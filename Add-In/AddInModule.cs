/*
 * Copyright (c) 2013 Skip Mercier
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 * This software is provided 'as-is', without any express or implied warranty. 
 * In no event will the authors be held liable for any damages arising from the use of this software.
 * Permission is granted to anyone to use this software for any purpose, including commercial 
 * applications, and to alter it and redistribute it freely, subject to the following restrictions:
 * 
 * 1. The origin of this software must not be misrepresented; you must not claim that you wrote 
 *    the original software. If you use this software in a product, an acknowledgment in the 
 *    product documentation would be appreciated but is not required.
 * 2. Altered source versions must be plainly marked as such, and must not be misrepresented as
 *    being the original software.
 * 3. This notice may not be removed or altered from any source distribution.
 * 
 * 
 * Modifications (c) 2009 Anthony Jones:
 * 06-07-10: Anthony Jones: Added threading to http requests to improve response
 * 2009: Added help context
 * 2013: Cleaned up code, misc changes
 * 
 */
using System;
using System.Collections;
using System.Globalization;
using System.Collections.Generic;
using System.Threading;
using System.Diagnostics;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO;

using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;

using VmcController.AddIn.Commands;
using VmcController;
using System.Windows.Threading;


namespace VmcController.AddIn
{
    /// <summary>
    /// Add-in for Vista Media Center for remote TCP control
    /// </summary>    
    public sealed class AddInModule : IAddInModule, IAddInEntryPoint, IDisposable
    {
        #region Member Variables

        public static string APP_NAME = "Media Center Network Controller";
        public static string DATA_DIR = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + APP_NAME;

        /// <summary>
        /// Used to signal that the plug-in should exit
        /// </summary>
        private EventWaitHandle m_waitHandle = new ManualResetEvent(false);

        /// <summary>
        /// The base TCP port number to use
        /// </summary>
        public static int m_basePortNumber;

        /// <summary>
        /// HTTP Socket Server
        /// </summary>
        private HttpSocketServer m_httpServer;

        private Logger logger;

        private bool needs_cleanup;

        private bool _disposed;

       
        #endregion

        #region IAddInModule Members
        /// <summary>
        /// Enables an application to initialize its private variables and allocate system resources.
        /// </summary>
        /// <param name="appInfo">A collection of the attributes and corresponding values that were specified in the application element used to register the application</param>
        /// <param name="entryPointInfo">A collection of the attributes and corresponding values that were specified in the entrypoint element used to register the application's entry points</param>
        public void Initialize(Dictionary<string, object> appInfo, Dictionary<string, object> entryPointInfo)
        {
            _disposed = false;

            needs_cleanup = true;

            if (!Directory.Exists(AddInModule.DATA_DIR))
            {
                Directory.CreateDirectory(AddInModule.DATA_DIR);
            }

            logger = new Logger("Add-In", false);

            //  Set the base port number
            int.TryParse(entryPointInfo["context"].ToString(), out m_basePortNumber);           

            m_httpServer = new HttpSocketServer();
            m_httpServer.InitServer();

            //  Sets the wait handle to the launch method will not exit
            m_waitHandle.Reset();
        }

        /// <summary>
        /// Enables an application to save its state information and free system resources
        /// </summary>
        public void Uninitialize()
        {
            cleanup();              
        }
        #endregion

        #region IAddInEntryPoint Members
        /// <summary>
        /// Starts running an application
        /// Because the Windows Media Center object is guaranteed to be valid only until the Launch
        /// method returns, an on-demand application must make all calls to the Windows Media Center
        /// API within the context of the application's Launch method. Calling the Windows Media
        /// Center object after the Launch method returns can result in a fatal error. If your
        /// application's Launch method spawns multiple threads, those threads must all be terminated
        /// before your Launch method returns. After calling the Launch method, if you do not call
        /// any method directly from the host object, .NET remoting releases unused objects every
        /// five minutes. To avoid this, use the host object or use the objects within five minutes
        /// to prevent them from being released.
        /// </summary>
        /// <param name="host">An application uses this interface to access other interfaces provided by the Microsoft.MediaCenter namespace</param>
        [STAThread]
        public void Launch(AddInHost host)
        {
            try
            {
                //  Lower the priority of this thread
                Thread.CurrentThread.Priority = ThreadPriority.BelowNormal;

                //Prevents occurrence of the program has stopped responding?
                System.Threading.Thread.Sleep(200);

                //MediaExperienceWrapper.Instance.GoToFullScreen();

                logger.Write("Starting http server...");

                //  Setup HTTP socket listener
                m_httpServer.StartListening(GetPortNumber(m_basePortNumber) + 10);

                if (System.IO.File.Exists(System.Environment.GetEnvironmentVariable("windir") + "\\ehome\\vmcController.xml"))
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(System.Environment.GetEnvironmentVariable("windir") + "\\ehome\\vmcController.xml");
                    XmlNode startupCommand = doc.DocumentElement.SelectSingleNode("startupMacro");

                    if (startupCommand != null)
                    {
                        MacroCmd macro = new MacroCmd();
                        OpResult result = macro.Execute(startupCommand.InnerText);
                        result = null;
                        macro = null;
                    }
                    doc = null;
                }

                //  Wait until exit request from host
                m_waitHandle.WaitOne();
            }
            catch (Exception e)
            {
                logger.Write("Exception in wait handle on Launch method: " + e.Message);
            }
            finally
            {
                cleanup();
            }
        }      

        #endregion

        #region Private Supporting Methods

        private void cleanup()
        {
            if (needs_cleanup)
            {
                logger.Write("Exiting...");

                //  Allow our launch method to exit
                m_waitHandle.Set();

                m_httpServer.CleanUpOnExit();

                needs_cleanup = false;
            }
        }

        /// <summary>
        /// Gets the port number for the tcp server.
        /// </summary>
        /// <param name="basePort">The base port.</param>
        /// <returns></returns>
        public static int GetPortNumber(int basePort)
        {
            string principalName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            int sessionId = Process.GetCurrentProcess().SessionId;
            Trace.TraceInformation("Windows Session #{0} Identity: {1}", sessionId, principalName);

            if (principalName.IndexOf("Mcx") > 0 && sessionId != 1)
                return basePort + int.Parse(principalName.Substring(principalName.LastIndexOf("Mcx") + 3, 1), CultureInfo.InvariantCulture);
            else if (sessionId == 1)
                return basePort;
            else
                return basePort;
               //Not sure why on my last two machines (Win8) sessionId was 2 so I'm commenting out -Skip Mercier
               //throw new InvalidOperationException("Unable to determine correct port number");
        }

        /// <summary>
        /// Gets the assembly version info.
        /// </summary>
        /// <value>The version info.</value>
        public static string VersionInfo
        {
            get
            {
                return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(3);
            }
        }    

        /// <summary>
        /// Retrieves the linker timestamp.
        /// </summary>
        /// <returns></returns>
        private DateTime RetrieveLinkerTimestamp()
        {
            string filePath = System.Reflection.Assembly.GetCallingAssembly().Location;
            const int c_PeHeaderOffset = 60;
            const int c_LinkerTimestampOffset = 8;
            byte[] b = new byte[2048];
            System.IO.Stream s = null;

            try
            {
                s = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                s.Read(b, 0, 2048);
            }
            finally
            {
                if (s != null) s.Close();
            }

            int i = System.BitConverter.ToInt32(b, c_PeHeaderOffset);
            int secondsSince1970 = System.BitConverter.ToInt32(b, i + c_LinkerTimestampOffset);
            DateTime dt = new DateTime(1970, 1, 1, 0, 0, 0);
            dt = dt.AddSeconds(secondsSince1970);
            dt = dt.AddHours(TimeZone.CurrentTimeZone.GetUtcOffset(dt).Hours);
            return dt;
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                // Need to dispose managed resources if being called manually
                if (disposing)
                {
                    if (m_waitHandle != null) m_waitHandle.Close();
                    if (m_httpServer != null) m_httpServer.Dispose();
                    if (logger != null) logger.Dispose();
                    _disposed = true;
                }                
            }
        }

        #endregion
    }
}
