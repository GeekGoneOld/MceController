/*
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

using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;
//using Microsoft.Ehome.Epg;

using VmcController.AddIn.Commands;

namespace VmcController.AddIn
{
    /// <summary>
    /// Add-in for Vista Media Center for remote TCP control
    /// </summary>
    public sealed class AddInModule : IAddInModule, IAddInEntryPoint
    {
        #region Member Variables
        /// <summary>
        /// Used to signal that the plug-in should exit
        /// </summary>
        private EventWaitHandle m_waitHandle = new ManualResetEvent(false);

        /// <summary>
        /// The base TCP port number to use
        /// </summary>
        private int m_basePortNumber;

        /// <summary>
        /// TCP Socket Server
        /// </summary>
        private TcpSocketServer m_socketServer = new TcpSocketServer();

        /// <summary>
        /// HTTP Socket Server
        /// </summary>
        private HttpSocketServer m_httpServer = new HttpSocketServer();

        /// <summary>
        /// Available commands
        /// </summary>
        private RemoteCommands m_remoteCommands = new RemoteCommands();

        #endregion

        #region IAddInModule Members
        /// <summary>
        /// Enables an application to initialize its private variables and allocate system resources.
        /// </summary>
        /// <param name="appInfo">A collection of the attributes and corresponding values that were specified in the application element used to register the application</param>
        /// <param name="entryPointInfo">A collection of the attributes and corresponding values that were specified in the entrypoint element used to register the application's entry points</param>
        public void Initialize(Dictionary<string, object> appInfo, Dictionary<string, object> entryPointInfo)
        {
            //  Set the base port number
            int.TryParse(entryPointInfo["context"].ToString(), out m_basePortNumber);

            //  Sets the wait handle to the launch method will not exit
            m_waitHandle.Reset();

            //  Initializes the EPG guide data
            //Guide.Initialize();
        }

        /// <summary>
        /// Enables an application to save its state information and free system resources
        /// </summary>
        public void Uninitialize()
        {
            //  Allow our launch method to exit
            m_waitHandle.Set();

            //  Release the guide resources
            //Guide.Uninitialize();
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
        public void Launch(AddInHost host)
        {
            try
            {
                //  Lower the priority of this thread
                System.Threading.Thread.CurrentThread.Priority = ThreadPriority.Lowest;

                //  Setup TCP socket listener
                m_socketServer.StartListening(GetPortNumber(m_basePortNumber));
                m_socketServer.NewMessage += new EventHandler<SocketEventArgs>(m_socketServer_NewMessage);
                m_socketServer.Connected += new EventHandler<SocketEventArgs>(m_socketServer_Connected);
                m_socketServer.Disconnected += new EventHandler<SocketEventArgs>(m_socketServer_Disconnected);

                //  Setup HTTP socket listener
                m_httpServer.StartListening(GetPortNumber(m_basePortNumber) + 10);
                m_httpServer.NewRequest += new EventHandler<HttpEventArgs>(m_httpServer_NewRequest);

                //EventLog.WriteEntry("VmcController Client AddIn", "Listening on port " + m_socketServer.PortNumber + " (Version " + VersionInfo + ")", EventLogEntryType.Information);

                if (System.IO.File.Exists(System.Environment.GetEnvironmentVariable("windir") + "\\ehome\\vmcController.xml"))
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(System.Environment.GetEnvironmentVariable("windir") + "\\ehome\\vmcController.xml");
                    XmlNode startupCommand = doc.DocumentElement.SelectSingleNode("startupMacro");

                    if (startupCommand == null)
                    {
                        //AddInHost.Current.MediaCenterEnvironment.Dialog("startup node not found", "", DialogButtons.Ok, 5, false);
                    }
                    else
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
            catch (Exception ex)
            {
                //EventLog.WriteEntry("VmcController Client AddIn", "Exception in Launch: " + ex.ToString(), EventLogEntryType.Error);
            }
            finally
            {
                // Close RemotedWindowsMediaPlayer
                //new GetNowPlaying(true);

                //  Shutdown listener
                if (m_socketServer.PortNumber > 0)
                    m_socketServer.StopListening();
            }
        }

        /// <summary>
        /// Handles the Connected event of the m_socketServer control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="SocketServer.SocketEventArgs"/> instance containing the event data.</param>
        void m_socketServer_Connected(object sender, SocketEventArgs e)
        {
            m_socketServer.SendMessage(String.Format(
                "204 Connected (Clients: {0} Version: {1} Build Date: {2})\r\n",
                m_socketServer.Count, VersionInfo,
                RetrieveLinkerTimestamp().ToShortDateString()), e.TcpClient);
        }

        /// <summary>
        /// Handles the Disconnected event of the m_socketServer control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="SocketServer.SocketEventArgs"/> instance containing the event data.</param>
        void m_socketServer_Disconnected(object sender, SocketEventArgs e)
        {
            // Close RemotedWindowsMediaPlayer
            //new GetNowPlaying(true);
        }

        /// <summary>
        /// Handles the received commands of the m_socketServer control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="args">The <see cref="SocketServer.SocketEventArgs"/> instance containing the event data.</param>
        void m_socketServer_NewMessage(object sender, SocketEventArgs e)
        {
            OpResult opResult = new OpResult(OpStatusCode.BadRequest);
            try
            {
                if (e.Message.Length == 0)
                    return;

                if (e.Message.Equals("help", StringComparison.InvariantCultureIgnoreCase))
                {
                    opResult = m_remoteCommands.CommandList(GetPortNumber(m_basePortNumber));
                }
                else if (e.Message.Equals("exit", StringComparison.InvariantCultureIgnoreCase))
                {
                    m_socketServer.CloseClient(e.TcpClient);
                    return;
                }
                else
                {
                    string[] command = e.Message.Split(new char[] { ' ' }, 2);
                    if (command.Length == 0)
                        return;

                    opResult = m_remoteCommands.Execute(command[0], (command.Length == 2 ? command[1] : string.Empty));
                }
                m_socketServer.SendMessage(string.Format("{0} {1}\r\n", (int)opResult.StatusCode, opResult.StatusText), e.TcpClient);
                if (opResult.StatusCode == OpStatusCode.Ok)
                {
                    m_socketServer.SendMessage(opResult.ToString(), e.TcpClient);
                    m_socketServer.SendMessage(".\r\n", e.TcpClient);
                }

            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
            }
        }

        /// <summary>
        /// Handles the received commands of the m_httpServer control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="args">The <see cref="HttpServer.HttpEventArgs"/> instance containing the event data.</param>
        void m_httpServer_NewRequest(object sender, HttpEventArgs e)
        {
            Thread http_thread = new Thread(new ParameterizedThreadStart(m_httpServer_NewRequest_thread));
            http_thread.Start(e);
        }
        void m_httpServer_NewRequest_thread(Object o)
        {
            HttpEventArgs e = (HttpEventArgs)o;
            m_httpServer_NewRequest_thread(e);
        }
        void m_httpServer_NewRequest_thread(HttpEventArgs e)
        {
            OpResult opResult = new OpResult(OpStatusCode.BadRequest);
            string sCommand = "";
            string sParam = "";
            string sBody = "";
            string sTempBody = "";

            try
            {
                // Show error for index
                if (e.Request.Length == 0)
                {
                    sCommand = "<i>No command specified.</i>";
                    sParam = "<i>No parameters specified.</i>";
                }
                else
                {
                    string[] req = e.Request.Split(new char[] { '?' }, 2); //Strip off "?"
                    string[] cmd_stack = req[0].Split(new char[] { '/' });
                    for (int idx = 0; idx < cmd_stack.Length; idx++)
                    {
                        sTempBody = "";
                        string[] command = cmd_stack[idx].Split(new char[] { ' ' }, 2);
                        if (command.Length == 0)
                            return;
                        sCommand = command[0];
                        sParam = (command.Length == 2 ? command[1] : string.Empty);
                        if (sCommand.Equals("help", StringComparison.InvariantCultureIgnoreCase))
                            opResult = m_remoteCommands.CommandListHTML(GetPortNumber(m_basePortNumber));
                        else if (sCommand.Equals("format", StringComparison.InvariantCultureIgnoreCase))
                        {
                            ICommand formatter = new customCmd(sBody);
                            opResult = formatter.Execute(sParam);
                            sBody = "";
                        }
                        else opResult = m_remoteCommands.Execute(sCommand, sParam);
                        sTempBody = opResult.ToString();
                        if (sParam.Length == 0) sParam = "<i>No parameters specified.</i>";
                        if (opResult.StatusCode != OpStatusCode.Ok && opResult.StatusCode != OpStatusCode.Success)
                        {
                            sTempBody = string.Format("<h1>ERROR<hr>Command: {0}<br>Params: {1}<br>Returned: {2} - {3}<hr>See <a href='help'>Help</a></h1>", sCommand, sParam, opResult.StatusCode, opResult.ToString());
                            //if (sBody.Length > 0) sBody += "<HR>";
                            //sBody += sTempBody;
                            //break;
                        }
                        else if (opResult.StatusCode != OpStatusCode.OkImage)
                        {
                            if (sTempBody.Length > 0)
                            {
                                if (sTempBody.TrimStart()[0] != '<') sTempBody = "<pre>" + sTempBody + "</pre>";
                            }
                            else
                            {
                                sTempBody = string.Format("<h1>Ok<hr>Last Command: '{0}'<br>Params: {1}<br>Returned: {2}<hr>See <a href='help'>Help</a></h1>", sCommand, sParam, opResult.StatusCode);
                            }
                            //if (sBody.Length > 0) sBody += "<HR>";
                            //sBody += sTempBody;
                        }
                        if (sBody.Length > 0) sBody += "<HR>";
                        sBody += sTempBody;
                    }
                }
                if (opResult.StatusCode == OpStatusCode.OkImage) m_httpServer.SendImage(opResult.ToString(), opResult.StatusText, e.HttpSocket);
                else m_httpServer.SendPage(string.Format("{0}\r\n", sBody), e.HttpSocket);
            }
            catch (Exception ex)
            {
                m_httpServer.SendPage(string.Format("<html><body>EXCEPTION: {0}<hr></body></html>",
                        ex.Message), e.HttpSocket);
                Trace.TraceError(ex.ToString());
            }
        }

        #endregion

        #region Private Supporting Methods
        /// <summary>
        /// Gets the port number for the tcp server.
        /// </summary>
        /// <param name="basePort">The base port.</param>
        /// <returns></returns>
        private int GetPortNumber(int basePort)
        {
            string principalName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            int sessionId = Process.GetCurrentProcess().SessionId;
            Trace.TraceInformation("Windows Session #{0} Identity: {1}", sessionId, principalName);

            if (principalName.IndexOf("Mcx") > 0 && sessionId != 1)
                return basePort + int.Parse(principalName.Substring(principalName.LastIndexOf("Mcx") + 3, 1), CultureInfo.InvariantCulture);
            else if (sessionId == 1)
                return basePort;
            else
                throw new InvalidOperationException("Unable to determine correct port number");
        }

        /// <summary>
        /// Gets the assembly version info.
        /// </summary>
        /// <value>The version info.</value>
        public static string VersionInfo
        {
            get {
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
    }
}
