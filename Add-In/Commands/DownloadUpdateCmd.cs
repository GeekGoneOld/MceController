/*
 * Copyright (c) 2013 Skip Mercier
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
 */
using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Net;
using Microsoft.MediaCenter.Hosting;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for FullScreen command.
	/// </summary>
	public class DownloadUpdateCmd: ICommand
	{
        public const string DL_32_BIT_FILE_NAME = "WMCController32.msi";
        public const string DL_64_BIT_FILE_NAME = "WMCController64.msi";
        public const string DL_BASE_URL = "https://github.com/GeekGoneOld/MceController/releases/latest/";

        private bool is64bit = false;

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "- not available yet but will download the latest version of the plugin to the user's desktop";
        }

        private string getFilePath(string truncatedFileName)
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + truncatedFileName + ".msi";
        }

        private string getTruncatedFileName(int counter)
        {
            string append = Convert.ToString(counter);
            if (is64bit) return DL_64_BIT_FILE_NAME.Substring(0, DL_64_BIT_FILE_NAME.IndexOf(".msi")) + " (" + append + ")";
            else return DL_32_BIT_FILE_NAME.Substring(0, DL_32_BIT_FILE_NAME.IndexOf(".msi")) + " (" + append + ")";
        }

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult Execute(string param)
        {            
            OpResult opResult = new OpResult();
            try
            {
/*                //Determine 32- or 64-bit
                if (AddInHost.Current.MediaCenterEnvironment.CpuClass.Contains("64")) is64bit = true;

                string truncatedFileName;
                if (is64bit) truncatedFileName = DL_64_BIT_FILE_NAME.Substring(0, DL_64_BIT_FILE_NAME.IndexOf(".msi"));
                else truncatedFileName = DL_32_BIT_FILE_NAME.Substring(0, DL_32_BIT_FILE_NAME.IndexOf(".msi"));

                int counter = 1;
                while (File.Exists(getFilePath(truncatedFileName)))
                {
                    truncatedFileName = getTruncatedFileName(counter);
                    counter++;
                }

                WebClient client = new WebClient();                
                client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");

                if (is64bit) client.DownloadFile(DL_BASE_URL + DL_64_BIT_FILE_NAME, getFilePath(truncatedFileName));
                else client.DownloadFile(DL_BASE_URL + DL_32_BIT_FILE_NAME, getFilePath(truncatedFileName));

*/                opResult.StatusCode = OpStatusCode.Success;
//                opResult.StatusText = "Downloaded " + truncatedFileName + " to desktop.";
                opResult.StatusText = "Didn't do anything, but wait 'til this works!";
            }
            catch (Exception ex)
            {
                opResult.StatusCode = OpStatusCode.Exception;
                opResult.StatusText = ex.Message;
            }
            return opResult;
        }

        #endregion
    }
}
