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
using System.Collections.Generic;
using Microsoft.MediaCenter.Hosting;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for FullScreen command.
	/// </summary>
	public class DownloadUpdateCmd: ICommand
	{
        public const string DL_32_BIT_FILE_NAME = "WMCController32.msi";
        public const string DL_64_BIT_FILE_NAME = "WMCController64.msi";
        public const string DL_BASE_URL = "https://api.github.com/repos/GeekGoneOld/MceController/releases/latest";

        private bool is64bit = false;

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "- downloads the latest version of the plugin from GitHub to the Media Center computer user's desktop";
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
                //don't use this on extender
                if (AddInModule.GetPortNumber(AddInModule.m_basePortNumber) != AddInModule.m_basePortNumber)
                {
                    opResult.StatusCode = OpStatusCode.BadRequest;
                    opResult.StatusText = "Command not available on extenders.";
                }
                else
                {
                    //Determine 32- or 64-bit
                    string baseFileName;
                    if (AddInHost.Current.MediaCenterEnvironment.CpuClass.Contains("64"))
                    {
                        is64bit = true;
                        baseFileName = DL_64_BIT_FILE_NAME.Substring(0, DL_64_BIT_FILE_NAME.IndexOf(".msi"));
                    }
                    {
                        is64bit = false;
                        baseFileName = DL_32_BIT_FILE_NAME.Substring(0, DL_64_BIT_FILE_NAME.IndexOf(".msi"));
                    }

                    WebClient client = new WebClient();
                    client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                    JObject jObject = JObject.Parse(client.DownloadString(DL_BASE_URL));

                    // find right asset's download url
                    string downloadURL = "";
                    foreach (JToken jToken in jObject["assets"].Children())
                    {
                        Asset asset = JsonConvert.DeserializeObject<Asset>(jToken.ToString());
                        if (asset.name.ToLower() == (baseFileName.ToLower() + ".msi"))
                        {
                            downloadURL = asset.browser_download_url;
                            break;
                        }
                    }

                    if (downloadURL == "")
                    {
                        opResult.StatusCode = OpStatusCode.BadRequest;
                        opResult.StatusText = "Cannot find latest release on GitHub.";
                    }
                    else
                    {
                        string truncatedFileName = baseFileName;
                        int counter = 1;
                        while (File.Exists(getFilePath(truncatedFileName)))
                        {
                            truncatedFileName = getTruncatedFileName(counter);
                            counter++;
                        }
                        client.DownloadFile(downloadURL, getFilePath(truncatedFileName));

                        opResult.StatusCode = OpStatusCode.Success;
                        opResult.StatusText = "Downloaded " + truncatedFileName + ".msi to desktop.";
                    }
                }
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

public class Asset
{
    public string name { get; set; }
    public string browser_download_url { get; set; }
}