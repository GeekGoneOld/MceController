/*
 * Copyright (c) 2007 Jonathan Bradshaw / 2012 Gert-Jan Niewenhuijse
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

/* This section Copyright (c) 2009 James Forrester
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
using System.Text.RegularExpressions;
using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for MsgBoxRich command.
	/// </summary>
    /// 
	public class MsgBoxRichCmd : ICommand
	{

        private DialogClosedCallback dlg;
        private DialogResult m_dlgResult;
        private bool responseReceived = false;
        
        private static Regex m_regex = new Regex("\"(?<caption>.+?)\"\\s+\"(?<message>.+?)\"\\s+(?<timeout>\\d+)\\s+\"(?<buttoncodes>.+?)\"\\s+\"(?<modal>.+?)\"\\s+\"(?<imagepath>.+?)\"");
        

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "\"caption\" \"message\" <timeout seconds> \"button codes\" \"modal|nonmodal\" \"imagepath\"";
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
                Match match = m_regex.Match(param);
                System.Collections.ArrayList buttonArray = new System.Collections.ArrayList();
                System.Collections.Hashtable buttonHT = new System.Collections.Hashtable();

                int customButtonID = 100;

                if (match.Groups["buttoncodes"].Value.Length > 0)
                {
                    string[] buttons = match.Groups["buttoncodes"].Value.Split(';');

                    foreach (string button in buttons)
                    {
                        switch (button)
                        {
                            case "OK":
                                buttonArray.Add(1);
                                break;
                            case "Cancel":
                                buttonArray.Add(2);
                                break;
                            case "Yes":
                                buttonArray.Add(4);
                                break;
                            case "No":
                                buttonArray.Add(8);
                                break;
                            default:
                                buttonArray.Add(button);
                                buttonHT.Add(customButtonID.ToString(),button);
                                customButtonID++;
                                break;
                        }
                    }
                }

                if (match.Success)
                {

                    responseReceived = false;
                    dlg = new DialogClosedCallback(On_DialogResult);

                    string imagefile = match.Groups["imagepath"].Value.Replace("/","\\");

                    // get latest image file from directory, notation example: @"c:\temp\test*.jpg"
                    if (imagefile.Contains("*"))
                    {
                        imagefile = GetFileInfo.GetNewestImage(imagefile);
                    }
                    
                    AddInHost.Current.MediaCenterEnvironment.Dialog(
                        match.Groups["message"].Value
                        ,match.Groups["caption"].Value
                        ,buttonArray
                        ,int.Parse(match.Groups["timeout"].Value)
                        ,match.Groups["modal"].Value == "modal" ? true:false
                        ,"file://" + imagefile
                        ,dlg);
                    
                    //wait for a response
                    while (!responseReceived)
                    {
                        System.Threading.Thread.Sleep(100);
                    }

                    string btnResult = "";

                    switch (m_dlgResult.ToString())
                    {
                        case "100":
                        case "101":
                        case "102":
                            btnResult = (string)buttonHT[m_dlgResult.ToString()];
                            break;
                        default:
                            btnResult = m_dlgResult.ToString();
                            break;

                    }

                    opResult.AppendFormat("response={0}", btnResult);

                    opResult.StatusCode = OpStatusCode.Success;

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

        private void On_DialogResult(DialogResult dlgResult)
        {
            m_dlgResult = dlgResult;
            responseReceived = true;
        }
    }
}
