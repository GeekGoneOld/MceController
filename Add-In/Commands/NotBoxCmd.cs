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
using System;
using System.Text.RegularExpressions;
using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;
using VmcController.AddIn.Metadata;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for NotBox command.
	/// </summary>
	public class NotBoxCmd : ICommand
	{
        private static Regex m_regex = new Regex("\"(?<message>.+?)\"\\s+(?<timeout>\\d+)\\s+\"(?<imagepath>.+?)\"");

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "\"message\" <timeout seconds> \"imagepath\"";
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
                if (match.Success)
                {
                    string imagefile = match.Groups["imagepath"].Value.Replace("/","\\");

                    // get latest image file from directory, notation example: @"c:\temp\test*.jpg"
                    if (imagefile.Contains("*"))
                    {
                        imagefile = GetFileInfo.GetNewestImage(imagefile);
                    }

                    DialogResultObject dialogObject = new DialogResultObject();
                    dialogObject.result = AddInHost.Current.MediaCenterEnvironment.DialogNotification(
                                            match.Groups["message"].Value,
                                            new System.Collections.ArrayList(1),
                                            int.Parse(match.Groups["timeout"].Value),
                                            "file://" + imagefile);
                    opResult.StatusCode = OpStatusCode.Ok;
                    opResult.ContentObject = dialogObject;
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
