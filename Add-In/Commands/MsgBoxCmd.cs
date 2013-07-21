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
 */
using System;
using System.Text.RegularExpressions;
using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;
using VmcController.AddIn.Metadata;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for MsgBox command.
	/// </summary>
	public class MsgBoxCmd : ICommand
	{
        private static Regex m_regex = new Regex("\"(?<caption>.+?)\"\\s+\"(?<message>.+?)\"\\s+(?<timeout>\\d+)");

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "\"caption\" \"message\" <timeout seconds>";
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
                    DialogResultObject result = new DialogResultObject();
                    result.result = AddInHost.Current.MediaCenterEnvironment.Dialog(
                                        match.Groups["message"].Value,
                                        match.Groups["caption"].Value,
                                        DialogButtons.Ok,
                                        int.Parse(match.Groups["timeout"].Value),
                                        false);
                    opResult.StatusCode = OpStatusCode.Success;
                    opResult.ContentObject = result;
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
