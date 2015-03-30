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
using System.Diagnostics;
using System.Security;
using System.Security.Permissions;
using System.Security.Principal;
using System.Windows.Forms;

namespace VmcController.AddIn.Commands
{
    class RestartMcCmd : ICommand
    {

        #region ICommand Members

        public string ShowSyntax()
        {
            return "Restarts Media Center (not available on extenders)";
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
            //don't use this on extender
            if (AddInModule.GetPortNumber(AddInModule.m_basePortNumber) != AddInModule.m_basePortNumber)
            {
                opResult.StatusCode = OpStatusCode.BadRequest;
                opResult.StatusText = "Command not available on extenders.";
            }
            else
                foreach (Process p in Process.GetProcesses())
                {
                    try
                    {
                        if (p.MainModule.ModuleName.Contains("ehshell"))
                            try
                            {
                                p.Kill();
                                Process.Start("ehshell.exe");
                                opResult.StatusCode = OpStatusCode.Success;
                            }
                            catch (Exception ex)
                            {
                                opResult.StatusCode = OpStatusCode.Exception;
                                opResult.StatusText = ex.Message;
                            }
                    }
                    catch (Exception)
                    {
                    }
                }
            return opResult;
        }

        #endregion
    }
}