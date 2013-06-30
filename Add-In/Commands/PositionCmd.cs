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
using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for Position command.
	/// </summary>
	public class PositionCmd: ICommand
	{
        private bool m_set = true;

        public PositionCmd(bool bSet)
        {
            m_set = bSet;
        }
        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            string s;
            if (m_set) s = "<seconds> - sets the current position";
            else s = "- returns the current position in seconds";
            return s;
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
            opResult.StatusCode = OpStatusCode.Success;
            try
            {
                if (AddInModule.getMediaExperience() == null)
                {
                    opResult.StatusCode = OpStatusCode.BadRequest;
                    opResult.AppendFormat("No media playing");
                }
                else if (m_set)
                {
                    TimeSpan position = TimeSpan.FromSeconds(double.Parse(param));
                    AddInHost.Current.MediaCenterEnvironment.MediaExperience.Transport.Position = position;
                }
                else
                {
                    TimeSpan position = AddInHost.Current.MediaCenterEnvironment.MediaExperience.Transport.Position;
                    opResult.AppendFormat("Position={0}", position);
                }
            }
            catch (Exception ex)
            {
                opResult.StatusCode = OpStatusCode.Exception;
                opResult.AppendFormat(ex.Message);
            }
            return opResult;
        }

        #endregion
    }
}
