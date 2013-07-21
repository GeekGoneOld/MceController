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
using VmcController.AddIn.Metadata;
using System.Collections.Generic;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for Position command.
	/// </summary>
	public class PositionCmd: IExperienceCommand
	{
        private bool m_set = true;

        public PositionCmd(bool bSet)
        {
            m_set = bSet;
        }

        #region ExperienceICommand Members

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
        public OpResult ExecuteMediaExperience(string param)
        {
            OpResult opResult = new OpResult();
            opResult.StatusCode = OpStatusCode.Success;
            try
            {
                if (MediaExperienceWrapper.Instance == null)
                {
                    opResult.StatusCode = OpStatusCode.BadRequest;
                    opResult.StatusText = "No media playing";
                }
                else if (m_set)
                {
                    TimeSpan position = TimeSpan.FromSeconds(double.Parse(param));
                    MediaExperienceWrapper.Instance.Transport.Position = position;
                }
                else
                {
                    PositionObject pObject = new PositionObject();
                    pObject.position = MediaExperienceWrapper.Instance.Transport.Position;
                    opResult.StatusCode = OpStatusCode.Success;
                    opResult.ContentObject = pObject;
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
