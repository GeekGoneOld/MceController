/*
 * This module is loosly based on J.Bradshaw's CapabilitiesCmd.cs
 * 
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
 * History:
 * 2009-01-23 Created Anthony Jones
 * 
 */
using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;
using System.Reflection;

namespace VmcController.AddIn.Commands
{
    /// <summary>
    /// Summary description for FullScreen command.
    /// </summary>
    public class MediaMetaDataCmd : ICommand
    {
        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "- returns a key value pair list of current media experience";
        }

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <returns></returns>
        public OpResult Execute(string param)
        {
            var mce = AddInHost.Current.MediaCenterEnvironment.MediaExperience;

            // possible workaround for Win7/8 bug
            if (mce == null)
            {
                System.Threading.Thread.Sleep(200);
                mce = AddInHost.Current.MediaCenterEnvironment.MediaExperience;
                if (mce == null)
                {
                    try
                    {
                        var fi = AddInHost.Current.MediaCenterEnvironment.GetType().GetField("_checkedMediaExperience", BindingFlags.NonPublic | BindingFlags.Instance);
                        if (fi != null)
                        {
                            fi.SetValue(AddInHost.Current.MediaCenterEnvironment, false);
                            mce = AddInHost.Current.MediaCenterEnvironment.MediaExperience;
                        }

                    }
                    catch (Exception)
                    {
                        // give up 
                    }

                }
            }

            // Now try to read again
            OpResult opResult = new OpResult();
            try
            {
                if (AddInHost.Current.MediaCenterEnvironment.MediaExperience == null)
                {
                    opResult.StatusCode = OpStatusCode.BadRequest;
                    opResult.AppendFormat("No media playing");
                }
                else
                {
                    foreach (KeyValuePair<string, object> entry in AddInHost.Current.MediaCenterEnvironment.MediaExperience.MediaMetadata)
                    {
                        opResult.AppendFormat("{0}={1}", entry.Key, entry.Value);
                    }
                    opResult.StatusCode = OpStatusCode.Ok;
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
