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
using Newtonsoft.Json;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for Volume command.
	/// </summary>
	public class VolumeCmd : ICommand
	{
        public const int NO_VOLUME_STATE = -1;

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "<0-50|Up|Down|Mute|UnMute|Get> (cannot set volume on extender - volume fixed at 25)";
        }

        public int getVolume()
        {
            return (int)(AddInHost.Current.MediaCenterEnvironment.AudioMixer.Volume / 1310.7);
        }

        public bool isMuted()
        {
            return AddInHost.Current.MediaCenterEnvironment.AudioMixer.Mute;
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
                opResult.StatusCode = OpStatusCode.Success;
                VolumeState volumeState = new VolumeState();
                if (param.Equals("Up", StringComparison.InvariantCultureIgnoreCase))
                    //don't use this on extender
                    if (AddInModule.GetPortNumber(AddInModule.m_basePortNumber) != AddInModule.m_basePortNumber)
                    {
                        opResult.StatusCode = OpStatusCode.BadRequest;
                        opResult.StatusText = "Command not available on extenders.";
                    }
                    else
                    {
                        AddInHost.Current.MediaCenterEnvironment.AudioMixer.VolumeUp();
                    }
                else if (param.Equals("Down", StringComparison.InvariantCultureIgnoreCase))
                    //don't use this on extender
                    if (AddInModule.GetPortNumber(AddInModule.m_basePortNumber) != AddInModule.m_basePortNumber)
                    {
                        opResult.StatusCode = OpStatusCode.BadRequest;
                        opResult.StatusText = "Command not available on extenders.";
                    }
                    else
                    {
                        AddInHost.Current.MediaCenterEnvironment.AudioMixer.VolumeDown();
                    }
                else if (param.Equals("Mute", StringComparison.InvariantCultureIgnoreCase))
                    AddInHost.Current.MediaCenterEnvironment.AudioMixer.Mute = true;
                else if (param.Equals("UnMute", StringComparison.InvariantCultureIgnoreCase))
                    AddInHost.Current.MediaCenterEnvironment.AudioMixer.Mute = false;
                else if (!param.Equals("Get", StringComparison.InvariantCultureIgnoreCase))
                //don't use this on extender
                    if (AddInModule.GetPortNumber(AddInModule.m_basePortNumber) != AddInModule.m_basePortNumber)
                    {
                        opResult.StatusCode = OpStatusCode.BadRequest;
                        opResult.StatusText = "Command not available on extenders.";
                    }
                    else
                    {
                        int desiredLevel = int.Parse(param);
                        if (desiredLevel > 50 || desiredLevel < 0)
                        {
                            opResult.StatusCode = OpStatusCode.BadRequest;
                            opResult.StatusText = "Volume must be < 50 and > 0!";
                        }
                        int volume = (int)(AddInHost.Current.MediaCenterEnvironment.AudioMixer.Volume / 1310.7);
                        for (int level = volume; level > desiredLevel; level--)
                            AddInHost.Current.MediaCenterEnvironment.AudioMixer.VolumeDown();

                        for (int level = volume; level < desiredLevel; level++)
                            AddInHost.Current.MediaCenterEnvironment.AudioMixer.VolumeUp();
                    }
                volumeState.volume = getVolume();
                volumeState.is_muted = isMuted();
                opResult.ContentObject = volumeState;               
            }
            catch (Exception ex)
            {
                opResult.StatusCode = OpStatusCode.Exception;
                opResult.StatusText = ex.Message;
            }
            return opResult;
        }

        public class VolumeState : OpResultObject
        {
            public bool is_muted = false;
            public int volume = NO_VOLUME_STATE;
        }

        #endregion
    }
}
