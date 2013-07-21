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
using System.Collections.Generic;
using System.Text;
using VmcController.AddIn.Metadata;
using WMPLib;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for PlayRate command.
	/// </summary>
	public class PlayRateCmd : IWmpCommand
	{
        private bool m_set = true;

        public PlayRateCmd(bool bSet)
        {
            m_set = bSet;
        }

        #region IWmpCommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            String s;
            StringBuilder sb = new StringBuilder("");
            if (m_set)
            {               
                foreach (string value in Enum.GetNames(typeof(PlayRateEnum)))
                {
                    sb.AppendFormat("{0}|", value);
                }
                sb.Remove(sb.Length - 1, 1);
                s = "<" + sb.ToString() + "> - sets the play rate";
            }
            else
            {
                foreach (string value in Enum.GetNames(typeof(WMPPlayState)))
                {
                    string modValue = value.Remove(0, 5);
                    if (!modValue.Equals("Ready")) sb.AppendFormat("{0}|", modValue);
                }
                sb.Remove(sb.Length - 1, 1);
                s = "- returns the play state (one of " + sb.ToString() + ")";
            }
            return s;
        }

        public OpResult Execute(RemotedWindowsMediaPlayer remotePlayer, string param)
        {
            OpResult opResult = new OpResult();
            try
            {
                if (MediaExperienceWrapper.Instance == null || remotePlayer.getPlayState() == WMPPlayState.wmppsUndefined)
                {
                    opResult.StatusCode = OpStatusCode.BadRequest;
                    opResult.StatusText = "No media playing";
                }
                else if (m_set)
                {
                    if (param.Equals("")) throw new Exception("Not a supported playrate!");
                    PlayRateEnum playRate = (PlayRateEnum)Enum.Parse(typeof(PlayRateEnum), param, true);
                    switch (playRate)
                    {
                        case PlayRateEnum.Pause:
                            remotePlayer.getPlayerControls().pause();
                            break;
                        case PlayRateEnum.Play:
                            remotePlayer.getPlayerControls().play();
                            break;
                        case PlayRateEnum.Stop:
                            remotePlayer.getPlayerControls().stop();
                            break;
                        case PlayRateEnum.FR:
                            if (remotePlayer.getPlayerControls().get_isAvailable("FastReverse"))
                            {
                                remotePlayer.getPlayerControls().fastReverse();
                            }
                            else
                            {
                                throw new Exception("Not supported");
                            }
                            break;
                        case PlayRateEnum.FF:
                            if (remotePlayer.getPlayerControls().get_isAvailable("FastForward"))
                            {
                                remotePlayer.getPlayerControls().fastForward();
                            }
                            else
                            {
                                throw new Exception("Not supported");
                            }
                            break;
                        case PlayRateEnum.SkipBack:
                            remotePlayer.getPlayerControls().previous();
                            break;
                        case PlayRateEnum.SkipForward:
                            remotePlayer.getPlayerControls().next();
                            break;
                        default:
                            throw new Exception("Not a supported playrate!");
                    }
                    opResult.StatusCode = OpStatusCode.Success;
                }
                else
                {
                    WMPPlayState state = remotePlayer.getPlayState();
                    //string value = Enum.GetName(typeof(WMPPlayState), state).Remove(0, 5);
                    PlayStateObject pObject = new PlayStateObject();
                    pObject.play_state = CurrentState.getTruncatedPlayState(remotePlayer.getPlayState());
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

    enum PlayRateEnum
    {
        Stop,
        Pause,
        Play,
        FF,
        FR,
        SkipForward,
        SkipBack
    }
}
