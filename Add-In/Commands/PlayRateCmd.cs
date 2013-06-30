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
using System.Text;
using WMPLib;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for PlayRate command.
	/// </summary>
	public class PlayRateCmd : WmpICommand
	{
        private bool m_set = true;

        public PlayRateCmd(bool bSet)
        {
            m_set = bSet;
        }

        #region WmpICommand Members

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
                    sb.AppendFormat("{0}|", modValue);
                }
                sb.Remove(sb.Length - 1, 1);
                s = "- returns the play state (one of " + sb.ToString() + ")";
            }
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
            throw new NotImplementedException();
        }

        public OpResult Execute(RemotedWindowsMediaPlayer remotePlayer, string param)
        {
            OpResult opResult = new OpResult();
            try
            {
                if (remotePlayer.getPlayState() == WMPPlayState.wmppsUndefined)
                {
                    if (m_set)
                    {
                        opResult.StatusCode = OpStatusCode.Ok;
                        opResult.AppendFormat("No media playing");
                    }
                    else
                    {
                        opResult.StatusCode = OpStatusCode.BadRequest;
                        opResult.AppendFormat("No media playing");
                    }
                }
                else if (m_set)
                {
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
                    }
                    opResult.StatusCode = OpStatusCode.Success;
                }
                else
                {
                    WMPPlayState state = remotePlayer.getPlayState();
                    string value = Enum.GetName(typeof(WMPPlayState), state).Remove(0, 5);
                    opResult.AppendFormat("PlayState={0}", value);
                    opResult.StatusCode = OpStatusCode.Success;
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
