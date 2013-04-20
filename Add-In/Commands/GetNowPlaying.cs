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
	/// Summary description for NotBoxRich command.
	/// </summary>
    /// 
	public class GetNowPlaying : ICommand
	{
        private static RemotedWindowsMediaPlayer player;
        

        public GetNowPlaying(bool should_close)
        {
            if (player == null && !should_close)
            {
                player = new RemotedWindowsMediaPlayer();
                player.CreateControl();
                player.Invoke();
            }
            else if (!should_close)
            {
                player.Invoke();
            }
            else 
            {
                player.close();
            }
        }

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "get the now playing list (nonfunctional - for testing only)";
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

                if (player != null)
                {
                    WMPLib.IWMPPlaylist mediaPlaylist = player.currentPlaylist;

                    opResult.StatusCode = OpStatusCode.Ok;
                    opResult.StatusText = "RemoteWindowsMediaPlayer is non null";
                }
                else
                {
                    opResult.StatusCode = OpStatusCode.BadRequest;
                    opResult.StatusText = "RemoteWindowsMediaPlayer is NULL!";
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
