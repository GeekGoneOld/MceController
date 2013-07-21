/*
 * Copyright (c) 2013 Skip Mercier
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
using System.Collections.Generic;
using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;
using WMPLib;
using System.Threading;
using System.Collections;
using Newtonsoft.Json;

namespace VmcController.AddIn.Commands {

	/// <summary>
	/// Summary description for PlayMediaCmd command.
	/// </summary>
	public class PlayMediaCmd : ICommand {

        private WindowsMediaPlayer Player;
        private IWMPPlaylist m_playlist;
        private ArrayList m_indexes;
		private bool m_appendToQueue;


        public PlayMediaCmd(RemotedWindowsMediaPlayer remotePlayer, IWMPPlaylist playlist, bool appendToQueue)
        {
            initPlayer(remotePlayer);
            m_indexes = null;
            m_playlist = playlist;
            m_appendToQueue = appendToQueue;
        }

        public PlayMediaCmd(RemotedWindowsMediaPlayer remotePlayer, IWMPPlaylist playlist, ArrayList indexes, bool appendToQueue)
        {
            initPlayer(remotePlayer);
            m_indexes = indexes;
            m_playlist = playlist;
			m_appendToQueue = appendToQueue;
		}

        private void initPlayer(RemotedWindowsMediaPlayer remotePlayer)
        {
            if (remotePlayer == null)
            {
                Player = null;
            }
            else
            {
                Player = remotePlayer.getPlayer();
            }
        }

		#region ICommand Members

		/// <summary>
		/// Shows the syntax.
		/// </summary>
		/// <returns></returns>
		public string ShowSyntax() {
			return "<" + "IWMPPlaylist parameters>";
		}

		/// <summary>
		/// Executes the specified param.
		/// </summary>
		/// <param name="param">The param.</param>
		/// <param name="result">The result.</param>
		/// <returns></returns>
		public OpResult Execute(string param) {
            OpResult opResult = new OpResult();
            try
            {
                if (setNowPlaying())
                {
                    opResult.StatusCode = OpStatusCode.Success;                    
                }
                else
                {
                    opResult.StatusCode = OpStatusCode.BadRequest;
                }
            }
            catch (Exception ex)
            {
                opResult.StatusCode = OpStatusCode.Exception;
                opResult.StatusText = ex.Message;
            }
			return opResult;
		}

        private void setMediaItem(IWMPMedia item, int index)
        {
            if (item != null)
            {

                if (Player != null)
                {
                    Player.currentPlaylist.appendItem(item);
                }
                else
                {
                    bool append;
                    if (index == 0)
                    {
                        append = m_appendToQueue;
                    }
                    else
                    {
                        append = true;
                    }
                    AddInHost.Current.MediaCenterEnvironment.PlayMedia(MediaType.Audio, item.sourceURL, append);
                }
            }
        }

        public bool setNowPlaying()
        {
            if (Player != null)
            {
                if (Player.currentPlaylist == null)
                {
                    Player.currentPlaylist = Player.newPlaylist(m_playlist.name, "");
                }
                if (!m_appendToQueue)
                {
                    Player.currentPlaylist.clear();
                } 
            }

            if (m_indexes != null)
            {
                for (int j = 0; j < m_indexes.Count; j++)
                {
                    setMediaItem(m_playlist.get_Item((int)m_indexes[j]), j);
                    System.Threading.Thread.Sleep(100);
                }
            }
            else
            {
                for (int j = 0; j < m_playlist.count; j++)
                {
                    setMediaItem(m_playlist.get_Item(j), j);
                    System.Threading.Thread.Sleep(100);    //Helps with stopped responding errors?
                }
            }
            return true;
        }

		#endregion
	}
}