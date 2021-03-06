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
using VmcController.AddIn.Commands;
using Microsoft.MediaCenter;
using WMPLib;
using System.Xml;
using Newtonsoft.Json;
using System.Threading;
using System.IO;


namespace VmcController.AddIn
{
    public interface IBaseCommand
    {
        string ShowSyntax();
    }

    public interface ICommand : IBaseCommand
    {
        OpResult Execute(string param);
    }

    public interface IExperienceCommand : IBaseCommand
    {
        OpResult ExecuteMediaExperience(string param);
    }

    public interface IWmpCommand : IBaseCommand
    {
        OpResult Execute(RemotedWindowsMediaPlayer remotePlayer, string param);
    }

    /// <summary>
    /// Manages the list of available remote commands
    /// </summary>
    public class RemoteCommands : IDisposable
    {
        private Dictionary<string, IBaseCommand> m_commands = new Dictionary<string, IBaseCommand>();
        private ReaderWriterLockSlim cacheLock = new ReaderWriterLockSlim();
        AutoResetEvent waitHandle = new AutoResetEvent(true);
        protected bool _disposed;

        /// <summary>
        /// Initializes a new instance of the <see cref="RemoteCommands"/> class.
        /// </summary>
        public RemoteCommands()
        {
            _disposed = false;

            m_commands.Add("=== Input Commands: ==========", null);
            m_commands.Add("button-rec", new SendKeyCmd('R', false, true, false));
            m_commands.Add("button-left", new SendKeyCmd(0x25));
            m_commands.Add("button-up", new SendKeyCmd(0x26));
            m_commands.Add("button-right", new SendKeyCmd(0x27));
            m_commands.Add("button-down", new SendKeyCmd(0x28));
            m_commands.Add("button-ok", new SendKeyCmd(0x0d));
            m_commands.Add("button-back", new SendKeyCmd(0x08));
            m_commands.Add("button-info", new SendKeyCmd('D', false, true, false));
            m_commands.Add("button-ch-plus", new SendKeyCmd(0xbb, false, true, false));
            m_commands.Add("button-ch-minus", new SendKeyCmd(0xbd, false, true, false));
            m_commands.Add("button-dvdmenu", new SendKeyCmd('M', true, true, false));
            m_commands.Add("button-dvdaudio", new SendKeyCmd('A', true, true, false));
            m_commands.Add("button-dvdsubtitle", new SendKeyCmd('U', true, true, false));
            m_commands.Add("button-cc", new SendKeyCmd('C', true, true, false));
            m_commands.Add("button-mute", new SendKeyCmd(0x77));
            m_commands.Add("button-space", new SendKeyCmd(0x20));
            m_commands.Add("button-pause", new SendKeyCmd('P', false, true, false));
            m_commands.Add("button-play", new SendKeyCmd('P', true, true, false));
            m_commands.Add("button-stop", new SendKeyCmd('S', true, true, false));
            m_commands.Add("button-skipback", new SkipCmd(false));
            m_commands.Add("button-skipfwd", new SkipCmd(true));
            m_commands.Add("button-rew", new SendKeyCmd('B', true, true, false));
            m_commands.Add("button-fwd", new SendKeyCmd('F', true, true, false));
            m_commands.Add("button-zoom", new SendKeyCmd('Z', true, true, false));
            m_commands.Add("button-num-0", new SendKeyCmd(0x60));
            m_commands.Add("button-num-1", new SendKeyCmd(0x61));
            m_commands.Add("button-num-2", new SendKeyCmd(0x62));
            m_commands.Add("button-num-3", new SendKeyCmd(0x63));
            m_commands.Add("button-num-4", new SendKeyCmd(0x64));
            m_commands.Add("button-num-5", new SendKeyCmd(0x65));
            m_commands.Add("button-num-6", new SendKeyCmd(0x66));
            m_commands.Add("button-num-7", new SendKeyCmd(0x67));
            m_commands.Add("button-num-8", new SendKeyCmd(0x68));
            m_commands.Add("button-num-9", new SendKeyCmd(0x69));
            m_commands.Add("button-num-star", new SendKeyCmd('3', true, false, false));
            m_commands.Add("button-num-number", new SendKeyCmd('8', true, false, false));
            m_commands.Add("button-clear", new SendKeyCmd(0x1b));
            m_commands.Add("type", new SendStringCmd());

            m_commands.Add("=== Misc Commands: ==========", null);
            m_commands.Add("dvdrom", new DvdRomCmd());
            m_commands.Add("msgbox", new MsgBoxCmd());
            m_commands.Add("msgboxrich", new MsgBoxRichCmd());
            m_commands.Add("notbox", new NotBoxCmd());
            m_commands.Add("notboxrich", new NotBoxRichCmd());
            m_commands.Add("goto", new NavigateToPage());
            m_commands.Add("announce", new AnnounceCmd());
            m_commands.Add("run-macro", new MacroCmd());
            m_commands.Add("suspend", new SuspendCmd());
            m_commands.Add("restartmc", new RestartMcCmd());
            m_commands.Add("server-settings", new ServerSettingsCmd());
            m_commands.Add("download-update", new DownloadUpdateCmd());

            m_commands.Add("=== Media Experience Commands: ==========", null);
            m_commands.Add("fullscreen", new FullScreenCmd());
            m_commands.Add("mediametadata", new MediaMetaDataCmd());
            m_commands.Add("playrate", new PlayRateCmd(true));
            m_commands.Add("playstate", new PlayRateCmd(false));
            m_commands.Add("position", new PositionCmd(true));
            m_commands.Add("position-get", new PositionCmd(false));
            m_commands.Add("media-info", new MediaInfoCmd());

            m_commands.Add("=== Environment Commands: ==========", null);
            m_commands.Add("mc-version", new VersionInfoCmd());
            m_commands.Add("addin-version", new VersionInfoPluginCmd());
            m_commands.Add("capabilities", new CapabilitiesCmd());
            m_commands.Add("changer-load", new ChangerCmd());

            m_commands.Add("=== Audio Mixer (Volume) Commands: ==========", null);
            m_commands.Add("volume", new VolumeCmd());

            m_commands.Add("=== Music Library Commands: ==========", null);
            m_commands.Add("music-list-artists", new MusicCmd(MusicCmd.LIST_ARTISTS));
            m_commands.Add("music-list-album-artists", new MusicCmd(MusicCmd.LIST_ALBUM_ARTISTS));
            m_commands.Add("music-list-albums", new MusicCmd(MusicCmd.LIST_ALBUMS));
            m_commands.Add("music-list-songs", new MusicCmd(MusicCmd.LIST_SONGS));
            m_commands.Add("music-list-playlists", new MusicCmd(MusicCmd.LIST_PLAYLISTS));
            m_commands.Add("music-list-details", new MusicCmd(MusicCmd.LIST_DETAILS));
            m_commands.Add("music-list-genres", new MusicCmd(MusicCmd.LIST_GENRES));
            m_commands.Add("music-list-recent", new MusicCmd(MusicCmd.LIST_RECENT));
            m_commands.Add("music-list-playing", new MusicCmd(MusicCmd.LIST_NOWPLAYING));
            m_commands.Add("music-list-current", new MusicCmd(MusicCmd.LIST_CURRENT));
            m_commands.Add("music-delete-playlist", new MusicCmd(MusicCmd.DELETE_PLAYLIST));
            m_commands.Add("music-play", new MusicCmd(MusicCmd.PLAY));
            m_commands.Add("music-queue", new MusicCmd(MusicCmd.QUEUE));
            m_commands.Add("music-shuffle", new MusicCmd(MusicCmd.SHUFFLE));
            m_commands.Add("music-cover", new MusicCmd(MusicCmd.SERV_COVER));
            m_commands.Add("music-clear-cache", new MusicCmd(MusicCmd.CLEAR_CACHE));
            m_commands.Add("music-list-stats", new MusicCmd(MusicCmd.LIST_STATS));

            m_commands.Add("=== Photo Library Commands: ==========", null);
            m_commands.Add("photo-clear-cache", new photoCmd(photoCmd.CLEAR_CACHE));
            m_commands.Add("photo-list", new photoCmd(photoCmd.LIST_PHOTOS));
            m_commands.Add("photo-play", new photoCmd(photoCmd.PLAY_PHOTOS));
            m_commands.Add("photo-queue", new photoCmd(photoCmd.QUEUE_PHOTOS));
            m_commands.Add("photo-tag-list", new photoCmd(photoCmd.LIST_TAGS));
            m_commands.Add("photo-serv", new photoCmd(photoCmd.SERV_PHOTO));
            m_commands.Add("photo-stats", new photoCmd(photoCmd.SHOW_STATS));

            m_commands.Add("=== Window State Commands: ==========", null);
            m_commands.Add("window-close", new SysCommand(SysCommand.SC_CLOSE));
            m_commands.Add("window-minimize", new SysCommand(SysCommand.SC_MINIMIZE));
            m_commands.Add("window-maximize", new SysCommand(SysCommand.SC_MAXIMIZE));
            m_commands.Add("window-restore", new SysCommand(SysCommand.SC_RESTORE));

            m_commands.Add("=== Reporting Commands: ==========", null);
            m_commands.Add("format", new customCmd());
        }

        /// <summary>
        /// Returns a multi-line list of commands and syntax
        /// </summary>
        /// <returns>string</returns>
        public OpResult CommandList()
        {
            return CommandList(0);
        }

        public OpResult CommandList(int port)
        {
            OpResult opResult = new OpResult(OpStatusCode.Ok);
            if (port != 0)
            {
                opResult.AppendFormat("=== Ports: ==========");
                opResult.AppendFormat("TCP/IP Socket port: {0}", port);
                opResult.AppendFormat("HTTP Server port: {0} (http://your_server:{1}/)", (port + 10), (port + 10));
            }
            opResult.AppendFormat("=== Connection Commands: ==========");
            opResult.AppendFormat("help - Shows this page");
            foreach (KeyValuePair<string, IBaseCommand> cmd in m_commands)
            {
                opResult.AppendFormat("{0} {1}", cmd.Key, (cmd.Value == null) ? "" : cmd.Value.ShowSyntax());
            }
            return opResult;
        }

        public OpResult CommandListHTML(int port)
        {
            OpResult opResult = new OpResult(OpStatusCode.Ok);

            string page_start =
                "<html><head><script LANGUAGE='JavaScript'>\r\n" +
                "function toggle (o){ var all=o.childNodes;  if (all[0].childNodes[0].innerText == '+') open(o,true);  else open(o,false);}\r\n" +
                "function open (o, b) { var all=o.childNodes;  if (b) {all[0].childNodes[0].innerText='-';all[1].style.display='inline';}  else {all[0].childNodes[0].innerText='+';all[1].style.display='none';}}\r\n" +
                "function toggleAll (b){ var all=document.childNodes[0].childNodes[1].childNodes; for (var i=0; i<all.length; i++) {if(all[i].id=='section') open(all[i],b)};}\r\n" +
                "</script></head>\r\n" +
                "<body><a id='top'>Jump to</a>: <a href='#commands'>Command List</a>, <a href='#examples'>Notes and Examples</a>, <a href='#bottom'>Bottom</a><hr><font size=+3><a id='commands'>Command List</a>:&nbsp;&nbsp;</font>[<a onclick='toggleAll(true);' >Open All</a>] | [<a onclick='toggleAll(false);' >Collapse All</a>]<hr>\r\n";
            string page_end =
                "</pre></span></div>\r\n" +
                "<br><hr><b><a id='examples'>Note: URLs must be correctly encoded</a></b><hr><br>\r\n" +
                "<b>Note - The following custom examples require that:</b><br>\r\n" +
                "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1 - Custom formats artist_browse and artist_list are defined in the &quot;music.template&quot; file<br>\r\n" +
                "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2 - The &quot;music.template&quot; file has been copied to the ehome directory (usually C:\\Windows\\ehome)<br>\r\n" +
                "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3 - Windows Media Center has been restarted after #1 and #2<br>\r\n" +
                "<br><b>Working track browser using custom formats: (can be slow... but this works as an album browser)</b><br>\r\n" +
                "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Display complete artist list linked to albums: <a href='music-list-details%20template:artist_list'>http://hostname:40510/music-list-details%20template:artist_list</a><br>\r\n" +
                "<br><b>Examples using artist filter: (warning can be very slow with large libraries)</b><br>\r\n" +
                "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;All artists: <a href='music-list-artists'>http://hostname:40510/music-list-artists</a><br>\r\n" +
                "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;All albums: <a href='music-list-albums'>http://hostname:40510/music-list-albums</a><br>\r\n" +
                "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;All albums by artists starting with the letter &quot;A&quot;: <a href='music-list-albums%20artist:a'>http://hostname:40510/music-list-albums%20artist:a</a><br>\r\n" +
                "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Play the tenth and thirteenth song in your collection: <a href='music-play%20indexes:10,13'>http://hostname:40510/music-play%20indexes:10,13</a><br>\r\n" +
                "<br><b>Examples using custom formats and artist match: (can be slow...)</b><br>\r\n" +
                "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Display pretty albums and tracks by the first artist starting with &quot;Jack&quot;: <a href='music-list-details%20template:artist_browse%20artist:jack'>http://hostname:40510/music-list-details%20template:artist_browse%20artist:jack</a><br>\r\n" +
                "<br><b>More help:</b><br>\r\n" +
                "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Help on the music commands: <a href='music-list-artists%20-help'>http://hostname:40510/music-list-artists%20-help</a><br>\r\n" +
                "<br><hr>\r\n" +
                "<a id='bottom'>Generated by</a>: Windows Media Center TCP/IP Controller (<a href='http://github.com/GeekGoneOld/MceController'>WMCController Home</a>)\r\n" +
                "<hr>Jump to: <a href='#top'>Top</a>, <a href='#commands'>Command List</a>, <a href='#examples'>Notes and Examples</a><br>\r\n" +
                "<script LANGUAGE='JavaScript'>toggleAll(false);</script></body></html>\r\n";

            string header_start = "</pre></span></div><br><div id='section' onclick='toggle(this)' style='border:solid 1px black;'><font size=+1 style='font:15pt courier;'><span>-</span>";
            string header_end = "</font><span style='display:'><pre>";

            opResult.AppendFormat("{0}", page_start);

            if (port != 0)
            {
                opResult.AppendFormat("{0} Ports: {1}", header_start, header_end);
                opResult.AppendFormat("HTTP command/server port: {0} (e.g http://your_server:{1}/", "4051x", 40510);
                opResult.AppendFormat("HTTP server port (deprecated): {0} (e.g http://your_server:{1}/", "4041x", 40410);
                opResult.AppendFormat("Streaming data server TCP port (deprecated): {0} (e.g. use PuTTY to connect to your_server:{1} raw)", "4040x", 40400);
            }
            opResult.AppendFormat("{0} Connection Commands: {1}", header_start, header_end);
            opResult.AppendFormat("<a href='/help'>help</a> - Shows this page");
            foreach (KeyValuePair<string, IBaseCommand> cmd in m_commands)
            {
                if (cmd.Key.StartsWith("==="))
                    opResult.AppendFormat(cmd.Key.Replace("==========", header_end).Replace("===", header_start));
                else
                    opResult.AppendFormat("<a href='/{0}'>{1}</a> {2}",
                        cmd.Key, cmd.Key, (cmd.Value == null) ? "" : cmd.Value.ShowSyntax().Replace("<", "&lt;").Replace(">", "&gt;"));
            }
            opResult.AppendFormat("{0}", page_end);
            return opResult;
        }

        private OpResult getSettings(ServerSettings settings, bool isBuilding)
        {
            settings.is_building = isBuilding;
            WindowsMediaPlayer Player = new WindowsMediaPlayer();
            settings.is_cache_outdated = MusicCmd.get_is_cache_outdated(MusicCmd.get_audio_item_count((IWMPMediaCollection2)Player.mediaCollection));
            ((ServerSettingsCmd)m_commands["server-settings"]).set(settings);
            return ((ICommand)m_commands["server-settings"]).Execute("");
        }

        public class LibraryStats : OpResultObject
        {
            public int albums = 0;
            public int album_artists = 0;
            public int artists = 0;
            public int genres = 0;
            public int songs = 0;
            public int playlists = 0;
        }

        public OpResult ExecuteLibraryStats(ServerSettings settings)
        {           
            if (cacheLock.TryEnterReadLock(0))
            {
                OpResult opResult = new OpResult();
                opResult.StatusCode = OpStatusCode.Success;
                LibraryStats library_stats = new LibraryStats();
                try {                    
                    int[] list_codes = new int[] {MusicCmd.LIST_ALBUMS, MusicCmd.LIST_ALBUM_ARTISTS, MusicCmd.LIST_ARTISTS, MusicCmd.LIST_GENRES, 
                        MusicCmd.LIST_DETAILS, MusicCmd.LIST_PLAYLISTS};                   
                    foreach (int i in list_codes)
                    {
                        MusicCmd cmd = new MusicCmd(i);
                        switch (i)
                        {
                            case MusicCmd.LIST_ALBUMS:
                                library_stats.albums = cmd.ExecuteStats().ResultCount;
                                break;
                            case MusicCmd.LIST_ALBUM_ARTISTS:
                                library_stats.album_artists = cmd.ExecuteStats().ResultCount;
                                break;
                            case MusicCmd.LIST_ARTISTS:
                                library_stats.artists = cmd.ExecuteStats().ResultCount;
                                break;
                            case MusicCmd.LIST_GENRES:
                                library_stats.genres = cmd.ExecuteStats().ResultCount;
                                break;
                            case MusicCmd.LIST_DETAILS:
                                library_stats.songs = cmd.ExecuteStats().ResultCount;
                                break;
                            case MusicCmd.LIST_PLAYLISTS:
                                library_stats.playlists = cmd.ExecuteStats().ResultCount;
                                break;
                        }
                    }                    
                }
                finally
                {
                    cacheLock.ExitReadLock();
                }
                opResult.ContentObject = library_stats;                                                
                return opResult;
            }
            else
            {
                return getSettings(settings, true);
            }
        }

        public OpResult ExecuteServerSettings(string command, ServerSettings settings)
        {
            if (cacheLock.TryEnterReadLock(0))
            {
                try { }
                finally
                {
                    cacheLock.ExitReadLock();
                }
                return getSettings(settings, false);
            }
            else
            {
                return getSettings(settings, true);
            }
        }

        public OpResult Execute(RemotedWindowsMediaPlayer remotePlayer, String command, string param)
        {
            return Execute(command, param, null, remotePlayer);
        }

        /// <summary>
        /// Executes a command with the given parameter string and returns a string return
        /// </summary>
        /// <param name="command">command name string</param>
        /// <param name="param">parameter string</param>
        /// <param name="playlist">now playing playlist, may be null</param>
        /// <param name="result">string</param>
        /// <returns></returns>
        public OpResult Execute(String command, string param, ServerSettings settings)
        {
            return Execute(command, param, settings, null);
        }

        public OpResult Execute(String command, string param, ServerSettings settings, RemotedWindowsMediaPlayer remotePlayer)
        {
            command = command.ToLower();
            if (m_commands.ContainsKey(command))
            {
                try
                {                    
                    if (command.Equals("music-list-stats"))
                    {
                        return ExecuteLibraryStats(settings);
                    }
                    else if (m_commands[command] is MusicCmd)
                    {
                        //Make sure cache is not being modified before executing any of the music-* commands
                        if (cacheLock.TryEnterReadLock(10))
                        {
                            try
                            {
                                return ((ICommand)m_commands[command]).Execute(param);
                            }
                            finally
                            {
                                cacheLock.ExitReadLock();
                            }
                        }
                        else
                        {
                            return getSettings(settings, true);
                        }
                    }
                    else if (m_commands[command] is IExperienceCommand || m_commands[command] is IWmpCommand)
                    {
                        //Only allow one thread at a time to access remoted player
                        waitHandle.WaitOne();
                        try
                        {
                            if (m_commands[command] is IWmpCommand) return ((IWmpCommand)m_commands[command]).Execute(remotePlayer, param);
                            else return ((IExperienceCommand)m_commands[command]).ExecuteMediaExperience(param);
                        }
                        finally
                        {                            
                            waitHandle.Set();
                        }
                    }
                    else
                    {
                        return ((ICommand)m_commands[command]).Execute(param);
                    }
                }
                catch (Exception ex)
                {
                    OpResult opResult = new OpResult();
                    opResult.StatusCode = OpStatusCode.Exception;
                    opResult.StatusText = ex.Message;
                    opResult.AppendFormat(ex.Message);
                    return opResult;
                }
            }
            else
            {
                return new OpResult(OpStatusCode.BadRequest);
            }
        }

        public void ExecuteCacheBuild(XmlDocument doc, Logger serverLogger)
        {
            cacheLock.EnterWriteLock();
            try
            {
                DateTime now = DateTime.Now;
                XmlNode lastCacheTimeNode = doc.DocumentElement.SelectSingleNode("lastCacheTime");
                if (lastCacheTimeNode != null)
                {
                    lastCacheTimeNode.InnerText = Convert.ToString(now);
                    //Save settings
                    doc.PreserveWhitespace = true;
                    doc.Save(AddInModule.DATA_DIR + "\\settings.xml");
                }

                serverLogger.Write("Updating cache now in progress");

                MusicCmd.clear_cache();

                Logger musicLogger = new Logger("MusicCmd", true);
                //Build caches for full library creation in Emote
                MusicCmd music = new MusicCmd(MusicCmd.LIST_DETAILS);
                music.ExecuteCacheBuild(musicLogger);
                music = new MusicCmd(MusicCmd.LIST_ARTISTS);
                music.ExecuteCacheBuild(musicLogger);
                music = new MusicCmd(MusicCmd.LIST_GENRES);
                music.ExecuteCacheBuild(musicLogger);
                musicLogger.Close();
                serverLogger.Write("Cache update finished");
            }
            finally
            {
                cacheLock.ExitWriteLock();
            }
        }

        public bool ExecuteCacheCheck()
        {
            bool isCacheVerified = false;
            cacheLock.EnterReadLock();
            try
            {
                WindowsMediaPlayer Player = new WindowsMediaPlayer();
                if (!MusicCmd.delete_old_cache(MusicCmd.get_audio_item_count((IWMPMediaCollection2)Player.mediaCollection)))
                {
                    //Only verify if cache wasn't deleted by check above
                    try
                    {
                        FileInfo fi = new FileInfo(MusicCmd.get_cache_filepath(MusicCmd.LIST_DETAILS, ""));
                        if (fi.Exists)
                        {
                            fi = new FileInfo(MusicCmd.get_cache_filepath(MusicCmd.LIST_ARTISTS, ""));
                            if (fi.Exists)
                            {
                                fi = new FileInfo(MusicCmd.get_cache_filepath(MusicCmd.LIST_GENRES, ""));
                                if (fi.Exists) isCacheVerified = true;
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                }                
            }
            finally
            {
                cacheLock.ExitReadLock();
            }
            return isCacheVerified;
        }

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                // Need to dispose managed resources if being called manually
                if (disposing)
                {
                    if (waitHandle != null) waitHandle.Close();
                    if (cacheLock != null) cacheLock.Dispose();
                    _disposed = true;
                }
            }
        }

        #endregion
    }
}
