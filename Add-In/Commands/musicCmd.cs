/*
 * This module is build on top of on J.Bradshaw's vmcController
 * Implements audio media library functions
 * 
 * Copyright (c) 2013 Skip Mercier
 * Copyright (c) 2009 Anthony Jones
 * 
 * Portions copyright (c) 2007 Jonathan Bradshaw
 * 
 * This software code module is provided 'as-is', without any express or implied warranty. 
 * In no event will the authors be held liable for any damages arising from the use of this software.
 * Permission is granted to anyone to use this software for any purpose, including commercial 
 * applications, and to alter it and redistribute it freely.
 * 
 * History:
 * Anthony Jones: 2010-06-07 Added music-list-album-artists command, added filter parameters to templates, added an "alpha" change template (A > B > C...), recently played list improvements
 * Anthony Jones: 2010-03-10 Added stats command
 * Anthony Jones: 2010-03-04 Added recently played commands, reworked some of the template code, many bug fixes
 * Anthony Jones: 2009-10-14 Added list by album artist
 * Anthony Jones: 2009-09-04 Reworked getArtistCmd to this, added caching
 * Skip Mercier: 2013 Lots of enhancements mainly concerning the use of data from the remoted WMP instance
 * 
 */

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections;
using System.Collections.Generic;
using System.Web;

using Microsoft.MediaCenter;
using VmcController;
using Microsoft.MediaCenter.UI;
using WMPLib;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


namespace VmcController.AddIn.Commands
{

    /// <summary>
    /// Summary description for getArtists commands.
    /// </summary>
    public class MusicCmd : MusicICommand
    {
        private RemotedWindowsMediaPlayer remotePlayer = null;
        private WindowsMediaPlayer Player = null;
        private Logger logger;

        private string debug_last_action = "none";

        public const int LIST_ARTISTS = 1;
        public const int LIST_ALBUMS = 2;
        public const int LIST_SONGS = 3;
        public const int LIST_DETAILS = 4;
        public const int PLAY = 5;
        public const int QUEUE = 6;
        public const int SERV_COVER = 7;
        public const int CLEAR_CACHE = 8;
        public const int LIST_ALBUM_ARTISTS = 9;
        public const int LIST_GENRES = 10;
        public const int LIST_RECENT = 11;
        public const int LIST_STATS = 12;
        public const int LIST_PLAYLISTS = 13;
        public const int LIST_NOWPLAYING = 14;
        public const int LIST_CURRENT = 15;
        public const int DELETE_PLAYLIST = 16;
        public const int SHUFFLE = 17;

        private int which_command = -1;

        private static Dictionary<string, string> m_templates = new Dictionary<string, string>();
        private state the_state = new state();
        private int result_count = 0;
        private string artist_filter = "";
        private string album_filter = "";
        private string genre_filter = "";
        private string song_filter = "";
        private string request_params = "";
        private bool m_cache_only = false;
        private bool m_stats_only = false;

        //private static string CACHE_DIR = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\VMC_Controller";
        private static string CACHE_MUSIC_CMD_DIR = AddInModule.DATA_DIR + "\\music_cmd_cache";
        private static string CACHE_VER_FILE = CACHE_MUSIC_CMD_DIR + "\\ver";

        private const string DEFAULT_DETAIL_ARTIST_START = "album_artist=%artist%";
        private const string DEFAULT_DETAIL_ALBUM_START = "     album=%album% (%albumYear%; %albumGenre%)";
        private const string DEFAULT_DETAIL_SONG = "          %if-songTrackNumber%track=%songTrackNumber%. %endif%song=%song% (%songLength%)";
        private const string DEFAULT_DETAIL_TRACK_ARTIST = "                   song_artist=%song_artist%";
        private const string DEFAULT_DETAIL_ALBUM_END = "          total album tracks=%albumTrackCount%";
        private const string DEFAULT_DETAIL_ARTIST_END = "     total artist tracks=%artistTrackCount%";
        private const string DEFAULT_DETAIL_FOOTER = "total artists found=%artistCount%\r\ntotal albums found=%albumCount%\r\ntotal tracks found=%trackCount%";
        private const string DEFAULT_STATS = "track_count=%track_count%\r\nartist_count=%artist_count%\r\nalbum_count=%album_count%\r\ngenre_count=%genre_count%\r\ncache_age=%cache_age%\r\navailable_templates=%available_templates%";

        private const string DEFAULT_IMAGE = "default.jpg";

        private static bool init_run = false;


        public MusicCmd(int i)
        {
            which_command = i;
            m_cache_only = false;
            init();
        }

        public MusicCmd(int i, bool cache_only)
        {
            which_command = i;
            m_cache_only = cache_only;
            logger = new Logger("MusicCmd", false);
            init();
        }

        public MusicCmd(RemotedWindowsMediaPlayer rPlayer, bool enqueue)
        {
            if (enqueue)
            {
                which_command = QUEUE;
            }
            else
            {
                which_command = PLAY;
            }
            remotePlayer = rPlayer;
            init();
        }

        public void setStatsOnly()
        {
            m_stats_only = true;
        }

        private void init()
        {
            if (!init_run)
            {
                init_run = true;
                loadTemplate();
            }
            Player = new WindowsMediaPlayer();
        }

        /// <summary>
        /// Checks to see if params contain an exact-hour param for setting cache time
        /// returns -1 if param not found
        /// </summary>
        public static int check_cache_command(string param)
        {
            if (param != null && param.Contains("exact-hour:"))
            {
                string album = param.Substring(param.IndexOf("exact-album:") + "exact-album:".Length);
                string hourString = trim_parameter(album);
                int hour = Convert.ToInt32(hourString);
                if (hour >= 0 && hour <= 12) return hour;
                else return -1;
            }
            return -1;
        }

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            string s = "[-help] [[exact-]artist:[*]artist_filter] [[exact-]album:[*]album_filter] [[exact-]genre:[*]genre_filter] [indexes:id1,id2] [template:template_name]- list / play from audio collection";
            switch (which_command)
            {
                case LIST_ARTISTS:
                    s = "[-help] [create-playlist:playlist_name] [[exact-]artist:[*]artist_filter] [[exact-]album:[*]album_filter] [[exact-]genre:[*]genre_filter] [template:template_name] - lists matching artists";
                    break;
                case LIST_ALBUM_ARTISTS:
                    s = "[-help] [create-playlist:playlist_name] [[exact-]artist:[*]artist_filter] [[exact-]album:[*]album_filter] [[exact-]genre:[*]genre_filter] [template:template_name] - lists matching album artists";
                    break;
                case LIST_ALBUMS:
                    s = "[-help] [create-playlist:playlist_name] [[exact-]artist:[*]artist_filter] [[exact-]album:[*]album_filter] [[exact-]genre:[*]genre_filter] [template:template_name] - list matching albums";
                    break;
                case LIST_SONGS:
                    s = "[-help] [create-playlist:playlist_name] [[exact-]artist:[*]artist_filter] [[exact-]album:[*]album_filter] [[exact-]genre:[*]genre_filter] [template:template_name] - list matching songs";
                    break;
                case LIST_GENRES:
                    s = "[-help] [create-playlist:playlist_name] [[exact-]artist:[*]artist_filter] [[exact-]album:[*]album_filter] [[exact-]genre:[*]genre_filter] [template:template_name] - list matching genres";
                    break;
                case LIST_PLAYLISTS:
                    s = "- list all playlists";
                    break;
                case LIST_DETAILS:
                    s = "[-help] [create-playlist:playlist_name] [exact-playlist:playlist_filter] [[exact-]artist:[*]artist_filter] [[exact-]album:[*]album_filter] [[exact-]genre:[*]genre_filter] [indexes:id1,id2] [template:template_name] - lists info on matching songs / albums / artists";
                    break;
                case PLAY:
                    s = "[-help] [exact-playlist:playlist_filter] [exact-song:song_filter] [[exact-]artist:[*]artist_filter] [[exact-]album:[*]album_filter] [[exact-]genre:[*]genre_filter] [indexes:id1,id2] - plays matching songs";
                    break;
                case QUEUE:
                    s = "[-help] [exact-playlist:playlist_filter] [exact-song:song_filter] [[exact-]artist:[*]artist_filter] [[exact-]album:[*]album_filter] [[exact-]genre:[*]genre_filter] [indexes:id1,id2] - adds matching songs to the now playing list";
                    break;
                case SHUFFLE:
                    s = "Changes play state to shuffle";
                    break;
                case SERV_COVER:
                    s = "[size-x:<width>] [size-y:<height>] [[exact-]artist:[*]artist_filter] [[exact-]album:[*]album_filter] [[exact-]genre:[*]genre_filter] [indexes:id1,id2] - serves the album cover of the first match";
                    break;
                case CLEAR_CACHE:
                    s = " - forces the cache to be cleared (normally only happens when the music library's length changes)";
                    break;
                case LIST_RECENT:
                    s = "[count:<how_many>] [template:template_name] - lists recently played / queued commands";
                    break;
                case LIST_STATS:
                    s = "[template:template_name] - lists stats (Artist / Album / Track counts, cache age, available templates, etc)";
                    break;
                case LIST_NOWPLAYING:
                    s = "- [index:id1] lists all songs in the current playlist or, if an index is supplied, playback will be set to that song in the list";
                    break;
                case LIST_CURRENT:
                    s = "- returns a key value pair list of current media similarly to mediametadata command";
                    break;
                case DELETE_PLAYLIST:
                    s = "[-help] [exact-playlist:playlist_filter] [indexes:id1,id2] deletes playlist specified by playlist_filter or, if indexes are supplied, only deletes items at indexes in the specified playlist";
                    break;
            }
            return s;
        }

        public OpResult showHelp(OpResult or)
        {
            or.AppendFormat("music-list-artists [~filters~] [~custom-template~] - lists all matching artists");
            or.AppendFormat("music-list-album-artists [~filters~] [~custom-template~] - lists all matching album artists - See WARNING");
            or.AppendFormat("music-list-songs [~filters~] [~custom-template~] - lists all matching songs");
            or.AppendFormat("music-list-albums [~filters~] [~custom-template~] - lists all matching albums");
            or.AppendFormat("music-list-genres [~filters~] [~custom-template~] - lists all matching genres");
            or.AppendFormat("music-list-playlists - lists all playlists");
            or.AppendFormat("music-list-playing [~index~] - lists songs in the current playlist or set playback to a specified song if index is supplied");
            or.AppendFormat("music-list-current - returns a key value pair list of current media similarly to mediametadata command");
            or.AppendFormat("music-delete-playlist [~filters~] [~index-list~] - deletes playlist specified by playlist_filter or, if indexes are supplied, only deletes items at indexes in the specified playlist");
            or.AppendFormat("music-play [~filters~] [~index-list~] - plays all matching songs");
            or.AppendFormat("music-queue [~filters~] [~index-list~] - queues all matching songs");
            or.AppendFormat("music-shuffle - sets playback mode to shuffle");
            or.AppendFormat("music-cover [~filters~] [~index-list~] [size-x:width] [size-y:height] - serves the cover image (first match)");
            or.AppendFormat(" ");
            or.AppendFormat("Where:");
            or.AppendFormat("     [~filters~] is one or more of: [~artist-filter~] [~album-filter~] [~genre-filter~] ");
            or.AppendFormat("     [~playlist-name~] is optional, can be an existing playlist to update, and must be combined with another filter below.");
            or.AppendFormat("     [~artist-filter~] is one of:");
            or.AppendFormat("          artist:<text> - matches track artists that start with <text> (\"artist:ab\" would match artists \"ABBA\" and \"ABC\")");
            or.AppendFormat("          artist:*<text> - matches track artists that have any words that start with <text> (\"artist:*ab\" would match \"ABBA\" and \"The Abstracts\")");
            or.AppendFormat("          exact-artist:<text> - matches the track artist that exactly matches <text> (\"exact-artist:ab\" would only match an artist names \"Ab\")");
            or.AppendFormat("     [~album-filter~] is one of:");
            or.AppendFormat("          album:<text> - matches albums that start with <text> (\"album:ab\" would match the album \"ABBA Gold\" and \"Abbey Road\")");
            or.AppendFormat("          exact-album:<text> - matches the album exactly named <text> (\"exact-album:ab\" would only match an album named \"Ab\")");
            or.AppendFormat("     [~genre-filter~] is one of:");
            or.AppendFormat("          genre:<text> - matches genre that start with <text> (\"genre:ja\" would match the genre \"Jazz\")");
            or.AppendFormat("          genre:*<text> - matches genres that contain <text> (\"genre:*rock\" would match \"Rock\" and \"Alternative Rock\")");
            or.AppendFormat("          exact-genre:<text> - matches the genere exactly named <text> (\"exact-genre:ja\" would only match an genre named \"Ja\")");
            or.AppendFormat("     [~playlist-filter~] is only:");
            or.AppendFormat("          exact-playlist:<text> - matches playlist exactly named <text>");
            or.AppendFormat("     [~song-filter~] is only:");
            or.AppendFormat("          exact-song:<text> - matches song exactly named <text>");
            or.AppendFormat("     [~index~] is of the form:");
            or.AppendFormat("          index:idx1 - specifies only one song in the current playlist by index");
            or.AppendFormat("     [~index-list~] is of the form:");
            or.AppendFormat("          indexes:idx1,idx2... - specifies one or more specific songs returned by the filter");
            or.AppendFormat("               Where idx1,idx2... is a comma separated list with no spaces (e.g. 'indexes:22,23,27')");
            or.AppendFormat("     [~custom-template~] is of the form:");
            or.AppendFormat("          template:<name> - specifies a custom template <name> defined in the \"music.template\" file");
            or.AppendFormat("     [size-x:~width~] - Resizes the served image, where ~width~ is the max width of the served image");
            or.AppendFormat("     [size-y:~height~] - Resizes the served image, where ~height~ is the max height of the served image");
            or.AppendFormat(" ");
            or.AppendFormat("Parameter Notes:");
            or.AppendFormat("     - Filter names containing two or more words must be enclosed in quotes.");
            or.AppendFormat("     - [~playlist-name~] must be combined with another filter and can be an existing playlist to update.");
            or.AppendFormat("     - [~song-filter~] can only be used with play and queue commands.");
            or.AppendFormat("     - Index numbers are just an index into the returned results and may change - they are not static!");
            or.AppendFormat("     - Both size-x and size-y must be > 0 or the original image will be returned without resizing.");
            or.AppendFormat(" ");
            or.AppendFormat(" ");
            or.AppendFormat("Examples:");
            or.AppendFormat("     music-list-artists - would return all artists in the music collection");
            or.AppendFormat("     music-list-album-artists - would return all album artists in the music collection");
            or.AppendFormat("          - WARNING: artists are filtered on the track level so this may be inconsistent");
            or.AppendFormat("     music-list-genres - would return all the genres in the music collection");
            or.AppendFormat("     music-list-artists artist:b - would return all artists in the music collection whose name starts with \"B\"");
            or.AppendFormat("     music-list-artists album:b - would return all artists in the music collection who have an album with a title that starts with \"B\"");
            or.AppendFormat("     music-list-albums artist:b - would return all albums by an artist whose name starts with \"B\"");
            or.AppendFormat("     music-list-albums artist:b album:*b - would return all albums that have a word starting with \"B\" by an artist whose name starts with \"B\"");
            or.AppendFormat("     music-list-albums genre:jazz - would return all the jazz albums");
            or.AppendFormat("     music-list-songs exact-artist:\"tom petty\" - would return all songs by \"Tom Petty\", but not songs by \"Tom Petty and the Heart Breakers \"");
            or.AppendFormat("     music-play exact-album:\"abbey road\" indexes:1,3 - would play the second and third songs (indexes are zero based) returned by the search for an album named \"Abbey Road\"");
            or.AppendFormat("     music-queue exact-artist:\"the who\" - would add all songs by \"The Who\" to the now playing list");
            return or;
        }

        private struct state
        {
            public string letter;

            public string genre;

            public string artist;
            public string nextArtist;

            public int artistTrackCount;
            public int artistAlbumCount;
            public string artistAlbumList;
            public int artistCount;

            public string album;
            public int albumTrackCount;
            public string albumYear;
            public string albumGenre;
            public string albumImage;
            public string nextAlbum;

            public string song;
            public string song_artist;
            public string songTrackNumber;
            public string songLength;
            public string songLocation;

            public int trackCount;
            public int albumCount;

            public void init()
            {
                letter = " ";
                artist = "";
                //artistIndex = 0;
                artistTrackCount = 0;
                artistAlbumCount = 0;
                artistAlbumList = "";
                artistCount = 0;
                nextArtist = "";

                album = "";
                //albumIndex = -1;
                albumTrackCount = 0;
                albumYear = "";
                albumGenre = "";
                albumImage = "";
                nextAlbum = "";

                song = "";
                song_artist = "";
                songTrackNumber = "";
                songLength = "";
                songLocation = "";

                trackCount = 0;
                albumCount = 0;
            }

            public void resetArtist(string new_artist)
            {
                artist = new_artist;
                nextArtist = "";

                artistTrackCount = 0;
                artistAlbumCount = 0;
                artistAlbumList = "";
                artistCount++;
            }

            public void resetAlbum(string new_album)
            {
                album = new_album;
                nextAlbum = "";

                if (new_album.Length > 0)
                {
                    albumList_add(new_album);
                    artistAlbumCount++;
                    albumCount++;
                }

                albumTrackCount = 0;
                albumYear = "";
                albumGenre = "";
                albumImage = "";

                //songTrackNumber = "";
                //songLength = "";

            }

            public void albumList_add(string s)
            {
                if (artistAlbumList.Length == 0) artistAlbumList = s;
                else artistAlbumList = artistAlbumList + ", " + s;
            }

            public void albumGenre_add(string s)
            {
                if (s == "") return;
                if (albumGenre.IndexOf(s) >= 0) return;
                if (albumGenre.Length == 0) albumGenre = s;
                else albumGenre = albumGenre + ", " + s;
            }

            public void findAlbumCover(string url)
            {
                string s = "";
                if (albumImage.Length > 0) return;

                try
                {
                    s = Path.GetDirectoryName(url) + @"\Folder.jpg";
                    if (File.Exists(s)) albumImage = s;
                    else
                    {
                        s = Path.GetDirectoryName(url) + @"\AlbumArtSmall.jpg";
                        if (File.Exists(s)) albumImage = s;
                    }
                }
                finally { }
            }

            public void add_song_to_album(IWMPMedia media_item)
            {
                if (albumYear == "" || albumYear.Length < 4) albumYear = media_item.getItemInfo("WM/OriginalReleaseYear");
                if (albumYear == "" || albumYear.Length < 4) albumYear = media_item.getItemInfo("WM/Year");
                genre = media_item.getItemInfo("WM/Genre");
                albumGenre_add(genre);

                if (albumImage.Length == 0) findAlbumCover(media_item.sourceURL);
                albumTrackCount++;
                artistTrackCount++;
                trackCount++;

                song = media_item.getItemInfo("Title");
                songLocation = media_item.sourceURL;
                songTrackNumber = media_item.getItemInfo("WM/TrackNumber");
                songLength = media_item.durationString;
                song_artist = media_item.getItemInfo("Author");

                string s = media_item.getItemInfo("WM/AlbumTitle");
                if (album != s) nextAlbum = s;

                s = media_item.getItemInfo("WM/AlbumArtist");
                if (s == "") s = "None";
                if (artist != s) nextArtist = s;

                return;
            }
        }

        public string do_conditional_replace(string s, string item, string v)
        {
            debug_last_action = "Conditional replace: Start - item: " + item;

            string value = "";
            try { value = v; }
            catch (Exception) { value = ""; }

            if (value == null) value = "";
            else value = value.Trim();

            int idx_start = -1;
            int idx_end = -1;
            debug_last_action = "Conditional replace: Checking Conditional - item: " + item;
            while ((idx_start = s.IndexOf("%if-" + item + "%")) >= 0)
            {
                if (value.Length == 0)
                {
                    if ((idx_end = s.IndexOf("%endif%", idx_start)) >= 0)
                        s = s.Substring(0, idx_start) + s.Substring(idx_end + 7);
                    else s = s.Substring(0, idx_start);
                }
                else
                {
                    if ((idx_end = s.IndexOf("%endif%", idx_start)) >= 0)
                        s = s.Substring(0, idx_end) + s.Substring(idx_end + 7);
                    s = s.Substring(0, idx_start) + s.Substring(idx_start + ("%if-" + item + "%").Length);
                }
            }
            debug_last_action = "Conditional replace: Doing replace - item: " + item;
            s = s.Replace("%" + item + "%", value);

            debug_last_action = "Conditional replace: End - item: " + item;

            return s;
        }

        public string file_includer(string s)
        {
            int idx_start = -1;
            int idx_end = -1;
            string fn = null;
            while ((idx_start = s.IndexOf("%file-include%")) >= 0)
            {
                if ((idx_end = s.IndexOf("%endfile%", idx_start)) >= 0)
                    fn = s.Substring((idx_start + ("%file-include%".Length)), (idx_end - (idx_start + ("%file-include%".Length))));
                else fn = s.Substring(idx_start + "%file-include%".Length);
                fn = fix_escapes(fn);

                string file_content = null;
                FileInfo fi = new FileInfo(fn);
                if (!fi.Exists) file_content = "";
                else
                {
                    StreamReader include_stream = File.OpenText(fn);
                    file_content = include_stream.ReadToEnd();
                    include_stream.Close();
                }
                s = s.Substring(0, idx_start) + file_content + s.Substring(idx_end + "%endfile%".Length);
            }
            return s;
        }

        private string basic_replacer(string s, string item, string value, int count, int index)
        {
            if (s.Length > 0) s = do_conditional_replace(s, item, value);
            s = do_conditional_replace(s, "resultCount", String.Format("{0}", count));
            if (index >= 0) s = do_conditional_replace(s, "index", String.Format("{0}", index));

            return s;
        }

        private string first_letter(string s)
        {
            if (s == null || s.Length == 0) return " ";

            string ret = "";
            if (s.StartsWith("the ", StringComparison.CurrentCultureIgnoreCase)) ret = s.Substring(4, 1);
            else ret = s.Substring(0, 1);

            ret = ret.ToUpper();
            if ("ABCDEFGHIJKLMNOPQRSTUVWXYZ".IndexOf(ret) < 0) ret = "#";

            return ret;
        }

        private string replacer(string s, int index)
        {
            /* The URL parameters*/
            s = do_conditional_replace(s, "all_filters", request_params);

            /* Genre */
            s = do_conditional_replace(s, "genre", the_state.genre);
            s = do_conditional_replace(s, "genreFilter", genre_filter);
            /* Artist */
            s = do_conditional_replace(s, "artist", the_state.artist);
            s = do_conditional_replace(s, "artistFilter", artist_filter);
            s = do_conditional_replace(s, "letter", first_letter(the_state.artist));
            s = do_conditional_replace(s, "artistTrackCount", String.Format("{0}", the_state.artistTrackCount));
            s = do_conditional_replace(s, "artistAlbumCount", String.Format("{0}", the_state.artistAlbumCount));
            s = do_conditional_replace(s, "artistCount", String.Format("{0}", the_state.artistCount));
            s = do_conditional_replace(s, "nextArtist", the_state.nextArtist);
            s = do_conditional_replace(s, "artistAlbumList", the_state.artistAlbumList);

            /* Album*/
            s = do_conditional_replace(s, "album", the_state.album);
            s = do_conditional_replace(s, "albumFilter", album_filter);
            //    s = s.Replace("%nextAlbum%", nextAlbum);
            s = do_conditional_replace(s, "albumTrackCount", String.Format("{0}", the_state.albumTrackCount));
            s = do_conditional_replace(s, "albumYear", the_state.albumYear);
            s = do_conditional_replace(s, "albumGenre", the_state.albumGenre);
            s = do_conditional_replace(s, "albumImage", the_state.albumImage);
            s = do_conditional_replace(s, "albumYear", the_state.albumYear);
            s = do_conditional_replace(s, "albumGenre", the_state.albumGenre);
            s = do_conditional_replace(s, "nextAlbum", the_state.nextAlbum);

            /* Song */
            s = do_conditional_replace(s, "song", the_state.song);
            s = do_conditional_replace(s, "song_artist", the_state.song_artist);
            s = do_conditional_replace(s, "songPath", the_state.songLocation);
            s = do_conditional_replace(s, "songTrackNumber", the_state.songTrackNumber);
            s = do_conditional_replace(s, "songLength", the_state.songLength);
            s = s.Replace("%index%", String.Format("{0}", index));

            s = do_conditional_replace(s, "trackCount", String.Format("{0}", the_state.trackCount));
            s = do_conditional_replace(s, "albumCount", String.Format("{0}", the_state.albumCount));

            return s;
        }

        public static string trim_parameter(string param)
        {
            if (param.Substring(0, 1) == "\"")
            {
                param = param.Substring(1);
                if (param.IndexOf("\"") >= 0) param = param.Substring(0, param.IndexOf("\""));
            }
            else if (param.IndexOf(" ") >= 0) param = param.Substring(0, param.IndexOf(" "));

            return param;
        }

        private bool loadTemplate()
        {
            bool ret = true;

            try
            {
                Regex re = new Regex("(?<lable>.+?)\t+(?<format>.*$?)");
                StreamReader fTemplate = File.OpenText("music.template");
                string sIn = null;
                while ((sIn = fTemplate.ReadLine()) != null)
                {
                    Match match = re.Match(sIn);
                    if (match.Success) m_templates.Add(match.Groups["lable"].Value, match.Groups["format"].Value);
                }
                fTemplate.Close();
            }
            catch { ret = false; }

            return ret;
        }

        private string fix_escapes(string s)
        {
            s = s.Replace("\r", "\\r");
            s = s.Replace("\n", "\\n");
            s = s.Replace("\t", "\\t");

            return s;
        }

        private string getTemplate(string template, string default_template)
        {
            string tmp = "";

            if (!m_templates.ContainsKey(template)) return default_template;

            tmp = m_templates[template];
            tmp = tmp.Replace("\\r", "\r");
            tmp = tmp.Replace("\\n", "\n");
            tmp = tmp.Replace("\\t", "\t");

            tmp = file_includer(tmp);

            return tmp;
        }

        public string make_cache_fn(string fn)
        {
            fn = fn.Replace("\\", "_");
            fn = fn.Replace(":", "_");
            fn = fn.Replace(" ", "%20");
            fn = fn.Replace("\"", "_");

            return fn;
        }

        public bool check_cache(string cur_ver)
        {
            bool ret_val = true;
            string cache_ver = "";

            try
            {
                cache_ver = System.IO.File.ReadAllText(CACHE_VER_FILE);
                if (cur_ver != cache_ver) ret_val = false;
            }
            catch (Exception) { ret_val = false; }

            if (!ret_val) clear_cache();

            return ret_val;
        }

        public void clear_cache()
        {
            try { System.IO.Directory.Delete(CACHE_MUSIC_CMD_DIR, true); }
            catch (Exception) { return; }
        }

        public void save_to_cache(string fn, string content, string cur_ver)
        {
            check_cache(cur_ver);

            string cached_file = CACHE_MUSIC_CMD_DIR + "\\" + fn;

            try
            {
                //Create dir if needed:
                if (!Directory.Exists(AddInModule.DATA_DIR)) Directory.CreateDirectory(AddInModule.DATA_DIR);
                if (!Directory.Exists(CACHE_MUSIC_CMD_DIR)) Directory.CreateDirectory(CACHE_MUSIC_CMD_DIR);

                FileInfo fi = new FileInfo(CACHE_VER_FILE);
                if (!fi.Exists)
                {
                    System.IO.File.WriteAllText(CACHE_VER_FILE, cur_ver);
                }
                System.IO.File.WriteAllText(cached_file, content);
            }
            catch (Exception)
            {
            }
        }

        public string get_cached(string fn, string cur_ver)
        {
            string cached = "";

            if (!check_cache(cur_ver)) return "";

            try
            {
                string cached_file = CACHE_MUSIC_CMD_DIR + "\\" + fn;

                FileInfo fi = new FileInfo(cached_file);
                if (fi.Exists)
                {
                    cached = System.IO.File.ReadAllText(cached_file);
                }
            }
            catch (Exception) { return ""; }

            return cached;
        }

        public ArrayList get_mrp(bool for_display)
        {
            ArrayList mrp_content = new ArrayList();
            string[] line_array;

            string mrp_file = AddInModule.DATA_DIR + "\\mrp_list.dat";

            try
            {
                FileInfo fi = new FileInfo(mrp_file);
                if (fi.Exists)
                {
                    System.IO.StreamReader file = new System.IO.StreamReader(mrp_file);
                    string line;
                    while ((line = file.ReadLine()) != null)
                    {
                        if (for_display)
                        {
                            mrp_content.Add(line);
                        }
                        else // When getting an array to look for existing entry ignore the count
                        {
                            line_array = line.Split('\t');
                            mrp_content.Add(line_array[0] + "\t" + line_array[1] + "\t" + line_array[2]);
                        }
                    }
                    file.Close();
                }
            }
            catch (Exception) { ; }

            return mrp_content;
        }

        public void add_to_mrp(string recent_text_type, string recent_text, string param, int track_count)
        {
            string mrp_file = AddInModule.DATA_DIR + "\\mrp_list.dat";
            string cmd = recent_text_type + "\t" + recent_text + "\t" + HttpUtility.UrlEncode(param);

            // read in existing list:
            ArrayList mrp_content_exists = get_mrp(false); // mrp without track count (which can change)
            ArrayList mrp_content_final = get_mrp(true);   // complete mrp for output

            // See of entry already exists (i.e. already been played)
            if (mrp_content_exists.Contains(cmd))
            {
                int idx = mrp_content_exists.IndexOf(cmd);
                mrp_content_final.RemoveAt(idx);
            }
            // Insert cmd at begining
            mrp_content_final.Insert(0, cmd + "\t" + track_count);
            // Trim to last 500 plays (arbitrary)
            if (mrp_content_exists.Count > 500) mrp_content_exists.RemoveRange(500, (mrp_content_exists.Count - 500));

            try
            {
                //Create dir if needed:
                if (!Directory.Exists(AddInModule.DATA_DIR)) Directory.CreateDirectory(AddInModule.DATA_DIR);

                System.IO.StreamWriter file = new System.IO.StreamWriter(mrp_file);
                foreach (string s in mrp_content_final) file.WriteLine(s);
                file.Close();
            }
            catch (Exception) { return; }

            return;
        }

        private OpResult list_recent(OpResult or, string template)
        {
            return list_recent(or, template, -1);
        }

        private OpResult list_recent(OpResult or, string template, int count)
        {
            ArrayList mrp_content = get_mrp(true);
            int result_count = mrp_content.Count;
            if (count > 0) result_count = (count < result_count) ? count : result_count;
            or.AppendFormat("{0}", basic_replacer(getTemplate(template + ".H", ""), "", "", result_count, -1)); // Header
            for (int j = 0; j < result_count; j++)
            {
                string list_item = (string)mrp_content[j];
                if (list_item.Length > 0)
                {
                    string s_out = "";
                    string[] values = list_item.Split('\t');
                    s_out = getTemplate(template + ".Entry", "%index%. %type%: %description% (%param%)");
                    s_out = basic_replacer(s_out, "full_type", values[0], result_count, j);
                    s_out = basic_replacer(s_out, "description", values[1], result_count, j);
                    s_out = basic_replacer(s_out, "param", values[2], result_count, j);
                    s_out = basic_replacer(s_out, "trackCount", values[3], result_count, j);
                    if (s_out.IndexOf("%title%") > 0)
                    {
                        string title = values[1];
                        if (title.IndexOf(":") > 0) title = title.Substring(title.LastIndexOf(":") + 1);
                        s_out = basic_replacer(s_out, "title", title, result_count, j);
                    }
                    if (s_out.IndexOf("%type%") > 0)
                    {
                        string type = values[0];
                        if (type.IndexOf("/") > 0) type = type.Substring(type.LastIndexOf("/") + 1);
                        s_out = basic_replacer(s_out, "type", type, result_count, j);
                    }

                    if (s_out.Length > 0) or.AppendFormat("{0}", s_out); // Entry
                }
            }
            or.AppendFormat("{0}", basic_replacer(getTemplate(template + ".F", "result_count=%resultCount%"), "", "", result_count, -1));

            return or;
        }

        private OpResult add_to_playlist(OpResult or, bool contains_query, IWMPPlaylist queried, string playlist_name)
        {
            if (contains_query)
            {
                IWMPPlaylistCollection playlistCollection = (IWMPPlaylistCollection)Player.playlistCollection;
                IWMPPlaylistArray current_playlists = playlistCollection.getByName(playlist_name);
                if (current_playlists.count > 0)
                {
                    IWMPPlaylist previous_playlist = current_playlists.Item(0);
                    for (int j = 0; j < queried.count; j++)
                    {
                        previous_playlist.appendItem(queried.get_Item(j));
                    }
                    or.ContentText = "Playlist " + playlist_name + " updated.";
                }
                else
                {
                    queried.name = playlist_name;
                    playlistCollection.importPlaylist(queried);
                    or.ContentText = "Playlist " + playlist_name + " added.";
                }
                or.StatusCode = OpStatusCode.Success;
            }
            else
            {
                or.StatusCode = OpStatusCode.BadRequest;
                or.StatusText = "Playlists can only be created using a query for a specific items.";
            }
            return or;
        }

        private IWMPPlaylistArray getAllUserPlaylists(IWMPPlaylistCollection collection)
        {
            return getUserPlaylistsByName(null, collection);
        }

        private IWMPPlaylistArray getUserPlaylistsByName(string query, IWMPPlaylistCollection collection)
        {
            if (query != null)
            {
                return collection.getByName(query);
            }
            else
            {
                return collection.getAll();
            }
        }        

        private string getSortAttributeFromQueryType(string query_type)
        {
            if (query_type.Equals("Album"))
            {
                return "WM/AlbumTitle";
            }
            else if (query_type.Equals("Album Artist"))
            {
                return "WM/AlbumArtist";
            }
            else if (query_type.Equals("Artist"))
            {
                return "Author";
            }
            else if (query_type.Equals("Genre"))
            {
                return "WM/Genre";
            }
            else if (query_type.Equals("Song"))
            {
                return "Title";
            }
            return "";
        }

        private IWMPPlaylist getPlaylistFromExactQuery(string query_text, string query_type, IWMPMediaCollection2 collection)
        {
            if (query_type.Equals("Album"))
            {
                return collection.getByAlbum(query_text);
            }
            else if (query_type.Equals("Album Artist"))
            {
                return collection.getByAttribute("WM/AlbumArtist", query_text);
            }
            else if (query_type.Equals("Artist"))
            {
                return collection.getByAuthor(query_text);
            }
            else if (query_type.Equals("Genre"))
            {
                return collection.getByGenre(query_text);
            }
            else if (query_type.Equals("Song"))
            {
                return collection.getByName(query_text);
                //mediaPlaylist = collection.getByAttribute("Title", query_text);
            }
            return null;
        }

        public string findAlbumPath(string url)
        {
            string path = "";
            try
            {
                path = Path.GetDirectoryName(url) + @"\Folder.jpg";
                if (File.Exists(path)) return path;
                else
                {
                    path = Path.GetDirectoryName(url) + @"\AlbumArtSmall.jpg";
                    if (File.Exists(path)) return path;
                }
            }
            finally { }
            return path;
        }

        public OpResult Execute(string param)
        {
            return Execute(param, null, null);
        }

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult Execute(string param, NowPlayingList nowPlaying, MediaItem currentMedia)
        {
            OpResult opResult = new OpResult();
            opResult.StatusCode = OpStatusCode.Json;

            if (currentMedia != null)
            {                
                opResult.ContentText = JsonConvert.SerializeObject(currentMedia, Formatting.Indented);
                return opResult;
            }
            else if (param.IndexOf("-help") >= 0)
            {
                opResult.StatusCode = OpStatusCode.Ok;
                opResult = showHelp(opResult);
                return opResult;
            }

            debug_last_action = "Execute: Start";
            bool should_enqueue = true;
            int size_x = 0;
            int size_y = 0;
            string create_playlist_name = null;
            string playlist_query = null;
            string template = "";
            string cache_fn = make_cache_fn(String.Format("{0}-{1}.txt", which_command, param));
            string cache_body = "";
            try
            {
                IWMPMediaCollection2 collection = (IWMPMediaCollection2)Player.mediaCollection;
                IWMPPlaylistCollection playlistCollection = (IWMPPlaylistCollection)Player.playlistCollection;

                int ver = Player.mediaCollection.getByAttribute("MediaType", "Audio").count;
                string cache_ver = String.Format("{0}", ver);
                cache_body = get_cached(cache_fn, cache_ver);
                if (cache_body.Length != 0 && create_playlist_name == null && playlist_query == null && !m_stats_only)
                {
                    opResult.ContentText = setCachedFlag(cache_body);
                    return opResult;
                }

                IWMPQuery query = collection.createQuery();
                IWMPPlaylistArray playlists = null;
                IWMPPlaylist mediaPlaylist = null;
                IWMPMedia media_item;

                ArrayList query_indexes = new ArrayList();

                Library metadata = new Library(m_stats_only);

                bool has_query = false;
                bool has_exact_query = false;
                
                string query_text = "";
                string query_type = "";

                debug_last_action = "Execution: Parsing params";

                request_params = HttpUtility.UrlEncode(param);

                if (param.Contains("create-playlist:"))
                {
                    create_playlist_name = param.Substring(param.IndexOf("create-playlist:") + "create-playlist:".Length);
                    create_playlist_name = trim_parameter(create_playlist_name);
                }

                if (param.Contains("exact-genre:"))
                {
                    string genre = param.Substring(param.IndexOf("exact-genre:") + "exact-genre:".Length);
                    genre = trim_parameter(genre);
                    query.addCondition("WM/Genre", "Equals", genre);
                    genre_filter = genre;
                    query_text = genre;
                    query_type = "Genre";
                    has_query = true;
                    has_exact_query = true;
                }
                else if (param.Contains("genre:*"))
                {
                    string genre = param.Substring(param.IndexOf("genre:*") + "genre:*".Length);
                    genre = trim_parameter(genre);
                    query.addCondition("WM/Genre", "BeginsWith", genre);
                    query.beginNextGroup();
                    query.addCondition("WM/Genre", "Contains", " " + genre);
                    genre_filter = genre;
                    query_text = genre;
                    query_type = "Genre";
                    has_query = true;
                }
                else if (param.Contains("genre:"))
                {
                    string genre = param.Substring(param.IndexOf("genre:") + "genre:".Length);
                    genre = trim_parameter(genre);
                    query.addCondition("WM/Genre", "BeginsWith", genre);
                    genre_filter = genre;
                    query_text = genre;
                    query_type = "Genre";
                    has_query = true;
                }

                if (param.Contains("exact-artist:"))
                {
                    string artist = param.Substring(param.IndexOf("exact-artist:") + "exact-artist:".Length);
                    artist = trim_parameter(artist);
                    query.addCondition("Author", "Equals", artist);
                    artist_filter = artist;
                    if (query_text.Length > 0)
                    {
                        query_text += ": ";
                        query_type += "/";
                    }
                    query_text += artist;
                    query_type += "Artist";
                    has_query = true;
                    has_exact_query = true;
                }
                else if (param.Contains("exact-album-artist:"))
                {
                    string album_artist = param.Substring(param.IndexOf("exact-album-artist:") + "exact-album-artist:".Length);
                    album_artist = trim_parameter(album_artist);
                    query.addCondition("WM/AlbumArtist", "Equals", album_artist);
                    artist_filter = album_artist;
                    if (query_text.Length > 0)
                    {
                        query_text += ": ";
                        query_type += "/";
                    }
                    query_text += album_artist;
                    query_type += "Album Artist";
                    has_query = true;
                    has_exact_query = true;
                }
                else if (param.Contains("album-artist:*"))
                {
                    string album_artist = param.Substring(param.IndexOf("album-artist:*") + "album-artist:*".Length);
                    album_artist = trim_parameter(album_artist);
                    query.addCondition("WM/AlbumArtist", "BeginsWith", album_artist);
                    query.beginNextGroup();
                    query.addCondition("WM/AlbumArtist", "Contains", " " + album_artist);
                    artist_filter = album_artist;
                    if (query_text.Length > 0)
                    {
                        query_text += ": ";
                        query_type += "/";
                    }
                    query_text += album_artist;
                    query_type += "Album Artist";
                    has_query = true;
                }
                else if (param.Contains("album-artist:"))
                {
                    string album_artist = param.Substring(param.IndexOf("album-artist:") + "album-artist:".Length);
                    album_artist = trim_parameter(album_artist);
                    query.addCondition("WM/AlbumArtist", "BeginsWith", album_artist);
                    artist_filter = album_artist;
                    if (query_text.Length > 0)
                    {
                        query_text += ": ";
                        query_type += "/";
                    }
                    query_text += album_artist;
                    query_type += "Album Artist";
                    has_query = true;
                }
                else if (param.Contains("artist:*"))
                {
                    string artist = param.Substring(param.IndexOf("artist:*") + "artist:*".Length);
                    artist = trim_parameter(artist);
                    query.addCondition("Author", "BeginsWith", artist);
                    query.beginNextGroup();
                    query.addCondition("Author", "Contains", " " + artist);
                    artist_filter = artist;
                    if (query_text.Length > 0)
                    {
                        query_text += ": ";
                        query_type += "/";
                    }
                    query_text += artist;
                    query_type += "Artist";
                    has_query = true;
                }
                else if (param.Contains("artist:"))
                {
                    string artist = param.Substring(param.IndexOf("artist:") + "artist:".Length);
                    artist = trim_parameter(artist);
                    query.addCondition("Author", "BeginsWith", artist);
                    artist_filter = artist;
                    if (query_text.Length > 0)
                    {
                        query_text += ": ";
                        query_type += "/";
                    }
                    query_text += artist;
                    query_type += "Artist";
                    has_query = true;
                }

                if (param.Contains("exact-album:"))
                {
                    string album = param.Substring(param.IndexOf("exact-album:") + "exact-album:".Length);
                    album = trim_parameter(album);
                    query.addCondition("WM/AlbumTitle", "Equals", album);
                    album_filter = album;
                    if (query_text.Length > 0)
                    {
                        query_text += ": ";
                        query_type += "/";
                    }
                    query_text += album;
                    query_type += "Album";
                    has_query = true;
                    has_exact_query = true;
                }
                else if (param.Contains("album:"))
                {
                    string album = param.Substring(param.IndexOf("album:") + "album:".Length);
                    album = trim_parameter(album);
                    query.addCondition("WM/AlbumTitle", "BeginsWith", album);
                    album_filter = album;
                    if (query_text.Length > 0)
                    {
                        query_text += ": ";
                        query_type += "/";
                    }
                    query_text += album;
                    query_type += "Album";
                    has_query = true;
                }

                //This is not for a query but rather for playing/enqueing exact songs
                if (param.Contains("exact-song:"))
                {
                    string song = param.Substring(param.IndexOf("exact-song:") + "exact-song:".Length);
                    song = trim_parameter(song);
                    song_filter = song;
                    if (query_text.Length > 0)
                    {
                        query_text += ": ";
                        query_type += "/";
                    }
                    query_text += song;
                    query_type += "Song";
                    has_exact_query = true;
                }

                if (param.Contains("exact-playlist:"))
                {
                    playlist_query = param.Substring(param.IndexOf("exact-playlist:") + "exact-playlist:".Length);
                    playlist_query = trim_parameter(playlist_query);
                }

                // Indexes specified?
                if (param.Contains("indexes:"))
                {
                    string indexes = param.Substring(param.IndexOf("indexes:") + "indexes:".Length);
                    if (indexes.IndexOf(" ") >= 0) indexes = indexes.Substring(0, indexes.IndexOf(" "));
                    string[] s_idx = indexes.Split(',');
                    foreach (string s in s_idx)
                    {
                        if (s.Length > 0) query_indexes.Add(Int16.Parse(s));
                    }
                    if (query_text.Length > 0)
                    {
                        query_text += ": ";
                        query_type += "/";
                    }
                    query_type += "Tracks";
                    has_query = true;
                }

                if (!has_query) query_type = query_text = "All";

                // Cover size specified?
                if (param.Contains("size-x:"))
                {
                    string tmp_size = param.Substring(param.IndexOf("size-x:") + "size-x:".Length);
                    if (tmp_size.IndexOf(" ") >= 0) tmp_size = tmp_size.Substring(0, tmp_size.IndexOf(" "));
                    size_x = Convert.ToInt32(tmp_size);
                }
                if (param.Contains("size-y:"))
                {
                    string tmp_size = param.Substring(param.IndexOf("size-y:") + "size-y:".Length);
                    if (tmp_size.IndexOf(" ") >= 0) tmp_size = tmp_size.Substring(0, tmp_size.IndexOf(" "));
                    size_y = Convert.ToInt32(tmp_size);
                }
                // Use Custom Template?
                if (param.Contains("template:"))
                {
                    template = param.Substring(param.IndexOf("template:") + "template:".Length);
                    if (template.IndexOf(" ") >= 0) template = template.Substring(0, template.IndexOf(" "));
                }

                if (which_command == PLAY) should_enqueue = false;

                switch (which_command)
                {
                    case CLEAR_CACHE:
                        opResult.StatusCode = OpStatusCode.Ok;
                        clear_cache();                       
                        return opResult;
                    case LIST_GENRES:
                        if (create_playlist_name != null)
                        {
                            IWMPPlaylist genre_playlist = collection.getPlaylistByQuery(query, "Audio", "WM/Genre", true);
                            add_to_playlist(opResult, has_query, genre_playlist, create_playlist_name);
                            opResult.StatusCode = OpStatusCode.Ok;
                        }
                        else
                        {
                            IWMPStringCollection genres;
                            if (has_query)
                            {
                                genres = collection.getStringCollectionByQuery("WM/Genre", query, "Audio", "WM/Genre", true);
                            }
                            else
                            {
                                genres = collection.getAttributeStringCollection("WM/Genre", "Audio");
                            }
                            if (genres != null && genres.count > 0)
                            {
                                result_count = 0;
                                for (int k = 0; k < genres.count; k++)
                                {
                                    string item = genres.Item(k);
                                    if (item != null && !item.Equals(""))
                                    {
                                        if (!m_stats_only) metadata.addGenre(item);
                                        else result_count++;
                                    }
                                }
                                opResult.ResultCount = result_count;
                                metadata.trimToSize();
                                opResult = serializeObject(opResult, metadata, cache_fn, cache_ver);
                            }
                            else
                            {
                                opResult.StatusCode = OpStatusCode.BadRequest;
                                opResult.StatusText = "No genres found!";
                            }
                        }
                        return opResult;
                    case LIST_ARTISTS:
                        if (create_playlist_name != null)
                        {
                            IWMPPlaylist artists_playlist = collection.getPlaylistByQuery(query, "Audio", "Author", true);
                            add_to_playlist(opResult, has_query, artists_playlist, create_playlist_name);
                            opResult.StatusCode = OpStatusCode.Ok;
                        }
                        else
                        {
                            IWMPStringCollection artists;
                            if (has_query)
                            {
                                artists = collection.getStringCollectionByQuery("Author", query, "Audio", "Author", true);
                            }
                            else
                            {
                                artists = collection.getAttributeStringCollection("Author", "Audio");
                            }
                            if (artists != null && artists.count > 0)
                            {
                                result_count = 0;
                                for (int k = 0; k < artists.count; k++)
                                {
                                    string item = artists.Item(k);
                                    if (item != null && !item.Equals(""))
                                    {
                                        if (!m_stats_only) metadata.addArtist(item);
                                        else result_count++;
                                    }
                                }
                                opResult.ResultCount = result_count;
                                metadata.trimToSize();
                                opResult = serializeObject(opResult, metadata, cache_fn, cache_ver);
                            }
                            else
                            {
                                opResult.StatusCode = OpStatusCode.BadRequest;
                                opResult.StatusText = "No artists found!";
                            }
                        }
                        return opResult;
                    case LIST_ALBUM_ARTISTS:
                        if (create_playlist_name != null)
                        {
                            IWMPPlaylist album_artists_playlist = collection.getPlaylistByQuery(query, "Audio", "WM/AlbumArtist", true);
                            add_to_playlist(opResult, has_query, album_artists_playlist, create_playlist_name);
                            opResult.StatusCode = OpStatusCode.Ok;
                        }
                        else
                        {
                            IWMPStringCollection album_artists;
                            if (has_query)
                            {
                                album_artists = collection.getStringCollectionByQuery("WM/AlbumArtist", query, "Audio", "WM/AlbumArtist", true);
                            }
                            else
                            {
                                album_artists = collection.getAttributeStringCollection("WM/AlbumArtist", "Audio");
                            }
                            if (album_artists != null && album_artists.count > 0)
                            {
                                result_count = 0;
                                for (int k = 0; k < album_artists.count; k++)
                                {
                                    string item = album_artists.Item(k);
                                    if (item != null && !item.Equals("") && !metadata.containsAlbumArtist(item))
                                    {
                                        if (!m_stats_only) metadata.addAlbumArtist(item);
                                        else result_count++;
                                    }
                                }
                                opResult.ResultCount = result_count;
                                metadata.trimToSize();
                                opResult = serializeObject(opResult, metadata, cache_fn, cache_ver);
                            }
                            else
                            {
                                opResult.StatusCode = OpStatusCode.BadRequest;
                                opResult.StatusText = "No album artists found!";
                            }
                        }
                        return opResult;
                    case LIST_ALBUMS:
                        if (create_playlist_name != null)
                        {
                            IWMPPlaylist albums_playlist = collection.getPlaylistByQuery(query, "Audio", "WM/AlbumTitle", true);
                            add_to_playlist(opResult, has_query, albums_playlist, create_playlist_name);
                            opResult.StatusCode = OpStatusCode.Ok;
                        }
                        else
                        {
                            IWMPStringCollection albums;
                            if (has_query)
                            {
                                albums = collection.getStringCollectionByQuery("WM/AlbumTitle", query, "Audio", "WM/AlbumTitle", true);
                            }
                            else
                            {
                                albums = collection.getAttributeStringCollection("WM/AlbumTitle", "Audio");
                            }
                            if (albums != null && albums.count > 0)
                            {
                                result_count = 0;
                                for (int k = 0; k < albums.count; k++)
                                {
                                    string item = albums.Item(k);
                                    if (item != null && !item.Equals(""))
                                    {
                                        if (!m_stats_only) metadata.addAlbum(new Album(item, collection.getByAlbum(item), m_stats_only));
                                        else result_count++;
                                    }
                                }
                                opResult.ResultCount = result_count;
                                metadata.trimToSize();
                                opResult = serializeObject(opResult, metadata, cache_fn, cache_ver);
                            }
                            else
                            {
                                opResult.StatusCode = OpStatusCode.BadRequest;
                                opResult.StatusText = "No albums found!";
                            }
                        }
                        return opResult;
                    case LIST_SONGS:
                        IWMPStringCollection songs;
                        if (has_query)
                        {
                            songs = collection.getStringCollectionByQuery("Title", query, "Audio", "Title", true);
                        }
                        else
                        {
                            songs = collection.getAttributeStringCollection("Title", "Audio");
                        }
                        if (songs != null && songs.count > 0)
                        {
                            for (int k = 0; k < songs.count; k++)
                            {
                                IWMPPlaylist playlist = collection.getByName(songs.Item(k));
                                if (playlist != null && playlist.count > 0)
                                {
                                    if (create_playlist_name != null)
                                    {
                                        add_to_playlist(opResult, has_query, playlist, create_playlist_name);
                                        opResult.StatusCode = OpStatusCode.Ok;
                                    }
                                    else
                                    {
                                        metadata.addSongs(playlist);
                                    }
                                }
                            }
                            if (create_playlist_name == null)
                            {
                                metadata.trimToSize();
                                opResult = serializeObject(opResult, metadata, cache_fn, cache_ver);
                            }
                        }
                        else
                        {
                            opResult.StatusCode = OpStatusCode.BadRequest;
                            opResult.StatusText = "No songs found!";
                        }
                        return opResult;
                    case LIST_PLAYLISTS:
                        result_count = 0;
                        playlists = getAllUserPlaylists(playlistCollection);
                        result_count = metadata.addPlaylists(playlistCollection, playlists);
                        metadata.trimToSize();
                        opResult.ResultCount = result_count;
                        opResult.ContentText = JsonConvert.SerializeObject(metadata, Formatting.Indented);
                        return opResult;
                    case LIST_RECENT:
                        if (param.Contains("count:"))
                        {
                            string scount = param.Substring(param.IndexOf("count:") + "count:".Length);
                            if (scount.IndexOf(" ") >= 0) scount = scount.Substring(0, scount.IndexOf(" "));
                            int count = Convert.ToInt32(scount);
                            list_recent(opResult, template, count);
                        }
                        else list_recent(opResult, template);
                        opResult.StatusCode = OpStatusCode.Ok;
                        return opResult;
                    case LIST_NOWPLAYING:
                        if (nowPlaying != null)
                        {
                            opResult.ContentText = JsonConvert.SerializeObject(nowPlaying, Formatting.Indented);
                        }
                        else
                        {
                            opResult.StatusCode = OpStatusCode.BadRequest;
                            opResult.StatusText = "Now playing is null!";
                        }
                        return opResult;
                    case DELETE_PLAYLIST:
                        if (playlist_query != null)
                        {
                            playlists = getUserPlaylistsByName(playlist_query, playlistCollection);
                            if (playlists.count > 0)
                            {
                                IWMPPlaylist mod_playlist = playlists.Item(0);
                                if (query_indexes.Count > 0)
                                {
                                    // Delete items indicated by indexes instead of deleting playlist
                                    for (int j = 0; j < query_indexes.Count; j++)
                                    {
                                        mod_playlist.removeItem(mod_playlist.get_Item((Int16)query_indexes[j]));
                                    }
                                    opResult.ContentText = "Items removed from playlist " + mod_playlist + ".";
                                }
                                else
                                {
                                    ((IWMPPlaylistCollection)Player.playlistCollection).remove(mod_playlist);
                                    opResult.ContentText = "Playlist " + mod_playlist + " deleted.";
                                }
                                opResult.StatusCode = OpStatusCode.Success;
                            }
                            else
                            {
                                opResult.StatusCode = OpStatusCode.BadRequest;
                                opResult.StatusText = "Playlist does not exist!";
                            }
                        }
                        else
                        {
                            opResult.StatusCode = OpStatusCode.BadRequest;
                            opResult.StatusText = "Must specify the exact playlist!";
                        }
                        return opResult;
                    case LIST_DETAILS:
                        // Get  query as a playlist
                        if (playlist_query != null)
                        {
                            //Return a specific playlist when music-list-details with exact-playlist is queried
                            playlists = getUserPlaylistsByName(playlist_query, playlistCollection);
                            if (playlists.count > 0)
                            {
                                Playlist aPlaylist = new Playlist(playlist_query);
                                mediaPlaylist = playlists.Item(0);
                                //Or return a playlist query
                                if (mediaPlaylist != null)
                                {
                                    aPlaylist.addItems(mediaPlaylist);
                                }
                                metadata.playlists.Add(aPlaylist);
                            }
                        }
                        else if (has_exact_query)
                        {
                            mediaPlaylist = getPlaylistFromExactQuery(query_text, query_type, collection);
                        }
                        else if (has_query)
                        {
                            string type = getSortAttributeFromQueryType(query_type);
                            mediaPlaylist = collection.getPlaylistByQuery(query, "Audio", type, true);
                        }

                        if (mediaPlaylist != null)
                        {
                            //Create playlist from query if supplied with playlist name
                            if (create_playlist_name != null)
                            {
                                add_to_playlist(opResult, has_query, mediaPlaylist, create_playlist_name);
                                return opResult;
                            }
                            else if (query_indexes.Count > 0)
                            {
                                for (int j = 0; j < query_indexes.Count; j++)
                                {
                                    media_item = mediaPlaylist.get_Item((Int16)query_indexes[j]);
                                    if (media_item != null)
                                    {
                                        metadata.addSong(media_item);
                                    }
                                }
                            }
                            else
                            {
                                if (query_type.Equals("Album"))
                                {
                                    Album album = new Album(query_text, m_stats_only);
                                    album.addTracks(mediaPlaylist);
                                    metadata.addAlbum(album);
                                }
                                else
                                {
                                    metadata.addSongs(mediaPlaylist);
                                }
                            }
                        }
                        else
                        {
                            if (logger != null)
                            {
                                logger.Write("Creating library metadata object");
                            }
                            //No query supplied so entire detailed library requested
                            //Parse all albums and return, no value album will be added as songs                            
                            IWMPStringCollection album_collection = collection.getAttributeStringCollection("WM/AlbumTitle", "Audio");
                            if (album_collection.count > 0)
                            {
                                result_count = 0;
                                for (int j = 0; j < album_collection.count; j++)
                                {
                                    if (album_collection.Item(j) != null)
                                    {
                                        //The collection seems to represent the abcense of an album as an "" string value
                                        IWMPPlaylist album_playlist = collection.getByAlbum(album_collection.Item(j));
                                        if (album_playlist != null)
                                        {                                                                                        
                                            if (!album_collection.Item(j).Equals(""))
                                            {
                                                Album album = new Album(album_collection.Item(j), m_stats_only);
                                                result_count += album.addTracks(album_playlist);
                                                metadata.addAlbum(album);
                                            }
                                            else
                                            {
                                                result_count += metadata.addSongs(album_playlist);
                                            }
                                        }
                                    }
                                }
                                metadata.trimToSize();
                            }
                        }
                        if (logger != null)
                        {
                            logger.Write("Starting serialization of metadata object.");
                        }
                        opResult.ResultCount = result_count;
                        opResult = serializeObject(opResult, metadata, cache_fn, cache_ver);                        
                        return opResult;
                    case PLAY:
                    case QUEUE:
                        if (has_exact_query)
                        {
                            mediaPlaylist = getPlaylistFromExactQuery(query_text, query_type, collection);
                        }
                        else if (has_query)
                        {
                            string type = getSortAttributeFromQueryType(query_type);
                            mediaPlaylist = collection.getPlaylistByQuery(query, "Audio", type, true);
                        }
                        else
                        {
                            mediaPlaylist = collection.getByAttribute("MediaType", "Audio");
                        }
                        //Play or enqueue
                        PlayMediaCmd pmc;
                        if (query_indexes.Count > 0)
                        {
                            result_count = query_indexes.Count;
                            for (int j = 0; j < query_indexes.Count; j++)
                            {
                                media_item = mediaPlaylist.get_Item(j);
                                if (media_item != null)
                                {
                                    query_text += ((j == 0) ? "" : ", ") + (Int16)query_indexes[j] + ". " + media_item.getItemInfo("Title");
                                }
                            }
                            pmc = new PlayMediaCmd(remotePlayer, mediaPlaylist, query_indexes, should_enqueue);
                        }
                        else
                        {
                            result_count = mediaPlaylist.count;
                            pmc = new PlayMediaCmd(remotePlayer, mediaPlaylist, should_enqueue);
                        }
                        opResult = pmc.Execute(null);

                        // Type, Artist, Album, Track, param, count
                        add_to_mrp(query_type, query_text, param, result_count); //Add to recent played list
                        return opResult;
                    case SERV_COVER:
                        if (has_exact_query)
                        {
                            mediaPlaylist = getPlaylistFromExactQuery(query_text, query_type, collection);
                        }
                        else if (has_query)
                        {
                            string type = getSortAttributeFromQueryType(query_type);
                            mediaPlaylist = collection.getPlaylistByQuery(query, "Audio", type, true);
                        }
                        else
                        {
                            mediaPlaylist = collection.getByAttribute("MediaType", "Audio");
                        }

                        try
                        {
                            if (query_indexes.Count > 0)
                            {
                                for (int j = 0; j < query_indexes.Count; j++)
                                {
                                    media_item = mediaPlaylist.get_Item((Int16)query_indexes[j]);
                                    if (media_item != null)
                                    {
                                        string album_path = findAlbumPath(media_item.sourceURL);
                                        photoCmd pc = new photoCmd(photoCmd.SERV_PHOTO);
                                        if (album_path.Length == 0) return pc.getPhoto(DEFAULT_IMAGE, "jpeg", size_x, size_y);
                                        else return pc.getPhoto(album_path, "jpeg", size_x, size_y);
                                    }
                                }
                            }
                            else
                            {
                                for (int j = 0; j < mediaPlaylist.count; j++)
                                {
                                    media_item = mediaPlaylist.get_Item(j);
                                    if (media_item != null)
                                    {
                                        string album_path = findAlbumPath(media_item.sourceURL);
                                        photoCmd pc = new photoCmd(photoCmd.SERV_PHOTO);
                                        if (album_path.Length == 0) return pc.getPhoto(DEFAULT_IMAGE, "jpeg", size_x, size_y);
                                        else return pc.getPhoto(album_path, "jpeg", size_x, size_y);
                                    }
                                }
                            }
                        }
                        catch (Exception ex) 
                        {
                            opResult.StatusCode = OpStatusCode.Exception;
                            opResult.StatusText = ex.Message;
                        }                  
                        return opResult;
                }
            }
            catch (Exception ex)
            {
                opResult = new OpResult();
                opResult.StatusCode = OpStatusCode.Exception;
                opResult.StatusText = ex.Message;
                opResult.AppendFormat("{0}", debug_last_action);
                opResult.AppendFormat("{0}", ex.Message);
            }

            debug_last_action = "Execute: End";

            return opResult;
        }

        private OpResult serializeObject(OpResult opResult, object metadata, string fn, string cur_ver)
        {
            opResult.ContentText = JsonConvert.SerializeObject(metadata, Formatting.Indented);
            if (logger != null)
            {
                logger.Write("Serialization finished.");
            }
            if (logger != null)
            {
                logger.Write("Writing to cache.");
            }
            save_to_cache(fn, opResult.ToString(), cur_ver);
            if (logger != null)
            {
                logger.Write("Writing to cache finished.");
                logger.Close();
            }
            if (m_cache_only)
            {
                opResult = new OpResult();
                opResult.StatusCode = OpStatusCode.Ok;
                opResult.StatusText = "Cache saved";
            }
            return opResult;
        }

        private string setCachedFlag(string cache_body)
        {
            JObject jObject = JObject.Parse(cache_body);
            jObject["from_cache"] = true;
            return jObject.ToString();
        }

        public class Library
        {
            public bool from_cache = false;
            private bool m_stats_only = false;

            public ArrayList albums = new ArrayList();
            public ArrayList songs = new ArrayList();
            public ArrayList genres = new ArrayList();
            public ArrayList artists = new ArrayList();
            public ArrayList album_artists = new ArrayList();
            public ArrayList playlists = new ArrayList();
            
            public Library(bool stats_only)
            {
                m_stats_only = stats_only;
            }

            public void addAlbum(Album album)
            {
                albums.Add(album);
            }

            public void addSong(IWMPMedia item)
            {
                songs.Add(new Song(item));
            }

            public int addSongs(IWMPPlaylist playlist)
            {
                int result_count = 0;
                for (int j = 0; j < playlist.count; j++)
                {
                    IWMPMedia item = playlist.get_Item(j);
                    if (item != null)
                    {
                        if (!m_stats_only) songs.Add(new Song(item));
                        else result_count++;
                    }
                }
                return result_count;
            }

            public void addGenre(string genre)
            {
                genres.Add(genre);
            }

            public void addArtist(string artist)
            {
                artists.Add(artist);
            }

            public bool containsAlbumArtist(string album_artist)
            {
                return album_artists.Contains(album_artist);
            }

            public void addAlbumArtist(string album_artist)
            {
                album_artists.Add(album_artist);
            }

            public int addPlaylists(IWMPPlaylistCollection playlistCollection, IWMPPlaylistArray list)
            {
                int result_count = 0;
                for (int j = 0; j < list.count; j++)
                {
                    bool containsAudio = false;
                    IWMPPlaylist playlist = list.Item(j);
                    string name = playlist.name;

                    if (!name.Equals("All Music") && !name.Contains("TV") && !name.Contains("Video") && !name.Contains("Pictures"))
                    {
                        for (int k = 0; k < playlist.count; k++)
                        {
                            try
                            {
                                if (playlist.get_Item(k).getItemInfo("MediaType").Equals("audio") && !playlistCollection.isDeleted(playlist))
                                {
                                    containsAudio = true;
                                }
                            }
                            catch (Exception)
                            {
                                //Ignore playlists with invalid items
                            }
                        }
                    }

                    if (containsAudio)
                    {
                        if (!m_stats_only) playlists.Add(new Playlist(name));
                        else result_count++;
                    }
                }
                return result_count;
            }

            public void trimToSize()
            {
                albums.TrimToSize();
                songs.TrimToSize();
                genres.TrimToSize();
                artists.TrimToSize();
                album_artists.TrimToSize();
                playlists.TrimToSize();
            }
        }

        public class Album
        {
            public string album = "";
            public string album_artist = "";
            public string year = "";
            public ArrayList genre = new ArrayList();
            public ArrayList tracks = new ArrayList();
            private bool m_stats_only = false;

            public Album(string name, IWMPPlaylist playlist, bool stats_only)
            {
                m_stats_only = stats_only;
                album = name;
                if (playlist != null)
                {
                    int count = 0;
                    if (playlist.count > 1) count = 2;
                    else count = 1;
                    for (int j = 0; j < count; j++)
                    {
                        IWMPMedia item = playlist.get_Item(j);
                        if (item != null)
                        {
                            album_artist = item.getItemInfo("WM/AlbumArtist");
                            year = item.getItemInfo("WM/OriginalReleaseYear");
                            if (year.Equals("") || year.Length < 4) year = item.getItemInfo("WM/Year");
                            if (!genre.Contains(item.getItemInfo("WM/Genre"))) genre.Add(item.getItemInfo("WM/Genre"));
                        }
                    }
                }
            }

            public Album(string name, bool stats_only)
            {
                m_stats_only = stats_only;
                album = name;
            }

            public int addTracks(IWMPPlaylist playlist)
            {
                int result_count = 0;
                for (int j = 0; j < playlist.count; j++)
                {
                    IWMPMedia item = playlist.get_Item(j);
                    if (item != null)
                    {
                        if (!m_stats_only) addTrack(item);
                        else result_count++;
                    }
                }
                return result_count;
            }

            public void addTrack(IWMPMedia item)
            {
                Track track = new Track(item);
                if (!genre.Contains(track.genre)) genre.Add(track.genre);
                year = track.year;
                album_artist = item.getItemInfo("WM/AlbumArtist");
                tracks.Add(track);
            }
        }

        public class Playlist
        {
            public string playlist = "";
            public ArrayList items = new ArrayList();

            public Playlist(string name)
            {
                playlist = name;
            }

            public void addItems(IWMPPlaylist playlist)
            {
                for (int j = 0; j < playlist.count; j++)
                {
                    IWMPMedia item = playlist.get_Item(j);
                    if (item != null)
                    {
                        items.Add(new PlaylistItem(j, item));
                    }
                }
            }
        }

        public class PlaylistItem : Song
        {
            public string number = "";

            public PlaylistItem(int index, IWMPMedia item)
                : base(item)
            {
                number = Convert.ToString(index + 1);
            }
        }

        public class Track : Song
        {
            public string track_number = "";

            public Track(IWMPMedia item) : base(item)
            {
                track_number = item.getItemInfo("WM/TrackNumber");
            }
        }

        public class Song
        {
            public string song = "";
            public string song_artist = "";
            public string genre = "";
            public string year = "";
            public string duration = "";

            public Song()
            {
            }

            public Song(IWMPMedia item)
            {
                song = item.getItemInfo("Title");
                song_artist = item.getItemInfo("Author");
                duration = item.durationString;
                year = item.getItemInfo("WM/OriginalReleaseYear");
                if (year.Equals("") || year.Length < 4) year = item.getItemInfo("WM/Year");
                genre = item.getItemInfo("WM/Genre");
            }
        }

        #endregion
    }
}
