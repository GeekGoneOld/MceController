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
using VmcController.AddIn.Metadata;


namespace VmcController.AddIn.Commands
{

    /// <summary>
    /// Summary description for getArtists commands.
    /// </summary>
    public class MusicCmd : ICommand
    {
        private RemotedWindowsMediaPlayer remotePlayer = null;
        private WindowsMediaPlayer Player = null;

        private string debug_last_action = "none";

        public const int LIST_ARTISTS = 1;
        public const int LIST_ALBUMS = 2;
        public const int LIST_SONGS = 3;
        public const int LIST_DETAILS = 4;
        public const int PLAY = 5;
        public const int QUEUE = 6;
        public const int SERV_COVER = 7;
        public const int CLEAR_CACHE = 8;
        public const int LIST_ALBUM_ARTISTS = 10;
        public const int LIST_GENRES = 11;
        public const int LIST_RECENT = 12;
        public const int LIST_STATS = 13;
        public const int LIST_PLAYLISTS = 14;
        public const int LIST_NOWPLAYING = 15;
        public const int LIST_CURRENT = 16;
        public const int DELETE_PLAYLIST = 17;
        public const int SHUFFLE = 18;

        private int which_command = -1;

        private static Dictionary<string, string> m_templates = new Dictionary<string, string>();
        private int result_count = 0;
        private string artist_filter = "";
        private string album_filter = "";
        private string genre_filter = "";
        private string song_filter = "";
        private string request_params = "";

        public static string CACHE_MUSIC_CMD_DIR = AddInModule.DATA_DIR + "\\music_cmd_cache";
        public static string CACHE_VER_FILE = CACHE_MUSIC_CMD_DIR + "\\ver";

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

        private void init()
        {
            if (!init_run)
            {
                init_run = true;
                loadTemplate();
            }
            Player = new WindowsMediaPlayer();
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
                    s = " list all playlists";
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
                    s = " changes play state to shuffle";
                    break;
                case SERV_COVER:
                    s = "[size-x:<width>] [size-y:<height>] [[exact-]artist:[*]artist_filter] [[exact-]album:[*]album_filter] [[exact-]genre:[*]genre_filter] [indexes:id1,id2] - serves the album cover of the first match";
                    break;
                case CLEAR_CACHE:
                    s = " clears and rebuilds the full cache";
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
                    s = "- returns state in JSON objects <current_track, play_state, volume, shuffle_mode, is_muted>, current_track contains currently playing metadata";
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

        public static bool delete_old_cache(string cur_ver)
        {
            //If number of audio items has changed (i.e. cur_ver) clear the cache
            try
            {
                if (get_is_cache_outdated(cur_ver))
                {
                    clear_cache();
                    return true;
                }
            }
            catch (Exception) 
            {
                //Return true here to indicate there's no cache
                clear_cache();
                return true;
            }
            return false;
        }

        public static void clear_cache()
        {
            try { System.IO.Directory.Delete(CACHE_MUSIC_CMD_DIR, true); }
            catch (Exception) { }
        }

        public static string get_cache_filepath(int command, string param)
        {
            string filename = String.Format("{0}-{1}.txt", command, param);
            filename = filename.Replace("\\", "_");
            filename = filename.Replace(":", "_");
            filename = filename.Replace(" ", "%20");
            filename = filename.Replace("\"", "_");
            return CACHE_MUSIC_CMD_DIR + "\\" + filename;
        }

        public static string get_audio_item_count(IWMPMediaCollection2 collection)
        {
            int ver = collection.getByAttribute("MediaType", "Audio").count;
            return String.Format("{0}", ver);
        }

        public static bool get_is_cache_outdated(string cur_ver)
        {                        
            try
            {
                string cache_ver = System.IO.File.ReadAllText(CACHE_VER_FILE);
                return !cur_ver.Equals(cache_ver);
            }
            catch (Exception)
            {
                return true;                
            }            
        }

        public void save_to_cache(string param, string opContent, string cur_ver)
        {
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
                System.IO.File.WriteAllText(get_cache_filepath(which_command, param), opContent);
            }
            catch (Exception)
            {
            }
        }

        public OpResult get_modified_cache(string cache_body)
        {
            JObject jObject = JObject.Parse(cache_body);
            jObject["from_cache"] = true;
            return new OpResult(jObject.ToString());
        }

        public string get_cached(string param, string cur_ver)
        {
            string cached = "";
            if (!param.Contains("stats") || !get_is_cache_outdated(cur_ver))
            {
                try
                {
                    string cached_file_path = get_cache_filepath(which_command, param);

                    FileInfo fi = new FileInfo(cached_file_path);
                    if (fi.Exists)
                    {
                        cached = System.IO.File.ReadAllText(cached_file_path);
                    }
                }
                catch (Exception)
                {
                } 
            }
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

        private OpResult add_to_playlist(IWMPPlaylist queried, string playlist_name)
        {
            OpResult opResult = new OpResult();
            PlaylistOpResultObject status = new PlaylistOpResultObject();
            status.playlist_name = playlist_name;
            IWMPPlaylistCollection playlistCollection = (IWMPPlaylistCollection)Player.playlistCollection;
            IWMPPlaylistArray current_playlists = playlistCollection.getByName(playlist_name);
            if (current_playlists.count > 0)
            {
                IWMPPlaylist previous_playlist = current_playlists.Item(0);
                for (int j = 0; j < queried.count; j++)
                {
                    previous_playlist.appendItem(queried.get_Item(j));
                }
                status.existing_playlist = true;
                opResult.StatusText = "Playlist updated";
            }
            else
            {
                queried.name = playlist_name;
                playlistCollection.importPlaylist(queried);
                status.existing_playlist = false;
                opResult.StatusText = "Playlist added";
            }
            opResult.StatusCode = OpStatusCode.Success;
            opResult.ContentObject = status;
            return opResult;
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

        private IWMPPlaylist getPlaylistFromExactQuery(string query_text, string query_type, IWMPMediaCollection2 collection, IWMPPlaylistCollection playlistCollection)
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
            }
            else if (query_type.Equals("Playlist"))
            {
                //Return a specific playlist when music-list-details with exact-playlist is queried
                IWMPPlaylistArray playlists = getUserPlaylistsByName(query_text, playlistCollection);
                if (playlists.count > 0)
                {
                    return playlists.Item(0);
                }
            }
            return null;
        }

        public ArrayList sortQueriedPlaylist(string query_type, IWMPPlaylist playlist)
        {
            List<PlaylistTrack> sortedList = null;                       
            if (query_type.Equals("Album")) { }
            else if (query_type.Equals("Album Artist")) { }
            else if (query_type.Equals("Artist"))
            {
                sortedList = new List<PlaylistTrack>();
            }
            else if (query_type.Equals("Genre"))
            {
                sortedList = new List<PlaylistTrack>();
            }
            else if (query_type.Equals("Song")) { }
            else if (query_type.Equals("Playlist")) { }
            if (sortedList != null)
            {
                for (int j = 0; j < playlist.count; j++)
                {
                    sortedList.Add(new PlaylistTrack(j, playlist.get_Item(j)));
                }
                sortedList.Sort();
                ArrayList indexes = new ArrayList();
                foreach (PlaylistTrack item in sortedList)
                {
                    indexes.Add(item.playlist_number - 1);
                }
                indexes.TrimToSize();
                return indexes;
            }
            else
            {
                return null;
            }            
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

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult Execute(string param)
        {
            return Execute(param, null);
        }

        public OpResult ExecuteCacheBuild(Logger logger)
        {
            return Execute("", logger);
        }

        public OpResult ExecuteStats()
        {
            return Execute("stats", null);
        }

        private OpResult Execute(string param, Logger logger)
        {
            OpResult opResult = new OpResult();
            opResult.StatusCode = OpStatusCode.Ok;
            
            if (param.IndexOf("-help") >= 0)
            {
                opResult = showHelp(opResult);
                return opResult;
            }

            debug_last_action = "Execute: Start";
            bool should_enqueue = true;
            int size_x = 0;
            int size_y = 0;
            string create_playlist_name = null;
            string template = "";
            string cache_body = "";
            try
            {               
                IWMPMediaCollection2 collection = (IWMPMediaCollection2)Player.mediaCollection;
                IWMPPlaylistCollection playlistCollection = (IWMPPlaylistCollection)Player.playlistCollection;
                                
                IWMPQuery query = collection.createQuery();
                IWMPPlaylistArray playlistArray = null;
                IWMPPlaylist mediaPlaylist = null;
                IWMPMedia media_item;

                string cur_ver = get_audio_item_count(collection);
                cache_body = get_cached(param, cur_ver);

                Library metadata = new Library(param);
                ArrayList query_indexes = new ArrayList();                

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
                    has_query = true;
                    has_exact_query = true;
                }

                if (param.Contains("exact-playlist:"))
                {
                    string playlist = param.Substring(param.IndexOf("exact-playlist:") + "exact-playlist:".Length);
                    query_text += trim_parameter(playlist);
                    query_type += "Playlist";
                    has_query = true;
                    has_exact_query = true;
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

                //Return cached results if not querying or creating a playlist
                if (cache_body.Length != 0 && create_playlist_name == null && !query_type.Equals("Playlist"))
                {
                    return get_modified_cache(cache_body);
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
                        clear_cache();
                        opResult.StatusCode = OpStatusCode.Success;
                        opResult.StatusText = "Cache cleared";
                        return opResult;
                    case LIST_GENRES:
                        if (create_playlist_name != null)
                        {
                            IWMPPlaylist genre_playlist = collection.getPlaylistByQuery(query, "Audio", "WM/Genre", true);
                            opResult = add_to_playlist(genre_playlist, create_playlist_name);
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
                                
                                opResult.ResultCount = metadata.addGenres(genres);
                                opResult = process_result(opResult, metadata, param, cur_ver, logger);
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
                            opResult = add_to_playlist(artists_playlist, create_playlist_name);
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
                                opResult.ResultCount = metadata.addArtists(artists);
                                opResult = process_result(opResult, metadata, param, cur_ver, logger);
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
                            opResult = add_to_playlist(album_artists_playlist, create_playlist_name);
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
                                opResult.ResultCount = metadata.addAlbumArtists(album_artists);
                                opResult = process_result(opResult, metadata, param, cur_ver, logger);
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
                            opResult = add_to_playlist(albums_playlist, create_playlist_name);                            
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

                                opResult.ResultCount = metadata.addAlbums(albums, collection);
                                opResult = process_result(opResult, metadata, param, cur_ver, logger);
                            }
                            else
                            {
                                opResult.StatusCode = OpStatusCode.BadRequest;
                                opResult.StatusText = "No albums found!";
                            }
                        }
                        return opResult;
                    case LIST_SONGS:
                        IWMPStringCollection song_collection;
                        if (has_query)
                        {
                            song_collection = collection.getStringCollectionByQuery("Title", query, "Audio", "Title", true);
                        }
                        else
                        {
                            song_collection = collection.getAttributeStringCollection("Title", "Audio");
                        }
                        if (song_collection != null && song_collection.count > 0)
                        {                            
                            mediaPlaylist = Player.newPlaylist(create_playlist_name, null);
                            for (int k = 0; k < song_collection.count; k++)
                            {
                                IWMPPlaylist playlist = collection.getByName(song_collection.Item(k));
                                if (playlist != null && playlist.count > 0)
                                {
                                    mediaPlaylist.appendItem(playlist.get_Item(0));
                                }
                            }
                            if (create_playlist_name != null)
                            {
                                opResult = add_to_playlist(mediaPlaylist, create_playlist_name);    
                            }
                            else
                            {
                                opResult.ResultCount = metadata.addSongs(mediaPlaylist);
                                opResult = process_result(opResult, metadata, param, cur_ver, logger);
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
                        playlistArray = getAllUserPlaylists(playlistCollection);
                        result_count = metadata.addPlaylists(playlistCollection, playlistArray);
                        opResult.ResultCount = result_count;
                        opResult.StatusCode = OpStatusCode.Success;
                        opResult.ContentObject = metadata;
                        return opResult;  //bypass cache
                    case LIST_RECENT:
                        if (param.Contains("count:"))
                        {
                            string scount = param.Substring(param.IndexOf("count:") + "count:".Length);
                            if (scount.IndexOf(" ") >= 0) scount = scount.Substring(0, scount.IndexOf(" "));
                            int count = Convert.ToInt32(scount);
                            list_recent(opResult, template, count);
                        }
                        else list_recent(opResult, template);
                        opResult.StatusCode = OpStatusCode.Success;
                        return opResult;
                    case DELETE_PLAYLIST:
                        if (query_type.Equals("Playlist"))
                        {                            
                            playlistArray = getUserPlaylistsByName(query_text, playlistCollection);
                            if (playlistArray.count > 0)
                            {
                                PlaylistOpResultObject playlistStatus = new PlaylistOpResultObject();
                                playlistStatus.playlist_name = query_text;

                                IWMPPlaylist mod_playlist = playlistArray.Item(0);                                
                                if (query_indexes.Count > 0)
                                {
                                    // Delete items indicated by indexes instead of deleting playlist
                                    for (int j = 0; j < query_indexes.Count; j++)
                                    {
                                        mod_playlist.removeItem(mod_playlist.get_Item((Int16)query_indexes[j]));
                                    }
                                    playlistStatus.existing_playlist = true;
                                    playlistStatus.playlist_deleted = false;
                                    playlistStatus.deleted_indexes = query_indexes;
                                    opResult.StatusText = "Songs removed from playlist";
                                }
                                else
                                {
                                    ((IWMPPlaylistCollection)Player.playlistCollection).remove(mod_playlist);
                                    playlistStatus.existing_playlist = true;
                                    playlistStatus.playlist_deleted = true;
                                    opResult.StatusText = "Playlist deleted";
                                }
                                opResult.StatusCode = OpStatusCode.Success;
                                opResult.ContentObject = playlistStatus;
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
                            opResult.StatusText = "Must specify exact playlist!";                            
                        }
                        return opResult;
                    case LIST_DETAILS:
                        // Get  query as a playlist                        
                        if (has_exact_query)
                        {
                            mediaPlaylist = getPlaylistFromExactQuery(query_text, query_type, collection, playlistCollection);
                        }
                        else if (has_query)
                        {
                            string type = getSortAttributeFromQueryType(query_type);
                            mediaPlaylist = collection.getPlaylistByQuery(query, "Audio", type, true);
                        }
                        //Create playlist from query if supplied with playlist name
                        if (create_playlist_name != null)
                        {
                            if (mediaPlaylist != null)
                            {
                                return add_to_playlist(mediaPlaylist, create_playlist_name);
                            }
                            else
                            {
                                opResult.StatusCode = OpStatusCode.BadRequest;
                                opResult.StatusText = "Playlists can only be created using a query for specific items.";
                            }                                
                        }
                        else if (mediaPlaylist != null)
                        {                            
                            result_count = 0;
                            if (query_indexes.Count > 0)
                            {
                                metadata.addSongs(mediaPlaylist, query_indexes);
                            }
                            else if (has_exact_query && query_type.Equals("Playlist"))
                            {
                                ArrayList items = new ArrayList();
                                Playlist item = new Playlist(query_text, metadata.is_stats_only);
                                result_count = item.addTracks(mediaPlaylist);
                                items.Add(item);
                                if (!metadata.is_stats_only)
                                {
                                    items.TrimToSize();
                                    metadata.playlists = items;
                                }                              
                            }
                            else if (has_exact_query && query_type.Equals("Album"))
                            {
                                ArrayList items = new ArrayList();
                                Album item = new Album(query_text, metadata.is_stats_only);
                                result_count = item.addTracks(mediaPlaylist);
                                items.Add(item);
                                if (!metadata.is_stats_only)
                                {
                                    items.TrimToSize();
                                    metadata.albums = items;
                                }
                            }
                            else
                            {
                                //Generate album
                                ArrayList albums = new ArrayList();
                                ArrayList songs = new ArrayList();
                                for (int j = 0; j < mediaPlaylist.count; j++)
                                {
                                    IWMPMedia item = mediaPlaylist.get_Item(j);
                                    if (item != null)
                                    {
                                        result_count++;
                                        string album_name = item.getItemInfo("WM/AlbumTitle");
                                        if (album_name != null && !album_name.Equals(""))
                                        {
                                            Album album = new Album(album_name, metadata.is_stats_only);
                                            album.tracks = new ArrayList();
                                            album.addTrack(album.tracks, item);
                                            int index = albums.IndexOf(album);
                                            if (index == -1) albums.Add(album);
                                            else
                                            {
                                                album = (Album)albums[index];
                                                album.addTrack(album.tracks, item);                                                
                                            }
                                        }
                                        else
                                        {
                                            Song song = new Song(item);
                                            if (!songs.Contains(song)) songs.Add(song);
                                        }
                                    }
                                }
                                if (!metadata.is_stats_only)
                                {
                                    metadata.albums = albums;
                                    metadata.songs = songs;
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
                                ArrayList albums = new ArrayList();
                                for (int j = 0; j < album_collection.count; j++)
                                {
                                    string name = album_collection.Item(j);
                                    if (name != null)
                                    {
                                        //The collection seems to represent the absence of an album as an "" string value
                                        IWMPPlaylist album_playlist = collection.getByAlbum(name);
                                        if (album_playlist != null)
                                        {                                                                                        
                                            if (!name.Equals(""))
                                            {
                                                Album item = new Album(name, metadata.is_stats_only);
                                                int count = item.addTracks(album_playlist);
                                                if (!albums.Contains(item))
                                                {
                                                    albums.Add(item);
                                                    result_count += count;
                                                }
                                            }
                                            else
                                            {
                                                result_count += metadata.addSongs(album_playlist);
                                            }
                                        }
                                    }
                                }                                
                                if (!metadata.is_stats_only)
                                {
                                    albums.TrimToSize();
                                    metadata.albums = albums;
                                }
                            }
                        }
                        if (logger != null)
                        {
                            logger.Write("Starting serialization of metadata object.");
                        }
                        opResult.ResultCount = result_count;
                        opResult = process_result(opResult, metadata, param, cur_ver, logger);                        
                        return opResult;
                    case PLAY:
                    case QUEUE:
                        if (has_exact_query)
                        {
                            mediaPlaylist = getPlaylistFromExactQuery(query_text, query_type, collection, playlistCollection);
                            //Since the Playlist.moveItem method does not seem to work                            
                            //if (query_indexes.Count <= 0)
                            //{
                            //    query_indexes = sortQueriedPlaylist(query_type, mediaPlaylist);
                            //}
                            //On second thought, this may not be best for playing, e.g., all songs by genre
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
                                    query_text += ((j == 0) ? "" : ", ") + (int)query_indexes[j] + ". " + media_item.getItemInfo("Title");
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
                            mediaPlaylist = getPlaylistFromExactQuery(query_text, query_type, collection, playlistCollection);
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
                            opResult = new OpResult();
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

        private OpResult process_result(OpResult opResult, Library metadata, string param, string cur_ver, Logger logger)
        {
            opResult.StatusCode = OpStatusCode.Success;
            opResult.ContentObject = metadata;            
            if (logger != null)
            {
                logger.Write("Writing to cache.");
            }
            //If stats_only, param will be "stats" so it won't kill the cache for each respective full metadata query
            save_to_cache(param, opResult.ToString(), cur_ver);            
            if (logger != null)
            {
                logger.Write("Writing to cache finished.");
                logger.Close();
            }
            if (logger != null)
            {
                opResult = new OpResult();
                opResult.StatusCode = OpStatusCode.Ok;
                opResult.StatusText = "Cache saved";
            } 
            return opResult;
        }

        public class PlaylistOpResultObject : OpResultObject
        {
            public ArrayList deleted_indexes = new ArrayList();
            public bool playlist_deleted = false;
            public bool existing_playlist = false;
            public string playlist_name = "";
        }

        #endregion
    }
}
