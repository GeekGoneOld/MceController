/*
 * This module is build on top of on J.Bradshaw's vmcController
 * Implements audio media library functions
 * 
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

namespace VmcController.AddIn.Commands
{
    /// <summary>
    /// Summary description for getArtists commands.
    /// </summary>
    public class musicCmd : ICommand
    {
        private static WMPLib.WindowsMediaPlayer Player = null;
        //private WMPLib.IWMPMedia media;
        //private WMPLib.IWMPPlaylist mediaPlaylist;

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

        private int which_command = -1;

        private static Dictionary<string, string> m_templates = new Dictionary<string, string>();
        private state the_state = new state();
        private int result_count = 0;
        private string artist_filter = "";
        private string album_filter = "";
        private string genre_filter = "";
        private string request_params = "";

        private static string CACHE_DIR = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\VMC_Controller";
        private static string CACHE_MUSIC_CMD_DIR = CACHE_DIR + "\\music_cmd_cache";
        private static string CACHE_VER_FILE = CACHE_MUSIC_CMD_DIR + "\\ver";

        private const string DEFAULT_DETAIL_ARTIST_START = "artist=%artist%";
        private const string DEFAULT_DETAIL_ALBUM_START = "     album=%album% (%albumYear%; %albumGenre%)";
        private const string DEFAULT_DETAIL_SONG = "          %if-songTrackNumber%track=%songTrackNumber%. %endif%song=%song% (%songLength%)";
        private const string DEFAULT_DETAIL_ALBUM_END = "          total album tracks=%albumTrackCount%";
        private const string DEFAULT_DETAIL_ARTIST_END = "     total artist tracks=%artistTrackCount%";
        private const string DEFAULT_DETAIL_FOOTER = "total artists found=%artistCount%\r\ntotal albums found=%albumCount%\r\ntotal tracks found=%trackCount%";
        private const string DEFAULT_STATS = "track_count=%track_count%\r\nartist_count=%artist_count%\r\nalbum_count=%album_count%\r\ngenre_count=%genre_count%\r\ncache_age=%cache_age%\r\navailable_templates=%available_templates%";

        private const string DEFAULT_IMAGE = "default.jpg";

        private static bool init_run = false;

        public musicCmd(int i)
        {
            which_command = i;

            if (!init_run)
            {
                init_run = true;
                loadTemplate();
                if (Player == null) Player = new WMPLib.WindowsMediaPlayer();
            }
        }

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            string s = "[-help] [[exact_]artist:[*]artist_filter] [[exact_]album:[*]album_filter] [[exact_]genre:[*]genre_filter] [indexes:id1,id2] [template:template_name]- list / play from audio collection";
            switch (which_command)
            {
                case LIST_ARTISTS:
                    s = "[-help] [[exact_]artist:[*]artist_filter] [[exact_]album:[*]album_filter] [[exact_]genre:[*]genre_filter] [template:template_name] - lists matching artists";
                    break;
                case LIST_ALBUM_ARTISTS:
                    s = "[-help] [[exact_]artist:[*]artist_filter] [[exact_]album:[*]album_filter] [[exact_]genre:[*]genre_filter] [template:template_name] - lists matching album artists";
                    break;
                case LIST_ALBUMS:
                    s = "[-help] [[exact_]artist:[*]artist_filter] [[exact_]album:[*]album_filter] [[exact_]genre:[*]genre_filter] [template:template_name] - list matching albums";
                    break;
                case LIST_SONGS:
                    s = "[-help] [[exact_]artist:[*]artist_filter] [[exact_]album:[*]album_filter] [[exact_]genre:[*]genre_filter] [template:template_name] - list matching songs";
                    break;
                case LIST_GENRES:
                    s = "[-help] [[exact_]artist:[*]artist_filter] [[exact_]album:[*]album_filter] [[exact_]genre:[*]genre_filter] [template:template_name] - list matching genres";
                    break;
                case LIST_DETAILS:
                    s = "[-help] [[exact_]artist:[*]artist_filter] [[exact_]album:[*]album_filter] [[exact_]genre:[*]genre_filter] [indexes:id1,id2] [template:template_name] - lists info on matching songs / albums / artists";
                    break;
                case PLAY:
                    s = "[-help] [[exact_]artist:[*]artist_filter] [[exact_]album:[*]album_filter] [[exact_]genre:[*]genre_filter] [indexes:id1,id2] - plays matching songs";
                    break;
                case QUEUE:
                    s = "[-help] [[exact_]artist:[*]artist_filter] [[exact_]album:[*]album_filter] [[exact_]genre:[*]genre_filter] [indexes:id1,id2] - adds matching songs to the now playing list";
                    break;
                case SERV_COVER:
                    s = "[size-x:<width>] [size-y:<height>] [[exact_]artist:[*]artist_filter] [[exact_]album:[*]album_filter] [[exact_]genre:[*]genre_filter] [indexes:id1,id2] - serves the album cover of the first match";
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
            or.AppendFormat("music-play [~filters~] [~index-list~] - plays all matching songs");
            or.AppendFormat("music-queue [~filters~] [~index-list~] - queues all matching songs");
            or.AppendFormat("music-cover [~filters~] [~index-list~] [size-x:width] [size-y:height] - serves the cover image (first match)");
            or.AppendFormat(" ");
            or.AppendFormat("Where:");
            or.AppendFormat("     [~filters~] is one or more of: [~artist-filter~] [~album-filter~] [~genre-filter~] ");
            or.AppendFormat("     [~artist-filter~] is one of:");
            or.AppendFormat("          artist:<text> - matches track artists that start with <text> (\"artist:ab\" would match artists \"ABBA\" and \"ABC\")");
            or.AppendFormat("          artist:*<text> - matches track artists that have any words that start with <text> (\"artist:*ab\" would match \"ABBA\" and \"The Abstracts\")");
            or.AppendFormat("          exact-artist:<text> - matches the track artist that exactly matches <text> (\"exact-artist:ab\" would only match an artist names \"Ab\")");
            or.AppendFormat("     [~album-filter~] is one of:");
            or.AppendFormat("          album:<text> - matches albums that start with <text> (\"album:ab\" would match the album \"ABBA Gold\" and \"Abbey Road\")");
            or.AppendFormat("          exact_album:<text> - matches the album exactly named <text> (\"exact_album:ab\" would only match an album named \"Ab\")");
            or.AppendFormat("     [~genre-filter~] is one of:");
            or.AppendFormat("          genre:<text> - matches genre that start with <text> (\"genre:ja\" would match the genre \"Jazz\")");
            or.AppendFormat("          genre:*<text> - matches genres that contain <text> (\"genre:*rock\" would match \"Rock\" and \"Alternative Rock\")");
            or.AppendFormat("          exact_genre:<text> - matches the genere exactly named <text> (\"exact_genre:ja\" would only match an genre named \"Ja\")");
            or.AppendFormat("     [~index-list~] is of the form:");
            or.AppendFormat("          indexes:idx1,idx2... - specifies one or more specific songs returned by the filter");
            or.AppendFormat("               Where idx1,idx2... is a comma separated list with no spaces (e.g. 'indexes:22,23,27')");
            or.AppendFormat("     [~custom-template~] is of the form:");
            or.AppendFormat("          template:<name> - specifies a custom template <name> defined in the \"music.template\" file");
            or.AppendFormat("     [size-x:~width~] - Resizes the served image, where ~width~ is the max width of the served image");
            or.AppendFormat("     [size-y:~height~] - Resizes the served image, where ~height~ is the max height of the served image");
            or.AppendFormat(" ");
            or.AppendFormat("Parameter Notes:");
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
            or.AppendFormat("     music-play exact_album:\"abbey road\" indexes:1,3 - would play the second and third songs (indexes are zero based) returned by the search for an album named \"Abbey Road\"");
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

            public void add_song_to_album(WMPLib.IWMPMedia media_item)
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

                string s = media_item.getItemInfo("WM/AlbumTitle");
                if (album != s) nextAlbum = s;

                s = media_item.getItemInfo("WM/AlbumArtist");
                if (s == "") s = media_item.getItemInfo("Author");
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
            s = do_conditional_replace(s, "songPath", the_state.songLocation);
            s = do_conditional_replace(s, "songTrackNumber", the_state.songTrackNumber);
            s = do_conditional_replace(s, "songLength", the_state.songLength);
            s = s.Replace("%index%", String.Format("{0}", index));

            s = do_conditional_replace(s, "trackCount", String.Format("{0}", the_state.trackCount));
            s = do_conditional_replace(s, "albumCount", String.Format("{0}", the_state.albumCount));

            return s;
        }

        private OpResult do_basic_list(OpResult or, WMPLib.IWMPStringCollection list, string list_type, string template, string default_template)
        {
            int result_count = list.count;
            or.AppendFormat("{0}", basic_replacer(getTemplate(template + ".H", ""), "", "", result_count, -1)); // Header
            string c_last = " ";
            for (int j = 0; j < list.count; j++)
            {
                string list_item = list.Item(j);
                if (list_item.Length > 0)
                {
                    string s_out = "";
                    if (first_letter(list_item) != c_last)
                    {
                        if (j > 0)
                        {
                            s_out = getTemplate(template + ".Alpha-", "");
                            s_out = basic_replacer(s_out, list_type, list.Item(j - 1), result_count, j);
                            s_out = basic_replacer(s_out, "letter", c_last, result_count, j);
                            if (s_out.Length > 0) or.AppendFormat("{0}", s_out); // Alpha-
                        }
                        c_last = first_letter(list_item);
                        s_out = getTemplate(template + ".Alpha+", "");
                        s_out = basic_replacer(s_out, list_type, list_item, result_count, j);
                        s_out = basic_replacer(s_out, "letter", c_last, result_count, j);
                        if (s_out.Length > 0) or.AppendFormat("{0}", s_out); // Alpha +
                    }
                    s_out = getTemplate(template + ".Entry", default_template);
                    s_out = basic_replacer(s_out, list_type, list_item, result_count, j);
                    s_out = basic_replacer(s_out, "letter", c_last, result_count, j);
                    s_out = do_conditional_replace(s_out, "all_filters", request_params);
                    s_out = do_conditional_replace(s_out, "genreFilter", genre_filter);
                    s_out = do_conditional_replace(s_out, "artistFilter", artist_filter);
                    s_out = do_conditional_replace(s_out, "albumFilter", album_filter);

                    //opResult.AppendFormat("artist={0}", artists.Item(j));
                    if (s_out.Length > 0) or.AppendFormat("{0}", s_out); // Entry
                }
            }
            if (result_count > 0) // Close the final alpha grouping
            {
                string s_out = getTemplate(template + ".Alpha-", "");
                s_out = basic_replacer(s_out, list_type, list.Item(result_count - 1), result_count, result_count);
                s_out = basic_replacer(s_out, "letter", c_last, result_count, result_count);
                if (s_out.Length > 0) or.AppendFormat("{0}", s_out); // Alpha-
            }
            return or;
        }

        private OpResult do_detailed_list(OpResult or, WMPLib.IWMPMedia media_item, int idx, string template)
        {

            string artist = "";
            string album = "";
            string letter = "";
            bool added = false;

            int index = idx;
            if (index < 0) 
                index = the_state.trackCount;

            if (media_item != null)
            {
                artist = media_item.getItemInfo("WM/AlbumArtist");
                if (artist == "") artist = media_item.getItemInfo("Author");
                album = media_item.getItemInfo("WM/AlbumTitle");
                letter = first_letter(artist);
            }

            // End of artist?
            if (artist != the_state.artist)
            {
                if (the_state.album.Length > 0)
                {
                    var s_out = replacer(getTemplate(template + ".Album-", DEFAULT_DETAIL_ALBUM_END), index);
                    if (s_out.Length > 0) or.AppendFormat("{0}", s_out);
                }
                if (the_state.artist.Length > 0)
                {
                    var s_out = replacer(getTemplate(template + ".Artist-", DEFAULT_DETAIL_ARTIST_END), index);
                    if (s_out.Length > 0) or.AppendFormat("{0}", s_out);
                }
                // End of current aplha?
                if (letter != the_state.letter && the_state.letter.Length == 1)
                {
                    var s_out = replacer(getTemplate(template + ".Alpha-", ""), index);
                    if (s_out.Length > 0) or.AppendFormat("{0}", s_out);
                }

                if (index >= 0)
                {
                    the_state.resetArtist(artist);
                    the_state.resetAlbum(album);
                }

                if (media_item != null)
                {
                    the_state.add_song_to_album(media_item);
                    added = true;
                    // Start new aplha?
                    if (letter != the_state.letter)
                    {
                        the_state.letter = letter;
                        var s_out = replacer(getTemplate(template + ".Alpha+", ""), index);
                        if (s_out.Length > 0) or.AppendFormat("{0}", s_out);
                    }
                    // Start new artist
                    if (the_state.artist.Length > 0)
                    {
                        var s_out = replacer(getTemplate(template + ".Artist+", DEFAULT_DETAIL_ARTIST_START), index);
                        if (s_out.Length > 0) or.AppendFormat("{0}", s_out);
                    }
                    // Start new album
                    if (the_state.album.Length > 0)
                    {
                        var s_out = replacer(getTemplate(template + ".Album+", DEFAULT_DETAIL_ALBUM_START), index);
                        if (s_out.Length > 0) or.AppendFormat("{0}", s_out);
                    }
                }
            }
            // End of album?
            else if (album != the_state.album)
            {
                if (the_state.album.Length > 0)
                {
                    var s_out = replacer(getTemplate(template + ".Album-", DEFAULT_DETAIL_ALBUM_END), index);
                    if (s_out.Length > 0) or.AppendFormat("{0}", s_out);
                }
                if (index >= 0) the_state.resetAlbum(album);
                if (media_item != null)
                {
                    the_state.add_song_to_album(media_item);
                    added = true;
                    if (the_state.album.Length > 0)
                    {
                        var s_out = replacer(getTemplate(template + ".Album+", DEFAULT_DETAIL_ALBUM_START), index);
                        if (s_out.Length > 0) or.AppendFormat("{0}", s_out);
                    }
                }
            }

            // Do track:
            if (media_item != null)
            {
                if (!added) the_state.add_song_to_album(media_item);
                var s_out = replacer(getTemplate(template + ".Entry", DEFAULT_DETAIL_SONG), index);
                if (s_out.Length > 0) or.AppendFormat("{0}", s_out);
            }

            return or;
        }

        public string trim_parameter(string param)
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
                if (!Directory.Exists(CACHE_DIR)) Directory.CreateDirectory(CACHE_DIR);
                if (!Directory.Exists(CACHE_MUSIC_CMD_DIR)) Directory.CreateDirectory(CACHE_MUSIC_CMD_DIR);

                FileInfo fi = new FileInfo(CACHE_VER_FILE);
                if (!fi.Exists)
                {
                    System.IO.File.WriteAllText(CACHE_VER_FILE, cur_ver);
                }

                System.IO.File.WriteAllText(cached_file, content);
            }
            catch (Exception) { return; }

            return;
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

            string mrp_file = CACHE_DIR + "\\mrp_list.dat";

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
            string mrp_file = CACHE_DIR + "\\mrp_list.dat";
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
                if (!Directory.Exists(CACHE_DIR)) Directory.CreateDirectory(CACHE_DIR);

                System.IO.StreamWriter file = new System.IO.StreamWriter(mrp_file);
                foreach (string s in mrp_content_final) file.WriteLine(s);
                file.Close();
            }
            catch (Exception) { return; }

            return;
        }

        private OpResult list_stats(OpResult or, string template)
        {
            or.AppendFormat("{0}", basic_replacer(getTemplate(template + ".H", ""), "", "", 0, -1)); // Header

            string s_out = getTemplate(template + ".Entry", DEFAULT_STATS);
            s_out = basic_replacer(s_out, "track_count",
                    String.Format("{0}", Player.mediaCollection.getByAttribute("MediaType", "Audio").count), 0, -1);

            WMPLib.IWMPMediaCollection2 collection = (WMPLib.IWMPMediaCollection2)Player.mediaCollection;
            WMPLib.IWMPQuery stats_query = collection.createQuery();

            s_out = basic_replacer(s_out, "artist_count",
                    String.Format("{0}", collection.getStringCollectionByQuery("Artist", stats_query, "Audio", "Artist", true).count), 0, -1);
            s_out = basic_replacer(s_out, "album_count",
                    String.Format("{0}", collection.getStringCollectionByQuery("WM/AlbumTitle", stats_query, "Audio", "WM/AlbumTitle", true).count), 0, -1);
            s_out = basic_replacer(s_out, "genre_count",
                    String.Format("{0}", collection.getStringCollectionByQuery("Genre", stats_query, "Audio", "Genre", true).count), 0, -1);

            if (s_out.IndexOf("%cache_age%") >= 0)
            {
                string cache_age = "No cache";
                try
                {
                    cache_age = System.IO.File.GetCreationTime(CACHE_VER_FILE).ToString();
                }
                catch (Exception) { ; }
                s_out = basic_replacer(s_out, "cache_age", cache_age, 0, -1);
            }

            if (s_out.IndexOf("%available_templates%") >= 0)
            {
                string template_list = "";
                foreach (KeyValuePair<string, string> t in m_templates)
                {
                    if (t.Key.EndsWith(".Entry"))
                    {
                        if (template_list.Length > 0) template_list += ", ";
                        template_list += t.Key.Substring(0, t.Key.Length - ".Entry".Length);
                    }
                }
                s_out = basic_replacer(s_out, "available_templates", template_list, 0, -1);
            }
            if (s_out.Length > 0) or.AppendFormat("{0}", s_out); // Entry

            or.AppendFormat("{0}", basic_replacer(getTemplate(template + ".F", ""), "", "", 0, -1)); // Footer

            return or;
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

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult Execute(string param)
        {
            debug_last_action = "Execute: Start";

            DateTime startTime = DateTime.Now;

            OpResult opResult = new OpResult();
            opResult.StatusCode = OpStatusCode.Ok;

            bool bFirst = false;
            int size_x = 0;
            int size_y = 0;

            string template = "";
            string cache_fn = make_cache_fn(String.Format("{0}-{1}.txt", which_command, param));
            string cache_body = "";
            bool is_cached = false;


            try
            {
                if (param.IndexOf("-help") >= 0)
                {
                    opResult = showHelp(opResult);
                    return opResult;
                }
                //if (Player == null) Player = new WMPLib.WindowsMediaPlayer();

                WMPLib.IWMPMediaCollection2 collection = (WMPLib.IWMPMediaCollection2)Player.mediaCollection;
                int ver = Player.mediaCollection.getByAttribute("MediaType", "Audio").count;//.GetHashCode();
                string cache_ver = String.Format("{0}", ver);
                WMPLib.IWMPQuery query = collection.createQuery();
                WMPLib.IWMPPlaylist mediaPlaylist = null;
                WMPLib.IWMPMedia media_item;

                ArrayList a_idx = new ArrayList();

                bool b_query = false;

                string recent_text = "";
                string recent_text_type = "";

                debug_last_action = "Execution: Parsing params";

                request_params = HttpUtility.UrlEncode(param);

                if (param.Contains("exact-genre:"))
                {
                    string genre = param.Substring(param.IndexOf("exact-genre:") + "exact-genre:".Length);
                    genre = trim_parameter(genre);
                    query.addCondition("Genre", "Equals", genre);
                    genre_filter = genre;
                    recent_text = genre;
                    recent_text_type = "Genre";
                    b_query = true;
                }
                else if (param.Contains("genre:*"))
                {
                    string genre = param.Substring(param.IndexOf("genre:*") + "genre:*".Length);
                    genre = trim_parameter(genre);
                    query.addCondition("Genre", "BeginsWith", genre);
                    query.beginNextGroup();
                    query.addCondition("Genre", "Contains", " " + genre);
                    genre_filter = genre;
                    recent_text = genre;
                    recent_text_type = "Genre";
                    b_query = true;
                }
                else if (param.Contains("genre:"))
                {
                    string genre = param.Substring(param.IndexOf("genre:") + "genre:".Length);
                    genre = trim_parameter(genre);
                    query.addCondition("Genre", "BeginsWith", genre);
                    genre_filter = genre;
                    recent_text = genre;
                    recent_text_type = "Genre";
                    b_query = true;
                }

                if (param.Contains("exact-artist:"))
                {
                    string artist = param.Substring(param.IndexOf("exact-artist:") + "exact-artist:".Length);
                    artist = trim_parameter(artist);
                    query.addCondition("Artist", "Equals", artist);
                    artist_filter = artist;
                    if (recent_text.Length > 0)
                    {
                        recent_text += ": ";
                        recent_text_type += "/";
                    }
                    recent_text += artist;
                    recent_text_type += "Artist";
                    b_query = true;
                }
                else if (param.Contains("artist:*"))
                {
                    string artist = param.Substring(param.IndexOf("artist:*") + "artist:*".Length);
                    artist = trim_parameter(artist);
                    query.addCondition("Artist", "BeginsWith", artist);
                    query.beginNextGroup();
                    query.addCondition("Artist", "Contains", " " + artist);
                    artist_filter = artist;
                    if (recent_text.Length > 0)
                    {
                        recent_text += ": ";
                        recent_text_type += "/";
                    }
                    recent_text += artist;
                    recent_text_type += "Artist";
                    b_query = true;
                }
                else if (param.Contains("artist:"))
                {
                    string artist = param.Substring(param.IndexOf("artist:") + "artist:".Length);
                    artist = trim_parameter(artist);
                    query.addCondition("Artist", "BeginsWith", artist);
                    artist_filter = artist;
                    if (recent_text.Length > 0)
                    {
                        recent_text += ": ";
                        recent_text_type += "/";
                    }
                    recent_text += artist;
                    recent_text_type += "Artist";
                    b_query = true;
                }

                if (param.Contains("exact-album:"))
                {
                    string album = param.Substring(param.IndexOf("exact-album:") + "exact-album:".Length);
                    album = trim_parameter(album);
                    query.addCondition("WM/AlbumTitle", "Equals", album);
                    album_filter = album;
                    if (recent_text.Length > 0)
                    {
                        recent_text += ": ";
                        recent_text_type += "/";
                    }
                    recent_text += album;
                    recent_text_type += "Album";
                    b_query = true;
                }
                //else if (param.Contains("album:*"))
                //{
                //    string artist = param.Substring(param.IndexOf("album:*") + "album:*".Length);
                //    if (album.IndexOf(" ") >= 0) album = album.Substring(0, album.IndexOf(" "));
                //    query.addCondition("WM/AlbumTitle", "BeginsWith", album);
                //    query.beginNextGroup();
                //    query.addCondition("WM/AlbumTitle", "Contains", " " + album);
                //}
                else if (param.Contains("album:"))
                {
                    string album = param.Substring(param.IndexOf("album:") + "album:".Length);
                    album = trim_parameter(album);
                    query.addCondition("WM/AlbumTitle", "BeginsWith", album);
                    album_filter = album;
                    if (recent_text.Length > 0)
                    {
                        recent_text += ": ";
                        recent_text_type += "/";
                    }
                    recent_text += album;
                    recent_text_type += "Album";
                    b_query = true;
                }

                // Indexes specified?
                if (param.Contains("indexes:"))
                {
                    string indexes = param.Substring(param.IndexOf("indexes:") + "indexes:".Length);
                    if (indexes.IndexOf(" ") >= 0) indexes = indexes.Substring(0, indexes.IndexOf(" "));
                    string[] s_idx = indexes.Split(',');
                    foreach (string s in s_idx)
                    {
                        if (s.Length > 0) a_idx.Add(Int16.Parse(s));
                    }
                    if (recent_text.Length > 0)
                    {
                        recent_text += ": ";
                        recent_text_type += "/";
                    }
                    recent_text_type += "Tracks";
                    b_query = true;
                }
                if (!b_query) recent_text_type = recent_text = "All";


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

                if (which_command == PLAY) bFirst = true;

                switch (which_command)
                {
                    case CLEAR_CACHE:
                        clear_cache();
                        return opResult;
                    case LIST_GENRES:
                        cache_body = get_cached(cache_fn, cache_ver);
                        if (cache_body.Length == 0)
                        {
                            WMPLib.IWMPStringCollection genres = collection.getStringCollectionByQuery("Genre", query, "Audio", "Genre", true);
                            do_basic_list(opResult, genres, "genre", template, "genre=%genre%");
                            result_count = genres.count;
                        }
                        else
                        {
                            is_cached = true;
                            opResult.ContentText = cache_body;
                        }
                        break;
                    case LIST_ARTISTS:
                        cache_body = get_cached(cache_fn, cache_ver);
                        if (cache_body.Length == 0)
                        {
                            WMPLib.IWMPStringCollection artists = collection.getStringCollectionByQuery("Artist", query, "Audio", "Artist", true);
                            do_basic_list(opResult, artists, "artist", template, "artist=%artist%");
                            result_count = artists.count;
                        }
                        else
                        {
                            is_cached = true;
                            opResult.ContentText = cache_body;
                        }
                        break;
                    case LIST_ALBUM_ARTISTS:
                        cache_body = get_cached(cache_fn, cache_ver);
                        if (cache_body.Length == 0)
                        {
                            WMPLib.IWMPStringCollection artists = collection.getStringCollectionByQuery("WM/AlbumArtist", query, "Audio", "WM/AlbumArtist", true);
                            do_basic_list(opResult, artists, "artist", template, "album_artist=%artist%");
                            result_count = artists.count;
                        }
                        else
                        {
                            is_cached = true;
                            opResult.ContentText = cache_body;
                        }
                        break;
                    case LIST_ALBUMS:
                        cache_body = get_cached(cache_fn, cache_ver);
                        if (cache_body.Length == 0)
                        {
                            WMPLib.IWMPStringCollection albums = collection.getStringCollectionByQuery("WM/AlbumTitle", query, "Audio", "WM/AlbumTitle", true);
                            do_basic_list(opResult, albums, "album", template, "album=%album%");
                            result_count = albums.count;
                        }
                        else
                        {
                            is_cached = true;
                            opResult.ContentText = cache_body;
                        }
                        break;
                    case LIST_SONGS:
                        cache_body = get_cached(cache_fn, cache_ver);
                        if (cache_body.Length == 0)
                        {
                            WMPLib.IWMPStringCollection songs = collection.getStringCollectionByQuery("Title", query, "Audio", "Title", true);
                            do_basic_list(opResult, songs, "song", template, "index=%index%, song=%song%");
                            result_count = songs.count;
                        }
                        else
                        {
                            is_cached = true;
                            opResult.ContentText = cache_body;
                        }
                        break;
                    case LIST_RECENT:
                        if (param.Contains("count:"))
                        {
                            string scount = param.Substring(param.IndexOf("count:") + "count:".Length);
                            if (scount.IndexOf(" ") >= 0) scount = scount.Substring(0, scount.IndexOf(" "));
                            int count = Convert.ToInt32(scount);
                            list_recent(opResult, template, count);
                        }
                        else list_recent(opResult, template);
                        return opResult;
                    case LIST_STATS:
                        cache_body = get_cached(cache_fn, cache_ver);
                        if (cache_body.Length == 0)
                        {
                            list_stats(opResult, template);
                        }
                        else
                        {
                            is_cached = true;
                            opResult.ContentText = cache_body;
                        }
                        break;
                    case LIST_DETAILS:
                    case PLAY:
                    case QUEUE:
                    case SERV_COVER:
                        if (which_command == LIST_DETAILS)
                        {
                            cache_body = get_cached(cache_fn, cache_ver);
                            if (cache_body.Length > 0)
                            {
                                is_cached = true;
                                opResult.ContentText = cache_body;
                                break;
                            }
                        }
                        if (which_command == SERV_COVER || which_command == LIST_DETAILS) the_state.init();

                        if (b_query) mediaPlaylist = collection.getPlaylistByQuery(query, "Audio", "Artist", true);
                        else mediaPlaylist = Player.mediaCollection.getByAttribute("MediaType", "Audio");

                        if (a_idx.Count > 0) result_count = a_idx.Count;
                        else result_count = mediaPlaylist.count;

                        // Header
                        opResult.AppendFormat("{0}", basic_replacer(getTemplate(template + ".H", ""), "", "", result_count, -1));

                        if (a_idx.Count > 0)
                        {
                            result_count = 0;
                            for (int j = 0; j < a_idx.Count; j++)
                            {
                                try { media_item = mediaPlaylist.get_Item((Int16)a_idx[j]); }
                                catch (Exception) { media_item = null; }
                                if (media_item != null)
                                {
                                    result_count++;
                                    if (which_command == LIST_DETAILS || which_command == SERV_COVER) // Display it
                                    {
                                        do_detailed_list(opResult, media_item, (Int16)a_idx[j], template);
                                        if (which_command == SERV_COVER)
                                        {
                                            photoCmd pc;
                                            pc = new photoCmd(photoCmd.SERV_PHOTO);
                                            if (the_state.albumImage.Length == 0)
                                                return pc.getPhoto(DEFAULT_IMAGE, "jpeg", size_x, size_y);
                                            else return pc.getPhoto(the_state.albumImage, "jpeg", size_x, size_y);
                                        }
                                    }
                                    else // Play / Queue it
                                    {
                                        recent_text += ((j == 0) ? "" : ", ") + (Int16)a_idx[j] + ". " + media_item.getItemInfo("Title");
                                        PlayMediaCmd pmc;
                                        pmc = new PlayMediaCmd(MediaType.Audio, !bFirst);
                                        bFirst = false;
                                        opResult = pmc.Execute(media_item.sourceURL);
                                    }
                                }
                            }
                        }
                        else
                        {
                            for (int j = 0; j < mediaPlaylist.count; j++)
                            {
                                media_item = mediaPlaylist.get_Item(j);
                                if (which_command == LIST_DETAILS || which_command == SERV_COVER) // Display it
                                {
                                    do_detailed_list(opResult, media_item, j, template);
                                    if (which_command == SERV_COVER)
                                    {
                                        photoCmd pc;
                                        pc = new photoCmd(photoCmd.SERV_PHOTO);
                                        if (the_state.albumImage.Length == 0)
                                            return pc.getPhoto(DEFAULT_IMAGE, "jpeg", size_x, size_y);
                                        else return pc.getPhoto(the_state.albumImage, "jpeg", size_x, size_y);
                                    }
                                }
                                else // Play / Queue it
                                {
                                    PlayMediaCmd pmc;
                                    pmc = new PlayMediaCmd(MediaType.Audio, !bFirst);
                                    bFirst = false;
                                    opResult = pmc.Execute(media_item.sourceURL);
                                }
                            }
                        }
                        if (which_command == LIST_DETAILS) do_detailed_list(opResult, null, -1, template);

                        if ((which_command == PLAY || which_command == QUEUE) && result_count > 0)
                        {
                            // Type, Artist, Album, Track, param, count
                            add_to_mrp(recent_text_type, recent_text, param, result_count); //Add to recent played list
                        }
                        break;
                }

                // Footer
                if (!is_cached)
                {
                    if (which_command != LIST_DETAILS) opResult.AppendFormat("{0}", basic_replacer(getTemplate(template + ".F", "result_count=%resultCount%"), "", "", result_count, -1));
                    else opResult.AppendFormat("{0}", replacer(getTemplate(template + ".F", "result_count=%index%"), result_count));
                    //opResult.AppendFormat("result_count={0}", result_count);
                    save_to_cache(cache_fn, opResult.ToString(), cache_ver);
                }
                string sub_footer = basic_replacer(getTemplate(template + ".C", "from_cache=%wasCached%\r\nellapsed_time=%ellapsedTime%"), "wasCached", is_cached.ToString(), -1, -1);
                TimeSpan duration = DateTime.Now - startTime;
                sub_footer = basic_replacer(sub_footer, "ellapsedTime", String.Format("{0}", duration.TotalSeconds), -1, -1);
                opResult.AppendFormat("{0}", sub_footer);
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

        #endregion
    }
}
