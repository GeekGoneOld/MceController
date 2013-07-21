/*
 * This module is build on top of on J.Bradshaw's vmcController
 * Implements audio media library functions
 * 
 * Copyright (c) 2009, 2010 Anthony Jones
 * 
 * Portions copyright (c) 2007 Jonathan Bradshaw
 * 
 * This software code module is provided 'as-is', without any express or implied warranty. 
 * In no event will the authors be held liable for any damages arising from the use of this software.
 * Permission is granted to anyone to use this software for any purpose, including commercial 
 * applications, and to alter it and redistribute it freely.
 * 
 * History:
 * 2010-06-07  Anthony Jones: Implemented photo play / queue (slideshows) and renamed cache dirs, misc (random image, specify size by width / height -OR- short_side / long_side, more caching)
 * 2010-03-10  Anthony Jones: Added available templates to the stats command
 * 2010-03-04  Anthony Jones: Fixed and tweaked
 * 2009-08-24  Anthony Jones: Added Threading to generate tag cache
 * 2009-08-01  Anthony Jones: Created
 * 
 * To Do:
 * - Elapsed time is incorrectly cached! 
 * - Add "is_cached" to custom templates
 * - New command to list current slide show contents
 * - Cache specific requests (clear specific caches manually - e.g. don't clear cached sized photos just because the tag cache is stale)
 * - List cache sizes?
 * - Make slideshows extender compatible (different slide show directory for each port?)
 * 
 */

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Runtime.Serialization.Formatters.Binary;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Threading;
using WMPLib;


namespace VmcController.AddIn.Commands
{
    /// <summary>
    /// Summary description for MsgBox command.
    /// </summary>
    public class photoCmd : ICommand
    {
        private static WindowsMediaPlayer Player = null;
        private IWMPMedia media;
        private IWMPPlaylist photo_media_play_list;

        private static Dictionary<string, List<int>> keywords;
        private static int cache_ver = -1;
        private static Dictionary<string, string> m_templates = new Dictionary<string, string>();

        private bool play = false;
        private int page_limit = int.MaxValue;
        private int page_start = 0;
        private int kw_count = 0;
        private DateTime first_date = new DateTime(9999, 12, 31);
        private DateTime last_date = new DateTime(1, 1, 1);
        private string param_is_tagged = "";
        private string param_not_tagged = "";
        private string param_start_date = "";
        private string param_end_date = "";
        private int last_img_served = -1;

        private string debug_last_action = "none";

        private static Thread tag_thread = null;
        private static bool in_generate_cache = false;
        private static bool in_generate_slideshow = false;
        private static bool init_run = false;

        public const int LIST_PHOTOS = 1;
        public const int PLAY_PHOTOS = 2;
        public const int QUEUE_PHOTOS = 3;
        public const int LIST_TAGS = 4;
        public const int CLEAR_CACHE = 5;
        public const int SERV_PHOTO = 6;
        public const int LIST_CACHE = 7;
        public const int SHOW_STATS = 8;

        private const string BUILD = ";;BUILD_ID;;";
        private const string ELAPSED = ";;GENERATE_TIME;;";

        private const string CACHE_ID = ";;CACHE_ID;;";
        private const string KEYWORD = ";;KW;;";
        private const string DATE = ";;DT;;";        

        private static string CACHE_DIR = AddInModule.DATA_DIR + "\\photo_cmd_cache";
        private static string CACHE_FILE = CACHE_DIR + "\\photoCmd.cache";
        private static string CACHE_RESIZE_DIR = CACHE_DIR + "\\Resized";
        private static string CACHE_QUERY_DIR = CACHE_DIR + "\\Queries";
        private static string CACHE_PAGE_DIR = CACHE_DIR + "\\Pages";

        private static string SLIDE_SHOW_DIR = AddInModule.DATA_DIR + "\\Slideshows";


        private const string DEFAULT_RESULT = "idx=%idx%\r\n%if-title%  Title=%title%\r\n%endif%  Filename=%filename%\r\n%if-datetaken%  Date Taken=%datetaken%\r\n%endif%%if-camera%  Camera=%camera%\r\n%endif%%if-tags%  Tags=%tags%\r\n%endif%";
        private const string DEFAULT_HEAD = "photos found=%results_count%\r\n";
        private const string DEFAULT_FOOT = "elapsed seconds=%elapsed_time%";

        private const string DEFAULT_PLAY_RESULT = "idx=%idx%. %if-title%  Title=%title%\r\n%endif%  Filename=%filename%. %if-datetaken%  Date Taken=%datetaken%";
        private const string DEFAULT_PLAY_HEAD = "photos added to queue=%results_count%\r\n";
        private const string DEFAULT_PLAY_FOOT = "elapsed seconds=%elapsed_time%";

        private const string DEFAULT_REPORT_HEAD = "";
        private const string DEFAULT_REPORT_RESULT = "tag=%tag%\r\n  images_tagged=%tag_count%";
        private const string DEFAULT_REPORT_FOOT = "tags found=%results_count%\r\nelapsed seconds=%elapsed_time%";

        private const string DEFAULT_STATS_HEAD = "photos_found=%results_count%\r\nstart_date=%start_date%\r\nend_date=%end_date%\r\ntag_count=%tag_count%%if-filter_is_tagged%\r\nfilter_is_tagged=%filter_is_tagged%%endif%%if-filter_not_tagged%\r\nfilter_not_tagged=%filter_not_tagged%%endif%\r\navailable_templates=%available_templates%";
        private const string DEFAULT_STATS_RESULT = "";
        private const string DEFAULT_STATS_FOOT = "";

        public int action = 1;

        private void reset_globals()
        {
            play = false;
            page_limit = int.MaxValue;
            page_start = 0;
            kw_count = 0;
            first_date = new DateTime(9999, 12, 31);
            last_date = new DateTime(1, 1, 1);
            param_is_tagged = "";
            param_not_tagged = "";
            param_start_date = "";
            param_end_date = "";
            debug_last_action = "none";
        }

        public photoCmd(int do_what)
        {
            action = do_what;
            /* Do init on the statics */
            if (!init_run)
            {
                init_run = true;
                loadTemplate();
                if (Player == null) Player = new WindowsMediaPlayer();
                photo_media_play_list = Player.mediaCollection.getByAttribute("MediaType", "Photo");
                validate_cache();
            }
        }

        public bool create_dirs()
        {
            //Create dirs if needed:
            if (!Directory.Exists(AddInModule.DATA_DIR))
            {
                try { Directory.CreateDirectory(AddInModule.DATA_DIR); }
                catch (Exception) { return false; }
            }
            if (!Directory.Exists(SLIDE_SHOW_DIR))
            {
                try { Directory.CreateDirectory(SLIDE_SHOW_DIR); }
                catch (Exception) { return false; }
            }
            if (!Directory.Exists(CACHE_DIR))
            {
                try { Directory.CreateDirectory(CACHE_DIR); }
                catch (Exception) { return false; }
            }
            if (!Directory.Exists(CACHE_RESIZE_DIR))
            {
                try { Directory.CreateDirectory(CACHE_RESIZE_DIR); }
                catch (Exception) { return false; }
            }
            if (!Directory.Exists(CACHE_QUERY_DIR))
            {
                try { Directory.CreateDirectory(CACHE_QUERY_DIR); }
                catch (Exception) { return false; }
            }
            if (!Directory.Exists(CACHE_PAGE_DIR))
            {
                try { Directory.CreateDirectory(CACHE_PAGE_DIR); }
                catch (Exception) { return false; }
            }
            
            return true;
        }

        public void serialize()
        {
            debug_last_action = "Opening cache file for serialization: " + CACHE_FILE;

            //Create dirs if needed:
            if (!Directory.Exists(CACHE_DIR)) create_dirs();

            Stream stream = File.Open(CACHE_FILE, FileMode.Create, FileAccess.ReadWrite);
            BinaryFormatter formatter = new BinaryFormatter();

            debug_last_action = "Writing serialization to " + CACHE_FILE;
            formatter.Serialize(stream, keywords);
            stream.Close();

            debug_last_action = "Done with serialization";

        }
        public bool deserialize()
        {
            debug_last_action = "Checking if file exists: " + CACHE_FILE;
            FileInfo fi = new FileInfo(CACHE_FILE);
            if (!fi.Exists) return false;

            debug_last_action = "Opening cache file for deserialization: " + CACHE_FILE;

            Stream stream = File.Open(CACHE_FILE, FileMode.Open, FileAccess.Read);
            BinaryFormatter formatter = new BinaryFormatter();

            debug_last_action = "Reading cache file for deserialization.";
            keywords = (Dictionary<string, List<int>>)formatter.Deserialize(stream);
            stream.Close();

            debug_last_action = "Checking version of cache";
            if (keywords.ContainsKey(CACHE_ID)) cache_ver = keywords[CACHE_ID][0];
            else cache_ver = -1;
            if (keywords.ContainsKey(BUILD))
            {
                if (keywords[BUILD][0] != System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build) cache_ver = -1;
            }
            else cache_ver = -1;

            return true;
        }

        private bool loadTemplate()
        {
            bool ret = true;

            try
            {
                Regex re = new Regex("(?<lable>.+?)\t+(?<format>.*$?)");
                StreamReader fTemplate = File.OpenText("photo.template");
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

            return tmp;
        }

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            string s = "";
            switch (action)
            {
                case CLEAR_CACHE:
                    s = "[-help] - Forces the index cache to reset (otherwise only resets when the number of photos changes)";
                    break;
                case LIST_TAGS:
                    s = "[-help] [template:<name>] - Lists all the tags found in the photo collection";
                    break;
                case SHOW_STATS:
                    s = "[-help] [template:<name>] [ids:index1,index2,...] [is-tagged:tag1;tag2...] [not-tagged:tag1;tag2...] [start-date:m-d-yyyy] [end-date:m-d-yyyy] - Lists stats of the current querry (filters, counts, etc)";
                    break;
                case LIST_PHOTOS:
                    s = "[-help] [is-tagged:tag1;...] [not-tagged:tag1;...] [start-date:m-d-yyyy] [end-date:m-d-yyyy] [template:<name>] ... - Lists all the matching images found in the photo collection";
                    break;
                case SERV_PHOTO:
                    s = "[random] [size-x:width] [size-y:height] [size-x:short] [size-y:long] [ids:index] - Serves the photo to a web browser (all filters apply, first or random matching image is returned)";
                    break;
                case LIST_CACHE:
                    s = "[-help] [template:<name>] ... - Lists all the tags found in the photo collection";
                    break;
                case PLAY_PHOTOS:
                case QUEUE_PHOTOS:
                default:
                    s = "[-help] [ids:index1,index2,...] [is-tagged:tag1;tag2...] [not-tagged:tag1;tag2...] [start-date:m-d-yyyy] [end-date:m-d-yyyy] - list / play from photo collection";
                    break;
            }
            
            return s;
        }

        public OpResult showHelp(OpResult or)
        {
            or.AppendFormat("photo-list [<Parameter List>] - lists all matching photos");
            or.AppendFormat("photo-play [<Parameter List>] - plays all matching photos");
            or.AppendFormat("photo-queue [<Parameter List>] - adds all match photos to the current slideshow");
            or.AppendFormat("photo-serv [ids:<indexlist>] [size-x:width] [size-y:height] [size-long:height] [size-short:height] [random] - serves (via http) the first/random image");
            or.AppendFormat("photo-tag-list [<Parameter List>] - lists all tags from all matching photos");
            or.AppendFormat("photo-stats [<Parameter List>]- lists 'stats' from the querry (filters, matches, etc)");
            or.AppendFormat("photo-clear-cache - forces the cache to reset (otherwise only resets when the number of photos changes)");
            or.AppendFormat(" ");
            or.AppendFormat("<Parameter List>: [ids:<indexlist>] [is-tagged:<taglist>] [not-tagged:<taglist>] [start-date:m-d-yyyy] [end-date:m-d-yyyy] [page-limit:x] [page-start:x] [template:name");
            or.AppendFormat("     [ids:<indexlist>] - lists / plays / queues the specified images");
            or.AppendFormat("          <indexlist> is a list of indexes (numbers) separated by commas (no spaces)");
            or.AppendFormat("          Note that specifying image indexes overrides all filters");
            or.AppendFormat("     [is-tagged:<taglist>] - filter images by included tags - limits results to images having all specified tags");
            or.AppendFormat("     [not-tagged:<taglist>] - filter images by excluded tags - limits results to images having none of the tags");
            or.AppendFormat("          <taglist> is a list of one or more tags that are used to filter images. Multiple tags should be separated ");
            or.AppendFormat("               with a ';' - no spaces allowed (use '_' in multi-word tags)");
            or.AppendFormat("     [start-date:m-d-yyyy] - filter images by date - limit results to images taken on or after the specified date");
            or.AppendFormat("     [end-date:m-d-yyyy] - filter images by date - limit results to images taken on or before the specified date");
            or.AppendFormat("     [page-limit:<x>] - where <x> is a number - Limits results to the first <x> images (photo-list only)");
            or.AppendFormat("     [page-start:<y>] - where <y> is a number - Results start with the <y> matching image (photo-list only)");
            or.AppendFormat("     [template:<name>] - where <name> is a custom template in the \"photo.template\" file");
            or.AppendFormat("     [size-x:<width>] - Resizes the served image, where <width> is the max width of the served image");
            or.AppendFormat("     [size-y:<height>] - Resizes the served image, where <height> is the max height of the served image");
            or.AppendFormat("     [size-short:<size_in_pix>] - Resizes the served image, where <size_in_pix> is the shorter side of the served image");
            or.AppendFormat("     [size-long:<size_in_pix>] - Resizes the served image, where <size_in_pix> is the max long size of the served image");
            or.AppendFormat("     [random] - (photo-serv only) serves a random image from the results (else serves the first image)");
            or.AppendFormat(" ");
            or.AppendFormat("Filter Notes:");
            or.AppendFormat("     - Multiple filters can be used");
            or.AppendFormat("     - The 'page-limit' and 'page-start' paramters can be used to page results");
            or.AppendFormat("     - The start-date and end-date filters are based on the 'DateTaken' value in the image's metadata");
            or.AppendFormat("     - If image ids are used all other filters are ignored!");
            or.AppendFormat(" ");
            or.AppendFormat("Resizeing Notes:");
            or.AppendFormat("     - Conflicting sizing can be specified, but only the most restrictive size (resulting in smallest image) ");
            or.AppendFormat("       will be used - the photo's ratio is preserved");
            or.AppendFormat(" ");
            or.AppendFormat("Custom Formatting Notes:");
            or.AppendFormat("     All custom formats must be defined in the \"photo.template\" file");
            or.AppendFormat("     The \"photo.template\" file must be manually coppied to the ehome directory (usually C:\\Windows\\ehome)");
            or.AppendFormat("     The \"photo.template\" file contains notes / examples on formatting");

            return or;
        }

        public void validate_cache()
        {
            debug_last_action = "Validating cache...";
            if (cache_ver != photo_media_play_list.count) deserialize();
            if (cache_ver != photo_media_play_list.count)
            {
                clear_cache();
                launch_generate_tags_list();
            }
            return;
        }
        public OpResult launch_generate_tags_list()
        {
            OpResult op_return = new OpResult(OpStatusCode.Ok);

            if (in_generate_cache)
            {
                if (keywords.ContainsKey(ELAPSED))
                    op_return.AppendFormat("Already generating cache, last time this operation took {0} seconds.", keywords[ELAPSED][0]);
                else op_return.AppendFormat("Already generating cache, First time generating - be patient.");
                return op_return;
            }
            else
            {
                tag_thread = new Thread(new ThreadStart(generate_tags_list));
                if (keywords != null && keywords.ContainsKey(ELAPSED))
                    op_return.AppendFormat("Generating cache, last time this operation took {0} seconds.", keywords[ELAPSED][0]);
                else
                {
                    keywords = new Dictionary<string, List<int>>();
                    string tmp_keyword = KEYWORD + "--Please_be_patient_generating_tags_list--";
                        keywords.Add(tmp_keyword, new List<int>());
                    op_return.AppendFormat("Generating cache, First time generating - be patient.");
                }
                tag_thread.Start();
            }

            return op_return;
        }


        public void generate_tags_list()
        {
            return;
        }

        public OpResult list_tags(OpResult opResult, string template)
        {
            //validate_cache();
            return list_tags(opResult, template, keywords);
        }
        public OpResult list_tags(OpResult opResult, string template, Dictionary<string, List<int>> kw_dic)
        {
            debug_last_action = "List tags: Start";
            //Sort keywords:
            List<string> keyword_list = new List<string>(kw_dic.Keys);
            keyword_list.Sort();
            kw_count = 0;
            string output_template = getTemplate(template, DEFAULT_REPORT_RESULT);
            output_template = file_includer(output_template);

            DateTime tmp_date = new DateTime();

            // Header (note no substitutions)
            string s_head = getTemplate(template + "+", DEFAULT_REPORT_HEAD);
            s_head = file_includer(s_head);
            opResult.AppendFormat("{0}", s_head);

            foreach (string s in keyword_list)
            {
                if (s.StartsWith(KEYWORD))
                {
                    string kw = s.Substring(KEYWORD.Length);
                    string s_out = do_conditional_replace(output_template, "tag_count", String.Format("{0}", kw_dic[s].Count));
                    s_out = do_conditional_replace(s_out, "tag", kw);
                    opResult.AppendFormat("{0}", s_out);
                    kw_count++;
                }
                else if (s.StartsWith(DATE))
                {
                    string date = s.Substring(DATE.Length);
                    string[] date_elements = date.Split('-');
                    if (date_elements.Length >= 3) tmp_date = new DateTime(Int32.Parse(date_elements[2]),
                            Int32.Parse(date_elements[0]), Int32.Parse(date_elements[1]));
                    if (tmp_date < first_date) first_date = tmp_date;
                    if (tmp_date > last_date) last_date = tmp_date;
                }
            }
            debug_last_action = "List tags: Finished foreach loop";

            debug_last_action = "List tags: Done";

            return opResult;
        }

        public OpResult show_stats(string template, HashSet<int> results)
        {
            OpResult op_return = new OpResult(OpStatusCode.Ok);
            string template_list = "";
            int pic_count = photo_media_play_list.count;

            if (results.Count > 0) pic_count = results.Count;

            string output_template = getTemplate(template + "+", DEFAULT_STATS_HEAD);
            output_template += getTemplate(template, DEFAULT_STATS_RESULT);
            output_template += getTemplate(template + "-", DEFAULT_STATS_FOOT);

            if (output_template.Length > 0)
            {
                output_template = file_includer(output_template);
                output_template = do_conditional_replace(output_template, "results_count", String.Format("{0}", pic_count));
                output_template = do_conditional_replace(output_template, "tag_count", String.Format("{0}", kw_count));
                output_template = do_conditional_replace(output_template, "start_date", String.Format("{0}", first_date.ToShortDateString()));
                output_template = do_conditional_replace(output_template, "end_date", String.Format("{0}", last_date.ToShortDateString()));
                output_template = do_conditional_replace(output_template, "filter_is_tagged", String.Format("{0}", param_is_tagged));
                output_template = do_conditional_replace(output_template, "filter_not_tagged", String.Format("{0}", param_not_tagged));
                output_template = do_conditional_replace(output_template, "filter_start_date", String.Format("{0}", param_start_date));
                output_template = do_conditional_replace(output_template, "filter_end_date", String.Format("{0}", param_end_date));
                if (output_template.IndexOf("%available_templates%") >= 0)
                {
                    foreach (KeyValuePair<string, string> t in m_templates)
                    {
                        if (!t.Key.EndsWith("+") && !t.Key.EndsWith("-"))
                        {
                            if (template_list.Length > 0) template_list += ", ";
                            template_list += t.Key;
                        }
                    }
                }
                /* get random image index in results if needed */
                if (output_template.Contains("%random_id%"))
                {
                    int random_index = new Random().Next(pic_count);
                    string random = "";

                    if (results.Count > 0)
                    {
                        HashSet<int>.Enumerator results_enum = results.GetEnumerator();
                        for (int i = 0; i < random_index; i++) results_enum.MoveNext();
                        random = "" + results_enum.Current;
                    }
                    else random = "" + random_index;
                    output_template = do_conditional_replace(output_template, "random_id", random);
                }
                output_template = do_conditional_replace(output_template, "available_templates", template_list);
                op_return.AppendFormat("{0}", output_template);
            }

            return op_return;
        }

        public HashSet<int> filter_is_tagged(HashSet<int> results, List<string> tags)
        {
            debug_last_action = "Filter is_tagged: Start";
            //HashSet<int> results = null;
            foreach (string s in tags)
            {
                if (s.Length > 0)
                {
                    string kw = KEYWORD + s;
                    if (!keywords.ContainsKey(kw))
                    {
                        debug_last_action = "Filter is_tagged: Clearing results - No keyword in dictionary: " + s;
                        results.Clear();
                        break;
                    }
                    else results.IntersectWith(keywords[kw]);
                }
                debug_last_action = "Filter is_tagged: Applied keyword: " + s;
            }
            debug_last_action = "Filter is_tagged: End";

            return results;
        }

        public HashSet<int> filter_not_tagged(HashSet<int> results, List<string> tags)
        {
            debug_last_action = "Filter not_tagged: Start";
            //HashSet<int> results = null;
            foreach (string s in tags)
            {
                if (s.Length > 0)
                {
                    string kw = KEYWORD + s;
                    if (keywords.ContainsKey(kw))
                    {
                        results.ExceptWith(keywords[kw]);
                    }
                }
                debug_last_action = "Filter not_tagged: Applied keyword: " + s;
            }
            debug_last_action = "Filter not_tagged: End";

            return results;
        }

        public string do_conditional_replace (string s, string item, string v)
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

        public string file_includer (string s)
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

        public string replacer(string s, WMPLib.IWMPMedia photo, int idx)
        {
            return s;
        }

        private int parse_int(string s)
        {
            int ret_val = 0;
            int idx = 0;

            while (idx < s.Length && s[idx] <= '9' && s[idx] >= '0')
                ret_val = (10 * ret_val) + (s[idx++] - '0');

            return ret_val;
        }

        private HashSet<int> parse_ids(string s)
        {
            HashSet<int> ret_set = new HashSet<int>();
            string s_id = s.Split(' ')[0];
            string[] id_list = s.Split(',');

            foreach (string id in id_list) ret_set.Add(parse_int(id));

            return ret_set;
        }

        public OpResult getPhoto(WMPLib.IWMPMedia photo, int r_w, int r_h)
        {
            return getPhoto(photo, r_w, r_h, 0, 0);
        }
        public OpResult getPhoto(WMPLib.IWMPMedia photo, int r_w, int r_h, int r_s, int r_l)
        {
            string filename = photo.getItemInfo("SourceURL");
            string filetype = "";
            try
            {
                filetype = photo.getItemInfo("FileType");
            }
            catch (Exception)
            {
                filetype = "jpeg";
            }
            return getPhoto(filename, filetype, r_w, r_h, r_s, r_l);
        }
        public OpResult getPhoto(string filename, string filetype, int r_w, int r_h)
        {
            return getPhoto(filename, filetype, r_w, r_h, 0, 0);
        }
        public OpResult getPhoto(string filename, string filetype, int r_w, int r_h, int r_s, int r_l)
        {
            OpResult opResult = new OpResult(OpStatusCode.OkImage);
            if (r_w > 0 || r_h > 0 || r_s > 0 || r_l > 0) filename = resize(filename, r_w, r_h, r_s, r_l);
            FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read);

            // Create a reader that can read bytes from the FileStream.
            BinaryReader reader = new BinaryReader(fs);
            byte[] bytes = new byte[fs.Length];
            int read;
            string sResponse = "";
            while ((read = reader.Read(bytes, 0, bytes.Length)) != 0)
            {
                // Read from the file and write the data to the network
                sResponse = sResponse + Encoding.ASCII.GetString(bytes, 0, read);
            }
            reader.Close();
            opResult.ImageData = bytes;
            return opResult;
        }

        private static Image resizeImage(Image original, int w, int h, int s, int l)
        {
            int original_w = original.Width;
            int original_h = original.Height;
            int new_w = w;
            int new_h = h;

            float percent = 100;
            float percent_w = 100;
            float percent_h = 100;

            // Figure out orientation for short size / long side measurements;
            if (original_w > original_h) // horizontal
            {
                if (w > 0 && l > 0) new_w = (w < l) ? w : l;
                else if (l > 0) new_w = l;
                if (h > 0 && s > 0) new_h = (h < s) ? h : s;
                else if (s > 0) new_h = s;
            }
            else // vertical or square
            {
                if (w > 0 && s > 0) new_w = (w < s) ? w : s;
                else if (s > 0) new_w = s;
                if (h > 0 && l > 0) new_h = (h < l) ? h : l;
                else if (l > 0) new_h = l;
            }

            if (new_w > 0) percent_w = ((float)new_w / (float)original_w);
            if (new_h > 0) percent_h = ((float)new_h / (float)original_h);

            percent = (percent_h < percent_w) ? percent_h : percent_w;

            int destWidth = (int)(original_w * percent);
            int destHeight = (int)(original_h * percent);

            Bitmap b = new Bitmap(destWidth, destHeight);
            Graphics g = Graphics.FromImage((Image)b);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            g.DrawImage(original, 0, 0, destWidth, destHeight);
            g.Dispose();

            return (Image)b;
        }

        private ImageCodecInfo getEncoderInfo(string mimeType)
        {
            // Get image codecs for all image formats
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();

            // Find the correct image codec
            for (int i = 0; i < codecs.Length; i++)
                if (codecs[i].MimeType == mimeType)
                    return codecs[i];
            return null;
        }
        private void saveJpeg(string path, Bitmap img, long quality)
        {
            // Encoder parameter for image quality
            EncoderParameter qualityParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);

            // Jpeg image codec
            ImageCodecInfo jpegCodec = this.getEncoderInfo("image/jpeg");

            if (jpegCodec == null)
                return;

            EncoderParameters encoderParams = new EncoderParameters(1);
            encoderParams.Param[0] = qualityParam;

            img.Save(path, jpegCodec, encoderParams);
        }

        public string resize(string fn, int max_w, int max_h)
        {
            return resize(fn, max_w, max_h, 0, 0);
        }

        public string resize(string fn, int max_w, int max_h, int max_s, int max_l)
        {
            string cached_file = fn;

            try
            {
                //Create dir if needed:
                if (!Directory.Exists(CACHE_RESIZE_DIR)) if(!create_dirs()) return fn;

                cached_file = cached_file.Replace("\\", "_");
                cached_file = cached_file.Replace(":", "_") + String.Format("_{0}_{1}_{2}_{3}.jpg", max_w, max_h, max_s, max_l);

                cached_file = CACHE_RESIZE_DIR + "\\" + cached_file;

                FileInfo fi = new FileInfo(cached_file);
                if (!fi.Exists)
                {
                    Image img = Image.FromFile(fn);
                    img = resizeImage(img, max_w, max_h, max_s, max_l);
                    saveJpeg(cached_file, new Bitmap(img), 85L);
                }
            }
            catch (Exception) { cached_file = fn; }
            return cached_file;
        }

        public void serialize_rendered_page(OpResult or, string param)
        {
            string fn = CACHE_PAGE_DIR + "\\" + make_cache_fn("p_" + param, 100);
            debug_last_action = "Opening page cache file for writing: " + fn;

            //Create dirs if needed:
            if (!Directory.Exists(CACHE_PAGE_DIR)) if (!create_dirs()) return;

            StreamWriter writer = new StreamWriter(fn, false);

            debug_last_action = "Writing page to " + fn;
            writer.Write(or.ToString());
            writer.Close();

            debug_last_action = "Done with serialization";
        }
        public OpResult deserialize_rendered_page(OpResult or, string param)
        {
            string fn = CACHE_PAGE_DIR + "\\" + make_cache_fn("p_" + param, 100);
            debug_last_action = "Checking if file exists: " + fn;
            FileInfo fi = new FileInfo(fn);
            if (!fi.Exists) return null;

            debug_last_action = "Opening cache file for reading: " + fn;

            StreamReader reader = new StreamReader(fn);

            debug_last_action = "Reading cache file.";
            or = new OpResult(reader.ReadToEnd());
            reader.Close();
            return or;
        }

        public void serialize_results(HashSet<int> results, string param)
        {
            string fn = CACHE_QUERY_DIR + "\\" + make_cache_fn("q_" + param, 100);
            debug_last_action = "Opening cache file for serialization: " + fn;

            //Create dirs if needed:
            if (!Directory.Exists(CACHE_QUERY_DIR)) if (!create_dirs()) return;

            Stream stream = File.Open(fn, FileMode.Create, FileAccess.ReadWrite);
            BinaryFormatter formatter = new BinaryFormatter();

            debug_last_action = "Writing serialization to " + fn;
            formatter.Serialize(stream, results);
            stream.Close();

            debug_last_action = "Done with serialization";

        }
        public HashSet<int> deserialize_results(HashSet<int> results, string param)
        {
            string fn = CACHE_QUERY_DIR + "\\" + make_cache_fn("q_" + param, 100);
            debug_last_action = "Checking if file exists: " + fn;
            FileInfo fi = new FileInfo(fn);
            if (!fi.Exists) return null;

            debug_last_action = "Opening cache file for deserialization: " + fn;

            Stream stream = File.Open(fn, FileMode.Open, FileAccess.Read);
            BinaryFormatter formatter = new BinaryFormatter();

            debug_last_action = "Reading cache file for deserialization.";
            results = (HashSet<int>)formatter.Deserialize(stream);
            stream.Close();

            return results;
        }

        public HashSet<int> get_list(HashSet<int> results, string param)
        {
            HashSet<int> ret = new HashSet<int>();

            // Check caches results:
            ret = deserialize_results(results, param);
            if (ret != null) return ret;

            // Date Limited?
            if (param.Contains("start-date:") || param.Contains("end-date:"))
            {
                DateTime start_date = new DateTime(1, 1, 1);
                DateTime end_date = new DateTime(9999, 12, 31);
                DateTime tmp_date = new DateTime();
                if (param.Contains("start-date:"))
                {
                    string date_String = param.Substring(param.IndexOf("start-date:") + "start-date:".Length);
                    if (date_String.IndexOf(" ") >= 0) date_String = date_String.Substring(0, date_String.IndexOf(" "));
                    string[] date_elements = date_String.Split('-');
                    if (date_elements.Length >= 3) start_date = new DateTime(Int32.Parse(date_elements[2]),
                            Int32.Parse(date_elements[0]), Int32.Parse(date_elements[1]));
                    param_start_date = start_date.ToShortDateString();
                    param_start_date = param_start_date.Replace("/", "-");
                }
                if (param.Contains("end-date:"))
                {
                    string date_String = param.Substring(param.IndexOf("end-date:") + "end-date:".Length);
                    if (date_String.IndexOf(" ") >= 0) date_String = date_String.Substring(0, date_String.IndexOf(" "));
                    string[] date_elements = date_String.Split('-');
                    if (date_elements.Length >= 3) end_date = new DateTime(Int32.Parse(date_elements[2]),
                            Int32.Parse(date_elements[0]), Int32.Parse(date_elements[1]));
                    param_end_date = end_date.ToShortDateString();
                    param_end_date = param_end_date.Replace("/", "-");
                }

                foreach (string s in keywords.Keys)
                {
                    if (s.StartsWith(DATE))
                    {
                        string date = s.Substring(DATE.Length);
                        string[] date_elements = date.Split('-');
                        if (date_elements.Length >= 3) tmp_date = new DateTime(Int32.Parse(date_elements[2]),
                                Int32.Parse(date_elements[0]), Int32.Parse(date_elements[1]));
                        if (tmp_date >= start_date && tmp_date <= end_date)
                            results.UnionWith(keywords[s]);
                    }
                }
            }
            else
            {
                // Full set of photos into result:
                for (int i = 0; i < photo_media_play_list.count; i++) results.Add(i);
            }
            if (param.Contains("is-tagged:"))
            {
                string tag_string = param.Substring(param.IndexOf("is-tagged:") + "is-tagged:".Length);
                if (tag_string.IndexOf(" ") >= 0) tag_string = tag_string.Substring(0, tag_string.IndexOf(" "));
                param_is_tagged = tag_string;
                List<string> tag_list = new List<string>(tag_string.Split(';'));
                results = filter_is_tagged(results, tag_list);
            }
            else param_is_tagged = "";
            if (param.Contains("not-tagged:"))
            {
                string tag_string = param.Substring(param.IndexOf("not-tagged:") + "not-tagged:".Length);
                if (tag_string.IndexOf(" ") >= 0) tag_string = tag_string.Substring(0, tag_string.IndexOf(" "));
                param_not_tagged = tag_string;
                List<string> tag_list = new List<string>(tag_string.Split(';'));
                results = filter_not_tagged(results, tag_list);
            }
            else param_not_tagged = "";

            // Cache results:
            serialize_results(results, param);

            return results;
        }

        public Dictionary<string, List<int>> generate_filtered_tags_list(HashSet<int> results)
        {
            
            return null;
        }

        public OpResult clear_cache()
        {
            OpResult op_return = new OpResult(OpStatusCode.Ok);

            try { System.IO.Directory.Delete(CACHE_RESIZE_DIR, true); }
            catch (Exception ex)
            {
                op_return.StatusCode = OpStatusCode.Exception;
                op_return.StatusText = ex.Message;
                op_return.AppendFormat("Exception trying to delete cache directory: \"{0}\"", CACHE_RESIZE_DIR);
                op_return.AppendFormat("{0}", ex.Message);
            }
            try { System.IO.Directory.Delete(CACHE_QUERY_DIR, true); }
            catch (Exception ex)
            {
                op_return.StatusCode = OpStatusCode.Exception;
                op_return.StatusText = ex.Message;
                op_return.AppendFormat("Exception trying to delete cache directory: \"{0}\"", CACHE_QUERY_DIR);
                op_return.AppendFormat("{0}", ex.Message);
            }
            if (!in_generate_cache)
            {
                try { System.IO.Directory.Delete(CACHE_DIR, true); }
                catch (Exception ex)
                {
                    op_return.StatusCode = OpStatusCode.Exception;
                    op_return.StatusText = ex.Message;
                    op_return.AppendFormat("Exception trying to delete cache directory: \"{0}\"", CACHE_DIR);
                    op_return.AppendFormat("{0}", ex.Message);
                }
            }
            else
            {
                op_return.StatusCode = OpStatusCode.BadRequest;
                op_return.AppendFormat("ERROR: Cannot delete cache while generating cache!");
                if (keywords.ContainsKey(ELAPSED))
                    op_return.AppendFormat("Currently generating cache, last time this operation took {0} seconds.", keywords[ELAPSED][0]);
                else op_return.AppendFormat("Currently generating cache, First time generating - be patient.");
            }
            return op_return;
        }

        public OpResult clear_slide_show()
        {
            OpResult op_return = new OpResult(OpStatusCode.Ok);

            if (!in_generate_slideshow)
            {
                try { System.IO.Directory.Delete(SLIDE_SHOW_DIR, true); }
                catch (Exception ex)
                {
                    op_return.StatusCode = OpStatusCode.Exception;
                    op_return.StatusText = ex.Message;
                    op_return.AppendFormat("Exception trying to delete slide show directory: \"{0}\"", SLIDE_SHOW_DIR);
                    op_return.AppendFormat("{0}", ex.Message);
                }
            }
            else
            {
                op_return.StatusCode = OpStatusCode.BadRequest;
                op_return.AppendFormat("ERROR: Cannot delete slide shoe while generating slide show!");
            }
            return op_return;
        }

        public string make_cache_fn(string fn, int max_len)
        {
            fn = fn.Replace("\\", "_");
            fn = fn.Replace(":", "_");
            fn = fn.Replace(" ", "%20");
            fn = fn.Replace("\"", "_");

            if (fn.Length > max_len) fn = fn.Substring(fn.Length - max_len);

            return fn;
        }

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult Execute(string param)
        {
            reset_globals();

            debug_last_action = "Execute: Start";

            HashSet<int> results = new HashSet<int>();
            
            DateTime startTime = DateTime.Now;

            OpResult opResult = new OpResult();
            opResult.StatusCode = OpStatusCode.Ok;

            string template = "";

            page_limit = 0xFFFF;
            page_start = 0;

            bool bFilter = false;
            string clean_filter_params = "";
            string clean_paging_params = "";

            try
            {
                if (param.IndexOf("-help") >= 0)
                {
                    opResult = showHelp(opResult);
                    return opResult;
                }
                if (action == CLEAR_CACHE)
                {
                    opResult = clear_cache();
                    launch_generate_tags_list();
                    return opResult;
                }
                if (Player == null) Player = new WMPLib.WindowsMediaPlayer();
                photo_media_play_list = Player.mediaCollection.getByAttribute("MediaType", "Photo");

                debug_last_action = "Execution: Parsing params";
                // "Paging"?
                if (param.Contains("page-limit:"))
                {
                    page_limit = parse_int(param.Substring(param.IndexOf("page-limit:") + "page-limit:".Length));
                    clean_paging_params += "page-limit:" + page_limit + " ";
                }
                if (param.Contains("page-start:"))
                {
                    page_start = parse_int(param.Substring(param.IndexOf("page-start:") + "page-start:".Length));
                    clean_paging_params += "page-start:" + page_start + " ";
                }

                // Use Custom Template?
                if (param.Contains("template:"))
                {
                    template = param.Substring(param.IndexOf("template:") + "template:".Length);
                    if (template.IndexOf(" ") >= 0) template = template.Substring(0, template.IndexOf(" "));
                    clean_paging_params += "template:" + template + " ";
                }

                // Check if any filters are used
                if (param.Contains("is-tagged:"))
                {
                    bFilter = true;
                    string tag_string = param.Substring(param.IndexOf("is-tagged:") + "is-tagged:".Length);
                    if (tag_string.IndexOf(" ") >= 0) tag_string = tag_string.Substring(0, tag_string.IndexOf(" "));
                    List<string> tag_list = new List<string>(tag_string.Split(';'));
                    tag_list.Sort();
                    clean_filter_params += "is-tagged:" + String.Join(";", tag_list.ToArray()) + " "; 
                }
                if (param.Contains("not-tagged:"))
                {
                    bFilter = true;
                    string tag_string = param.Substring(param.IndexOf("not-tagged:") + "not-tagged:".Length);
                    if (tag_string.IndexOf(" ") >= 0) tag_string = tag_string.Substring(0, tag_string.IndexOf(" "));
                    List<string> tag_list = new List<string>(tag_string.Split(';'));
                    tag_list.Sort();
                    clean_filter_params += "not-tagged:" + String.Join(";", tag_list.ToArray()) + " ";
                }
                if (param.Contains("start-date:"))
                {
                    bFilter = true;
                    string date_string = param.Substring(param.IndexOf("start-date:") + "start-date:".Length);
                    DateTime date;
                    if (date_string.IndexOf(" ") >= 0) date_string = date_string.Substring(0, date_string.IndexOf(" "));
                    string[] date_elements = date_string.Split('-');
                    if (date_elements.Length >= 3)
                    {
                        date = new DateTime(Int32.Parse(date_elements[2]),
                            Int32.Parse(date_elements[0]), Int32.Parse(date_elements[1]));
                        date_string = date.ToShortDateString();
                        date_string = date_string.Replace("/", "-");
                        clean_filter_params += "start-date:" + date_string + " ";
                    }
                }
                if (param.Contains("end-date:"))
                {
                    bFilter = true;
                    string date_string = param.Substring(param.IndexOf("end-date:") + "end-date:".Length);
                    DateTime date;
                    if (date_string.IndexOf(" ") >= 0) date_string = date_string.Substring(0, date_string.IndexOf(" "));
                    string[] date_elements = date_string.Split('-');
                    if (date_elements.Length >= 3)
                    {
                        date = new DateTime(Int32.Parse(date_elements[2]),
                            Int32.Parse(date_elements[0]), Int32.Parse(date_elements[1]));
                        date_string = date.ToShortDateString();
                        date_string = date_string.Replace("/", "-");
                        clean_filter_params += "end-date:" + date_string + " ";
                    }
                }

                // Results specified? OVERRIDES OTHER FILTERS!
                if (param.Contains("ids:"))
                {
                    results = parse_ids(param.Substring(param.IndexOf("ids:") + "ids:".Length));
                    clean_filter_params = "ids:" + results.GetHashCode();
                }

                // Validate caches:
                validate_cache();
                // look for page cache - use filter + paging params, return if found
                if (action != PLAY_PHOTOS && action != QUEUE_PHOTOS && action != SERV_PHOTO && action != SHOW_STATS) //Check cache
                {
                    try
                    {
                        OpResult or = new OpResult();
                        or.StatusCode = OpStatusCode.Ok;
                        or = deserialize_rendered_page(or, action + "_" + clean_filter_params + clean_paging_params);
                        if (or != null) return or;
                    }
                    catch (Exception) { ; }
                }

                switch (action)
                {
                    case SHOW_STATS:
                    case LIST_TAGS:
                        //int pic_count = photo_media_play_list.count;
                        if (bFilter)    // Generate subset of tags based on images
                        {
                            if (results.Count == 0) results = get_list(results, clean_filter_params);
                            //pic_count = results.Count;
                            opResult = list_tags(opResult, template, generate_filtered_tags_list(results));
                        }
                        else            // Full set of tags
                            opResult = list_tags(opResult, template);
                        if (action == SHOW_STATS) opResult = show_stats(template, results);
                        break;
                    case LIST_PHOTOS:
                    case PLAY_PHOTOS:
                    case QUEUE_PHOTOS:
                    case SERV_PHOTO:
                        if (results.Count == 0)
                        {
                            results = get_list(results, clean_filter_params);
                        }

                        // Output results
                        if (action == SERV_PHOTO)
                        {
                            int size_x = 0;
                            int size_y = 0;
                            int size_short = 0;
                            int size_long = 0;
                            
                            debug_last_action = "Execute: Serv Start";
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
                            if (param.Contains("size-short:"))
                            {
                                string tmp_size = param.Substring(param.IndexOf("size-short:") + "size-short:".Length);
                                if (tmp_size.IndexOf(" ") >= 0) tmp_size = tmp_size.Substring(0, tmp_size.IndexOf(" "));
                                size_short = Convert.ToInt32(tmp_size);
                            }
                            if (param.Contains("size-long:"))
                            {
                                string tmp_size = param.Substring(param.IndexOf("size-long:") + "size-long:".Length);
                                if (tmp_size.IndexOf(" ") >= 0) tmp_size = tmp_size.Substring(0, tmp_size.IndexOf(" "));
                                size_long = Convert.ToInt32(tmp_size);
                            }
                            // Return one photo (may be only one)
                            int random_index = -1;
                            int image_index = 0;
                            if (param.Contains("random")) random_index = new Random().Next(results.Count);
                            int result_count = 0;
                            foreach (int i in results)
                            {
                                if (random_index < 0 && result_count >= page_start && result_count < page_start + page_limit)
                                {
                                    image_index = i;
                                    break;
                                }
                                else if (random_index >= 0 && random_index <= result_count)
                                {
                                    image_index = i;
                                    break;
                                }
                                result_count++;
                            }
                            last_img_served = image_index;
                            opResult = getPhoto(photo_media_play_list.get_Item(image_index), size_x, size_y, size_short, size_long);
                            debug_last_action = "Execute: Serv End";
                        }
                        else
                        {
                            string template_result = (action == LIST_PHOTOS) ? getTemplate(template, DEFAULT_RESULT) : getTemplate(template, DEFAULT_PLAY_RESULT);
                            template_result = file_includer(template_result);

                            debug_last_action = "Execute: List: Header";
                            string s_head = (action == LIST_PHOTOS) ? getTemplate(template + "+", DEFAULT_HEAD):getTemplate(template + "+", DEFAULT_PLAY_HEAD);
                            s_head = file_includer(s_head);
                            s_head = do_conditional_replace(s_head, "results_count", String.Format("{0}", results.Count));
                            opResult.AppendFormat("{0}", s_head);

                            debug_last_action = "Execute: List: Items";

                            if (action == PLAY_PHOTOS)
                            {
                                clear_slide_show();
                            }

                            int result_count = 0;
                            string in_file = "";
                            string out_file = "";
                            foreach (int i in results)
                            {
                                if (action == PLAY_PHOTOS || action == QUEUE_PHOTOS) 
                                {
                                    //Create dir if needed:
                                    if (!Directory.Exists(SLIDE_SHOW_DIR)) create_dirs();
                                    // Add photo to slideshow directory
                                    in_file = photo_media_play_list.get_Item(i).getItemInfo("SourceURL");
                                    out_file = SLIDE_SHOW_DIR + "\\" + make_cache_fn(in_file, 100);
                                    File.Copy(in_file, out_file, true);
                                }
                                else if (result_count >= page_start && result_count < page_start + page_limit)
                                {
                                    string s = "";
                                    s = replacer(template_result, photo_media_play_list.get_Item(i), i);
                                    opResult.AppendFormat("{0}", s);
                                }
                                result_count++;
                            }
                            if (action == PLAY_PHOTOS || action == QUEUE_PHOTOS)
                            {
                                // Start playing:
                                Microsoft.MediaCenter.Hosting.AddInHost.Current.MediaCenterEnvironment.NavigateToPage(Microsoft.MediaCenter.PageId.Slideshow, SLIDE_SHOW_DIR);
                                //Old: Microsoft.MediaCenter.Hosting.AddInHost.Current.MediaCenterEnvironment.NavigateToPage(Microsoft.MediaCenter.PageId.MyPictures, SLIDE_SHOW_DIR);
                            }
                            debug_last_action = "Execute: List: Footer";
                            TimeSpan duration = DateTime.Now - startTime;
                            string s_foot = (action == LIST_PHOTOS) ? getTemplate(template + "-", DEFAULT_FOOT):getTemplate(template + "-", DEFAULT_PLAY_FOOT);
                            s_foot = file_includer(s_foot);
                            s_foot = do_conditional_replace(s_foot, "elapsed_time", String.Format("{0}", duration.TotalSeconds));
                            s_foot = do_conditional_replace(s_foot, "results_count", String.Format("{0}", results.Count));
                            opResult.AppendFormat("{0}", s_foot);
                        }
                         
                        break;
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

            if (action == LIST_TAGS)
            {
                debug_last_action = "Execute: List: Footer";
                TimeSpan duration = DateTime.Now - startTime;
                string s_foot = getTemplate(template + "-", DEFAULT_REPORT_FOOT);
                s_foot = file_includer(s_foot);
                s_foot = do_conditional_replace(s_foot, "elapsed_time", String.Format("{0}", duration.TotalSeconds));
                s_foot = do_conditional_replace(s_foot, "results_count", String.Format("{0}", kw_count));
                s_foot = do_conditional_replace(s_foot, "first_date", String.Format("{0}", first_date.ToShortDateString()));
                s_foot = do_conditional_replace(s_foot, "last_date", String.Format("{0}", last_date.ToShortDateString()));
                opResult.AppendFormat("{0}", s_foot);
            }

            // save page to cache - use filter + paging params
            // TODO: elapsed time is incorrectly cached! Add is_cached to custom templates
            if (opResult.StatusCode == OpStatusCode.Ok &&
                action != PLAY_PHOTOS && action != QUEUE_PHOTOS && action != SERV_PHOTO && action != SHOW_STATS) // Save cache
            {
                debug_last_action = "Execute: Saving page to cache";
                try { serialize_rendered_page(opResult, action + "_" + clean_filter_params + clean_paging_params); }
                catch (Exception e) { debug_last_action = "Execute: Save to Cache failed: " + e.Message; }
            }

            debug_last_action = "Execute: End";

            return opResult;
        }

        #endregion
    }
}
