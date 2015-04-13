/*
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 */
using System;
using System.Text;
using System.Diagnostics;
using System.Collections.Generic;
using VmcController.MceState;

namespace VmcController.Status 
{
    class ProcessCommands
    {
        public static string MakeHelpPage()
        {
            StringBuilder page = new StringBuilder();

            string page_start =
                "<html><head><script LANGUAGE='JavaScript'>\r\n" +
                "function toggle (o){ var all=o.childNodes;  if (all[0].childNodes[0].innerText == '+') open(o,true);  else open(o,false);}\r\n" +
                "function open (o, b) { var all=o.childNodes;  if (b) {all[0].childNodes[0].innerText='-';all[1].style.display='inline';}  else {all[0].childNodes[0].innerText='+';all[1].style.display='none';}}\r\n" +
                "function toggleAll (b){ var all=document.childNodes[0].childNodes[1].childNodes; for (var i=0; i<all.length; i++) {if(all[i].id=='section') open(all[i],b)};}\r\n" +
                "</script></head>\r\n" +
                "<body><a id='top'>Jump to</a>: <a href='#commands'>Command List</a>, <a href='#media-info'>Media Info</a>, <a href='#bottom'>Bottom</a><hr><font size=+3><a id='commands'>Command List</a>:&nbsp;&nbsp;</font>[<a onclick='toggleAll(true);' >Open All</a>] | [<a onclick='toggleAll(false);' >Collapse All</a>]<hr>\r\n";
            string page_end =
                "</pre></span></div>\r\n" +
                "<br><hr><font size=+3><a id='media-info'>Media Info</a>:</font><hr><br>\r\n" +
                "<b><font size=+1>Important note!</font></b><p>" +
                "This information comes from Media State Aggregation Service (MSAS) which was deprecated in Vista.  " +
                "It is known to work for Vista and 7, but its future is uncertain.  It is preferred that you use the command/info on " +
                "port 4051x instead.  This provides commands as well as feedback, but some feedback is missing (e.g. Live TV channel)." +
                "<p>" +
                "Since there may be multiple media sessions, all are reported in no guaranteed order.  This could cause a parser to miss some media " +
                "(e.g. MediaName could exist on multiple sessions).  To minimize this problem, any background session (Recording is the only known one) " +
                "tags are reported prefixed with BG_ (e.g. BG_MediaName)." +
                "<p>" +
                "<b>volume</b> : volume setting<br>" +
                "Possible values are:<br>" +
                "&nbsp;&nbsp;0-50<br>" +
                "<p>" +
                "<b>mute</b> : mute state<br>" +
                "Possible values are:<br>" +
                "&nbsp;&nbsp;false<br>" +
                "&nbsp;&nbsp;true<br>" +
                "<p>" +
                "<b>m_page</b> : current page displayed on Media Center<br>" +
                "Note that this is reported as two properties (m_page=&lt;value&gt; and &lt;value&gt;=true)<br>" +
                "Possible values are:<br>" +
                "&nbsp;&nbsp;FS_DVD<br>" +
                "&nbsp;&nbsp;FS_Extensibility<br>" +
                "&nbsp;&nbsp;FS_Guide<br>" +
                "&nbsp;&nbsp;FS_Home<br>" +
                "&nbsp;&nbsp;FS_Music<br>" +
                "&nbsp;&nbsp;FS_Photos<br>" +
                "&nbsp;&nbsp;FS_Radio<br>" +
                "&nbsp;&nbsp;FS_RecordedShows<br>" +
                "&nbsp;&nbsp;FS_TV<br>" +
                "&nbsp;&nbsp;FS_Unknown<br>" +
                "&nbsp;&nbsp;FS_Videos<br>" +
                "<p>" +
                "<b>m_mediaMode</b> : type of media playing or most recently played<br>" +
                "Note that this is reported as two properties (m_page=&lt;value&gt; and &lt;value&gt;=true)<br>" +
                "Possible values are:<br>" +
                "&nbsp;&nbsp;CD<br>" +
                "&nbsp;&nbsp;DVD<br>" +
                "&nbsp;&nbsp;PhoneCall<br>" +
                "&nbsp;&nbsp;Photos<br>" +
                "&nbsp;&nbsp;PVR<br>" +
                "&nbsp;&nbsp;Radio<br>" +
                "&nbsp;&nbsp;Recording<br>" +
                "&nbsp;&nbsp;StreamingContentAudio<br>" +
                "&nbsp;&nbsp;StreamingContentVideo<br>" +
                "&nbsp;&nbsp;TVTuner<br>" +
                "<p>" +
                "<b>m_playRate</b> : play rate of current media<br>" +
                "Note that this is reported as two properties (m_page=&lt;value&gt; and &lt;value&gt;=true)<br>" +
                "Possible values are:<br>" +
                "&nbsp;&nbsp;FF1<br>" +
                "&nbsp;&nbsp;FF2<br>" +
                "&nbsp;&nbsp;FF3<br>" +
                "&nbsp;&nbsp;Pause<br>" +
                "&nbsp;&nbsp;Play<br>" +
                "&nbsp;&nbsp;Rewind1<br>" +
                "&nbsp;&nbsp;Rewind2<br>" +
                "&nbsp;&nbsp;Rewind3<br>" +
                "&nbsp;&nbsp;SlowMotion1<br>" +
                "&nbsp;&nbsp;SlowMotion2<br>" +
                "&nbsp;&nbsp;SlowMotion3<br>" +
                "&nbsp;&nbsp;Stop<br>" +
                "<p>" +
                "<b>ArtistName</b> : artist name for current track<br>" +
                "<b>CallingPartyName</b> : calling party name for current phone call<br>" +
                "<b>CallingPartyNumber</b> : calling party number for current phone call<br>" +
                "<b>CurrentPicture</b> : name of current picture<br>" +
                "<b>DiscWriter_ProgressPercentageChanged</b> : percentage complete of current CD/DVD write operation<br>" +
                "<b>DiscWriter_ProgressTimeChanged</b> : elapsed time of current CD/DVD write operation<br>" +
                "<b>DiscWriter_SelectedFormat</b> : selected format of current CD/DVD write operation<br>" +
                "<b>DialogVisible</b> : set to true when a dialog is visible and false when dialog is closed<br>" +
                "<b>MediaName</b> : name of current media<br>" +
                "<b>MediaTime</b> : total length of current media<br>" +
                "<b>ParentalAdvisoryRating</b> : MPAA rating of current track or DVD<br>" +
                "<b>RadioFrequency</b> : frequency of current radio station<br>" +
                "<b>RepeatSet</b> : repeat mode state<br>" +
                "<b>Shuffle</b> : shuffle mode state<br>" +
                "<b>TitleNumber</b> : number of current title<br>" +
                "<b>TotalTracks</b> : total number of tracks on this album<br>" +
                "<b>TrackDuration</b> : total length of this track<br>" +
                "<b>TrackName</b> : name of this track<br>" +
                "<b>TrackNumber</b> : track number of this track<br>" +
                "<b>TrackTime</b> : elapsed time of this track<br>" +
                "<b>TransitionTime</b> : photo transition time<br>" +
                "<b>Visualization</b> : name of the current visualization<br>" +
                "<br>" +
                "<b>RequestForTuner</b> : don't really know...<br>" +
                "<b>Unknown</b> : don't really know...<br>" +
                "<br>" +
                "<b>DiscWriter_Start</b> : set to true when current CD/DVD write operation is started (useless)<br>" +
                "<b>DiscWriter_Stop</b> : set to true when current CD/DVD write operation is finished or stopped (useless)<br>" +
                "<b>Ejecting</b> : set to true when current CD/DVD is ejected (useless)<br>" +
                "<b>GuideLoaded</b> : set to true when a new TV guide is downloaded (useless)<br>" +
                "<b>Next</b> : set to true when skip next is pressed (useless)<br>" +
                "<b>NextFrame</b> : set to true when skip next frame is pressed (useless)<br>" +
                "<b>Prev</b> : set to true when skip prev is pressed (useless)<br>" +
                "<b>PrevFrame</b> : set to true when skip prev frame is pressed (useless)<br>" +
                "<br><hr>\r\n" +
                "<a id='bottom'>Generated by</a>: Windows Media Center TCP/IP Controller (<a href='http://github.com/GeekGoneOld/MceController'>WMCController Home</a>)\r\n" +
                "<hr>Jump to: <a href='#top'>Top</a>, <a href='#commands'>Command List</a>, <a href='#media-info'>Media Info</a><br>\r\n" +
                "<script LANGUAGE='JavaScript'>toggleAll(false);</script></body></html>\r\n";
            string header_start = "</pre></span></div><br><div id='section' onclick='toggle(this)' style='border:solid 1px black;'><font size=+1 style='font:15pt courier;'><span>-</span>";
            string header_end = "</font><span style='display:'><pre>";

            page.AppendFormat("{0}\r\n", page_start);

            page.AppendFormat("{0} Ports: {1}\r\n", header_start, header_end);
            page.AppendFormat("HTTP command/server port: {0} (e.g http://your_server:{1}/\r\n", "4051x", 40510);
            page.AppendFormat("HTTP server port (deprecated): {0} (e.g http://your_server:{1}/\r\n", "4041x", 40410);
            page.AppendFormat("Streaming data server TCP port (deprecated): {0} (e.g. use PuTTY to connect to your_server:{1} raw)", "4040x", 40400);

            page.AppendFormat("{0} Commands: {1}\r\n", header_start, header_end);
            page.AppendFormat("<a href='/help'>help</a> - displays this help page\r\n");
            page.AppendFormat("<a href='/media-info'>media-info</a> - displays media info in plain text (compatible with 4040x)\r\n");

            page.Append(page_end);


            return page.ToString();
        }

        public static string GetMediaInfo()
        {
			StringBuilder page = new StringBuilder();
            string prefix;
            bool pg_set = false;
            bool pr_set = false;
            bool mm_set = false;

            try
            {
                foreach (KeyValuePair<int, MediaState> mediaState in MediaStateDict.mediaStates)
                {
                    page.AppendFormat("SESSION={0}\r\n", mediaState.Key);

                    if (mediaState.Value.MediaMode == MediaState.MEDIASTATUSPROPERTYTAG.Recording)
                        prefix = "BG_";
                    else
                        prefix = "";

				    //  Provide current state information to the client
				    if (!string.IsNullOrEmpty(mediaState.Value.Volume)) {
                        page.AppendFormat(prefix + "Volume={0}\r\n", mediaState.Value.Volume);
				    }
                    if (!string.IsNullOrEmpty(mediaState.Value.Mute))
                    {
					    page.AppendFormat(prefix + "Mute={0}\r\n", mediaState.Value.Mute);
				    }
				    if (mediaState.Value.Page != MediaState.MEDIASTATUSPROPERTYTAG.Unknown) {
                        page.AppendFormat(prefix + "{0}=True\r\n", mediaState.Value.Page);
                        page.AppendFormat(prefix + "m_page={0}\r\n", mediaState.Value.Page);
                        if (prefix == "")
                            pg_set = true;
                    }
				    if (mediaState.Value.PlayRate != MediaState.MEDIASTATUSPROPERTYTAG.Unknown) {
                        page.AppendFormat(prefix + "{0}=True\r\n", mediaState.Value.PlayRate);
                        page.AppendFormat(prefix + "m_playRate={0}\r\n", mediaState.Value.PlayRate);
                        if (prefix == "")
                            pr_set = true;
                    }
                    if (mediaState.Value.MediaMode != MediaState.MEDIASTATUSPROPERTYTAG.Unknown)
                    {
                        page.AppendFormat(prefix + "{0}=True\r\n", mediaState.Value.MediaMode);
                        page.AppendFormat(prefix + "m_mediaMode={0}\r\n", mediaState.Value.MediaMode);
                        if (prefix == "")
                            mm_set = true;
                    }
                    foreach (KeyValuePair<string, object> item in mediaState.Value.MetaData)
					    page.AppendFormat(prefix + "{0}={1}\r\n", item.Key, item.Value);

				    //  Send the data to the connected client
				    Trace.TraceInformation(page.ToString());
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
            }
            finally
            {
            }
            page.AppendFormat("SESSION={0}\r\n", 0);
            if (!pg_set)
                page.AppendFormat("m_page={0}\r\n", MediaState.MEDIASTATUSPROPERTYTAG.Unknown);
            if (!pr_set)
                page.AppendFormat("m_playRate={0}\r\n", MediaState.MEDIASTATUSPROPERTYTAG.Unknown);
            if (!mm_set)
                page.AppendFormat("m_mediaMode={0}\r\n", MediaState.MEDIASTATUSPROPERTYTAG.Unknown);
            page.AppendFormat("status=Ok\r\n");
            return page.ToString();
        }
    }
}