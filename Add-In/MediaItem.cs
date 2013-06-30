using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMPLib;

namespace VmcController.AddIn
{
    public class MediaItem
    {
        public string album = "";
        public string album_artist = "";
        public string song = "";
        public string song_artist = "";
        public string number = "";
        public string duration = "";
        public string play_state = "";

        public MediaItem(IWMPMedia media)
        {
            if (media != null)
            {
                album = media.getItemInfo("WM/AlbumTitle");
                album_artist = media.getItemInfo("WM/AlbumArtist");
                song = media.getItemInfo("Title");
                song_artist = media.getItemInfo("Author");
                number = media.getItemInfo("WM/TrackNumber");
                duration = media.durationString; 
            }
        }

        public string getItemInfo(string bstrItemName)
        {
            if (bstrItemName != null)
            {
                if (bstrItemName.Equals("WM/AlbumTitle")) return album;
                if (bstrItemName.Equals("WM/AlbumArtist")) return album_artist;
                if (bstrItemName.Equals("Title")) return song;
                if (bstrItemName.Equals("Author")) return song_artist;
                if (bstrItemName.Equals("WM/TrackNumber")) return number;
                if (bstrItemName.Equals("Duration")) return duration; 
            }
            return null;
        }
    }
}
