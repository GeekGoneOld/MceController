using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMPLib;

namespace VmcController.AddIn.Metadata
{
    public class Song : IComparable<Song>
    {
        public string song = "";
        public string song_artist = "";
        public string genre = "";
        public string year = "";
        public string duration = "";

        public Song(string name)
        {
            song = name;
        }

        public Song(IWMPMedia item)
        {
            if (item != null)
            {
                song = item.getItemInfo("Title");
                song_artist = item.getItemInfo("Author");
                duration = item.durationString;
                year = item.getItemInfo("WM/OriginalReleaseYear");
                if (year.Equals("") || year.Length < 4) year = item.getItemInfo("WM/Year");
                genre = item.getItemInfo("WM/Genre"); 
            }
        }

        public override string ToString()
        {
            return song;
        }

        public bool Equals(Song s)
        {
            if (s == null || song == null) return false;
            return song.Equals(s.song);
        }

        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType() || song == null) return false;
            Song s = (Song)obj;
            return song.Equals(s.song);
        }

        public int CompareTo(Song s)
        {
            if (s == null || song == null) return 1;
            return song.CompareTo(s.song);            
        }
    }
}
