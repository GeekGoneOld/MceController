using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMPLib;

namespace VmcController.AddIn.Metadata
{
    public class Album : MediaSet
    {
        private bool m_stats_only;
        public string album = "";
        public string album_artist = "";
        public string year = "";
        public ArrayList genres = new ArrayList();
        public ArrayList tracks;


        public Album(string name, IWMPPlaylist playlist)
        {
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
                        if (!genres.Contains(item.getItemInfo("WM/Genre"))) genres.Add(item.getItemInfo("WM/Genre"));
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
            ArrayList items = new ArrayList();
            for (int j = 0; j < playlist.count; j++)
            {
                IWMPMedia item = playlist.get_Item(j);
                if (item != null)
                {
                    addTrack(items, item);
                }
            }
            items.TrimToSize();
            result_count = items.Count;
            if (!m_stats_only)
            {
                tracks = items;
            }
            return result_count;
        }

        public void addTrack(ArrayList items, IWMPMedia item)
        {
            Track track = new Track(item);
            if (!genres.Contains(track.genre)) genres.Add(track.genre);
            year = track.year;
            album_artist = item.getItemInfo("WM/AlbumArtist");
            items.Add(track);
        }

        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType()) return false;
            Album a = (Album)obj;
            return album.Equals(a.album);
        }
    }
}
