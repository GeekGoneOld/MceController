using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMPLib;

namespace VmcController.AddIn.Metadata
{
    public class Playlist : MediaSet
    {
        public string playlist = "";
        public ArrayList tracks;
        private bool m_stats_only;

        public Playlist(string name, bool stats_only)
        {
            m_stats_only = stats_only;
            playlist = name;
        }

        public int addTracks(IWMPPlaylist playlist)
        {
            ArrayList items = new ArrayList();
            for (int j = 0; j < playlist.count; j++)
            {
                IWMPMedia item = playlist.get_Item(j);
                if (item != null)
                {
                    items.Add(new PlaylistTrack(j, item));
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

        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType()) return false;
            Playlist p = (Playlist)obj;
            return playlist.Equals(p.playlist);
        }
    }
}
