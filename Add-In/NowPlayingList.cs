using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMPLib;

namespace VmcController.AddIn
{
    public class NowPlayingList
    {
        public ArrayList now_playing = new ArrayList();

        public NowPlayingList(IWMPPlaylist original)
        {
            for (int j = 0; j < original.count; j++)
            {
                IWMPMedia item = original.get_Item(j);
                if (item != null)
                {
                    now_playing.Add(new MediaItem(item));
                }                                
            }
        }

        public int Count()
        {
            return now_playing.Count;
        }

        public MediaItem get_Item(int index)
        {            
            return (MediaItem) now_playing[index];
        }
    }
}
