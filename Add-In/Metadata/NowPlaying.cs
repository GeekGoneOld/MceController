using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMPLib;

namespace VmcController.AddIn.Metadata
{
    public class NowPlaying : OpResultObject
    {
        public ArrayList now_playing = new ArrayList();


        public NowPlaying(RemotedWindowsMediaPlayer remotePlayer)
        {
            IWMPPlaylist playlist = remotePlayer.getNowPlaying();
            if (playlist != null && playlist.count > 0)
            {
                result_count = playlist.count;
                for (int j = 0; j < playlist.count; j++)
                {
                    IWMPMedia item = playlist.get_Item(j);
                    if (item != null)
                    {
                        now_playing.Add(new PlaylistTrack(j, item));
                    }
                }
            }            
        }
    }
}
