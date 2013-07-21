using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMPLib;

namespace VmcController.AddIn.Metadata
{
    public class PlaylistTrack : Track
    {
        public int playlist_number = 0;

        public PlaylistTrack(int index, IWMPMedia item) : base(item)
        {
            playlist_number = index + 1;
        }
    }
}
