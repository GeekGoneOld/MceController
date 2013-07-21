using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMPLib;

namespace VmcController.AddIn.Metadata
{
    public class Track : Song
    {
        public int number = 0;
        public string album = "";


        public Track(IWMPMedia item) : base(item)
        {
            if (item != null)
            {
                album = item.getItemInfo("WM/AlbumTitle");
                try
                {
                    number = Convert.ToInt32(item.getItemInfo("WM/TrackNumber"));
                }
                catch (Exception)
                {
                    number = 0;                    
                }
            }
        }
    }
}
