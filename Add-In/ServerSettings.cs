using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using WMPLib;

namespace VmcController.AddIn
{
    public class ServerSettings : OpResultObject
    {
        public string version = "";
        public bool is_building = false;
        private bool cache_outdated;
        public int cache_build_hour = 4;
        public int send_buffer_size = HttpSocketServer.MIN_SEND_BUFFER_SIZE;

        public ServerSettings()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);
            version = fileVersionInfo.ProductVersion;
        }

        public bool is_cache_outdated
        {
            set { cache_outdated = value; }
            get { return cache_outdated; }
        }
    }
}
