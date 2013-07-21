using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VmcController.AddIn.Metadata
{
    public class PlayStateObject : OpResultObject
    {
        [JsonConverter(typeof(StringEnumConverter))]
        public TruncatedWMPPlayState play_state;
    }
}
