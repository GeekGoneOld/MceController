using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VmcController.AddIn.Metadata;

namespace VmcController.AddIn
{
    public class OpResultObject : MediaSet
    {
        [JsonConverter(typeof(StringEnumConverter))]
        public OpStatusCode status_code = OpStatusCode.BadRequest;
        public string status_message;
    }
}
