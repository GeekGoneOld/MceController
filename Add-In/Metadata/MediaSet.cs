using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VmcController.AddIn.Metadata
{
    public class MediaSet
    {
        public int result_count = 0;

        public bool ShouldSerializeresult_count()
        {
            return(result_count != 0);
        }
    }
}
