using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace VmcController.AddIn
{
    public class MediaExperienceWrapper
    {
        private static MediaExperience mce;

        private MediaExperienceWrapper()
        {
        }

        public static MediaExperience Instance
        {
            get
            {
                if (mce == null) mce = AddInHost.Current.MediaCenterEnvironment.MediaExperience;

                // possible workaround for Win7/8 bug
                if (mce == null)
                {
                    System.Threading.Thread.Sleep(200);
                    mce = AddInHost.Current.MediaCenterEnvironment.MediaExperience;
                    if (mce == null)
                    {
                        try
                        {
                            var fi = AddInHost.Current.MediaCenterEnvironment.GetType().GetField("_checkedMediaExperience", BindingFlags.NonPublic | BindingFlags.Instance);
                            if (fi != null)
                            {
                                fi.SetValue(AddInHost.Current.MediaCenterEnvironment, false);
                                mce = AddInHost.Current.MediaCenterEnvironment.MediaExperience;
                            }

                        }
                        catch (Exception)
                        {
                            // give up 
                        }

                    }
                }
                return mce; 
            }
        }
    }
}
