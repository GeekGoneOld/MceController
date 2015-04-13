/*
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 * This software is provided 'as-is', without any express or implied warranty. 
 * In no event will the authors be held liable for any damages arising from the use of this software.
 * Permission is granted to anyone to use this software for any purpose, including commercial 
 * applications, and to alter it and redistribute it freely, subject to the following restrictions:
 * 
 * 1. The origin of this software must not be misrepresented; you must not claim that you wrote 
 *    the original software. If you use this software in a product, an acknowledgment in the 
 *    product documentation would be appreciated but is not required.
 * 2. Altered source versions must be plainly marked as such, and must not be misrepresented as
 *    being the original software.
 * 3. This notice may not be removed or altered from any source distribution.
 * 
 */
using System;
using System.Collections.Generic;
using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;
using Newtonsoft.Json;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for Volume command.
	/// </summary>
	public class MediaInfoCmd : IExperienceCommand
	{
        public const int NO_VOLUME_STATE = -1;

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "Gets info about current media (volume, play mode, playstate, playrate, position and metadata)";
        }

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult ExecuteMediaExperience(string param)
        {
            OpResult opResult = new OpResult();            
            try
            {
                opResult.StatusCode = OpStatusCode.Success;
                MediaInfo mediaInfo = new MediaInfo();

                //get volume
                mediaInfo.volume = (int)(AddInHost.Current.MediaCenterEnvironment.AudioMixer.Volume / 1310.7);
                mediaInfo.is_muted = AddInHost.Current.MediaCenterEnvironment.AudioMixer.Mute;

                //get metadata, play_state, play rate and position if available
                if (MediaExperienceWrapper.Instance != null)
                {
                    mediaInfo.play_state = Enum.GetName(typeof(PlayState), MediaExperienceWrapper.Instance.Transport.PlayState);
                    mediaInfo.play_rate = Enum.GetName(typeof(Play_Rate_enum), (Int32)MediaExperienceWrapper.Instance.Transport.PlayRate);
                    mediaInfo.position_sec = (Int32)Math.Round(MediaExperienceWrapper.Instance.Transport.Position.TotalSeconds);

                    if (MediaExperienceWrapper.Instance.MediaMetadata != null)
                    {
                        mediaInfo.metadata = MediaExperienceWrapper.Instance.MediaMetadata;

                        //search metadata to find type of media and duration
                        Object obj;
                        string objstr, str;
                        if (mediaInfo.metadata.TryGetValue("TrackDuration", out obj))
                        {
                            objstr = obj.ToString();
                            mediaInfo.duration_sec = Convert.ToInt32(objstr);
                            mediaInfo.play_mode = Enum.GetName(typeof(Play_Mode_enum), Play_Mode_enum.StreamingContentAudio);
                        }
                        else if (mediaInfo.metadata.TryGetValue("Duration", out obj))
                        {
                            TimeSpan ts;
                            objstr = obj.ToString();
                            TimeSpan.TryParse(objstr, out ts);
                            mediaInfo.duration_sec = (Int32)Math.Round(ts.TotalSeconds);
                            if (mediaInfo.metadata.ContainsKey("ChapterTitle"))
                            {
                                mediaInfo.play_mode = Enum.GetName(typeof(Play_Mode_enum), Play_Mode_enum.DVD);
                            }
                            else if (mediaInfo.metadata.TryGetValue("Name", out obj))
                            {
                                objstr = obj.ToString();
                                str = objstr.ToLower();
                                if (str.EndsWith("dvr_ms") | str.EndsWith("wtv"))
                                {
                                    mediaInfo.play_mode = Enum.GetName(typeof(Play_Mode_enum), Play_Mode_enum.PVR);
                                }
                                else
                                {
                                    mediaInfo.play_mode = Enum.GetName(typeof(Play_Mode_enum), Play_Mode_enum.StreamingContentVideo);
                                }
                            }
                        }
                    }
                    else
                    {
                        mediaInfo.play_mode = Enum.GetName(typeof(Play_Mode_enum), Play_Mode_enum.Undefined);
                    }
                }

                //convert times
                mediaInfo.position_hms = TimeSpan.FromSeconds((double)mediaInfo.position_sec);
                mediaInfo.duration_hms = TimeSpan.FromSeconds((double)mediaInfo.duration_sec);

                //flag data as valid
                mediaInfo.info_state = "Valid";

                opResult.ContentObject = mediaInfo;               
            }
            catch (Exception ex)
            {
                opResult.StatusCode = OpStatusCode.Exception;
                opResult.StatusText = ex.Message;
            }
            return opResult;
        }

        public class MediaInfo : OpResultObject
        {
            public bool is_muted = false;
            public int volume = NO_VOLUME_STATE;
            public string play_mode = Enum.GetName(typeof(Play_Mode_enum), Play_Mode_enum.Undefined);
            public string play_state = Enum.GetName(typeof(PlayState), PlayState.Undefined);
            public string play_rate = Enum.GetName(typeof(Play_Rate_enum), Play_Rate_enum.Stop);
            public Int32 position_sec = 0;
            public TimeSpan position_hms;
            public Int32 duration_sec = 0;
            public TimeSpan duration_hms;
            public IDictionary<string, object> metadata;
            public string info_state = "Unknown";
        }

        public enum Play_Mode_enum
        {
            Undefined = 0,
            StreamingContentAudio = 1,
            StreamingContentVideo = 2,
            PVR = 3,
            DVD = 4,
            TVTuner = 5,
            CD = 6,
            Photo = 7
        }

        public enum Play_Rate_enum
        {
            Stop = 0,
            Pause = 1,
            Play = 2,
            Ff1 = 3,
            Ff2 = 4,
            Ff3 = 5,
            Rewind1 = 6,
            Rewind2 = 7,
            Rewind3 = 8,
            Slowmotion1 = 9,
            Slowmotion2 = 10,
            Slowmotion3 = 11
        }

//        public enum PlayState
//        {
//            Stopped = 0,
//            Paused = 1,
//            Playing = 2,
//            Buffering = 3,
//            Finished = 4,
//            Undefined = 5,
//        }

        #endregion
    }
}
