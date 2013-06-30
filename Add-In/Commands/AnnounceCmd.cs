using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Speech.Synthesis;
using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;
using System.IO;
using System.Diagnostics;
using System.Threading;


namespace VmcController.AddIn.Commands
{
    public class Messenger
    {
        private string _MsgBoxRichCmdString;
        private bool _ready = false;

        public Messenger(string MsgBoxRichCmdString)
        {
            _MsgBoxRichCmdString = MsgBoxRichCmdString;
        }

        public bool Ready
        {
            get
            { return _ready; }
        }

        public void ShowMessage()
        {
            OpResult result;
            try
            {
                MsgBoxRichCmd msg = new MsgBoxRichCmd();
                result = msg.Execute(_MsgBoxRichCmdString);
            }
            catch (Exception) { }

            _ready = true;
        }


    }

    class AnnounceCmd : ICommand
    {

        private static Regex m_regex = new Regex("\"(?<caption>.+?)\"\\s+\"(?<message>.+?)\"\\s+(?<timeout>\\d+)\\s+\"(?<buttoncodes>.+?)\"\\s+\"(?<modal>.+?)\"\\s+\"(?<imagepath>.+?)\"\\s+\"(?<ssmltospeak>.+?)\"");

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "\"caption\" \"message\" <timeout seconds> \"button codes\" \"modal|nonmodal\" \"imagepath\" \"ssmltospeak\"";
        }

        private bool isPlaying()
        {
            //bool isPlaying = false;

            if (AddInModule.getMediaExperience() == null)
            {
                return false;
            }
            else if (AddInHost.Current.MediaCenterEnvironment.MediaExperience.Transport.PlayState == PlayState.Stopped
                    | AddInHost.Current.MediaCenterEnvironment.MediaExperience.Transport.PlayState == PlayState.Finished
                    | AddInHost.Current.MediaCenterEnvironment.MediaExperience.Transport.PlayState == PlayState.Undefined)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult Execute(string param)
        {
            OpResult opResult = new OpResult();

            //try
            //{
            Match match = m_regex.Match(param);

            if (match.Success)
            {


                //try
                //{
                //msgResult = msg.Execute(msgboxcmd);
                //}
                //catch (Exception ex)
                //{
                //EventLog.WriteEntry("VmcController Client AddIn", "Error in Announce: " + ex.ToString());
                //}

                //EventLog.WriteEntry("VmcController Client AddIn", "Past msgboxrich");
                //AddInHost.Current.MediaCenterEnvironment.Dialog("Doing the playback", "", DialogButtons.Ok, 5, true);


                //first show the message box
                //MsgBoxRichCmd msg = new MsgBoxRichCmd();
                //string msgboxcmd = "\"" + match.Groups["caption"].Value + "\" \"" + match.Groups["message"].Value + "\" " + match.Groups["timeout"].Value + " \"" + match.Groups["buttoncodes"].Value + "\" \"" + match.Groups["modal"].Value + "\" \"" + match.Groups["imagepath"].Value + "\"";

                Messenger msg = new Messenger("\"" + match.Groups["caption"].Value + "\" \"" + match.Groups["message"].Value + "\" " + match.Groups["timeout"].Value + " \"" + match.Groups["buttoncodes"].Value + "\" \"" + match.Groups["modal"].Value + "\" \"" + match.Groups["imagepath"].Value + "\"");
                //EventLog.WriteEntry("VmcController Client AddIn", "Executing msgboxrich: " + msgboxcmd);

                //OpResult msgResult;

                //AddInHost.Current.MediaCenterEnvironment.Dialog("Playback state: " + isPlaying.ToString(), "", DialogButtons.Ok, 5, true);
                //then if not playing, speak the message

                Thread t = new Thread(new ThreadStart(msg.ShowMessage));
                if (isPlaying() == true)
                {
                    t.Start();
                }
                else
                {
                    //AddInHost.Current.MediaCenterEnvironment.Dialog("Not playing. Proceeding", "", DialogButtons.Ok, 5, true);

                    FileInfo fi;
                    if (match.Groups["ssmltospeak"].Value.Length > 1)
                    {
                        fi = MakeAudioFile(match.Groups["ssmltospeak"].Value);
                    }
                    else
                    {
                        //AddInHost.Current.MediaCenterEnvironment.Dialog("Making the audio file from the message", "", DialogButtons.Ok, 5, true);
                        fi = MakeAudioFile(match.Groups["message"].Value);
                    }


                    //long fileDurationInMS = (fi.Length / 352000) * 1000;  //much older code


                    //long fileDurationInMS = ((fi.Length / 44100) * 1000)-3000;  //current code
                    //AddInHost.Current.MediaCenterEnvironment.Dialog("Duration: " + fileDurationInMS.ToString(), "", DialogButtons.Ok, 3, true);
                    /*
                    WMPLib.WindowsMediaPlayer Player = new WMPLib.WindowsMediaPlayer();
                    Player.settings.autoStart = true;
                    Player.URL = fi.FullName;

                        
                    */



                    //show the message box
                    //msgResult = msg.Execute(msgboxcmd);

                    t.Start();
                    //ThreadPriority PreviousPriority = Thread.CurrentThread.Priority;
                    //Thread.CurrentThread.Priority = ThreadPriority.Highest;
                    WMPLib.WindowsMediaPlayer Player = new WMPLib.WindowsMediaPlayer();
                    Player.settings.playCount = 1;
                    Player.settings.setMode("loop", false);
                    Player.URL = fi.FullName;

                    //AddInHost.Current.MediaCenterEnvironment.PlayMedia(MediaType.Audio, fi.FullName, false);

                    //DateTime ExpectedEndTime = DateTime.Now.AddMilliseconds(fileDurationInMS);
                    //AddInHost.Current.MediaCenterEnvironment.Dialog("Now we're playing", "", DialogButtons.Ok, 5, true);
                    //AddInHost.Current.MediaCenterEnvironment.Dialog("Playstate: " + (AddInHost.Current.MediaCenterEnvironment.MediaExperience == null).ToString(), "", DialogButtons.Ok, 5, true);
                    //SendKeyCmd StopCmd = new SendKeyCmd('S', true, true, false);
                    //OpResult StopResult = null;

                    //while (AddInHost.Current.MediaCenterEnvironment.MediaExperience.Transport.PlayState == PlayState.Playing)

                    //wait for the required duration to elapse


                    //while (isPlaying() == true)
                    //{
                    //  if (DateTime.Compare(DateTime.Now,ExpectedEndTime) >= 0)
                    //{

                    //ensure it only plays once, regardless of shuffle/repeat state. Don't know why we have to keep sending the stop command
                    //probably something to do with timing
                    //while (AddInHost.Current.MediaCenterEnvironment.MediaExperience.Transport.PlayState == PlayState.Playing)
                    //{
                    //StopResult = StopCmd.Execute("");

                    //System.Windows.Forms.Application.DoEvents(); //ensures that the stop command is responded to
                    //System.Threading.Thread.Sleep(100);
                    //}
                    //break;
                    // }

                    // System.Threading.Thread.Sleep(100);
                    //}
                    //Thread.CurrentThread.Priority = PreviousPriority;
                    fi.Delete();

                }

                //opResult.AppendFormat(msgResult.StatusText);
                opResult.AppendFormat("Ok");
                opResult.StatusCode = OpStatusCode.Ok;
            }

            //}
            //catch (Exception ex)
            //{
            //    EventLog.WriteEntry("VmcController Client AddIn", ex.ToString());
            //    opResult.StatusCode = OpStatusCode.Exception;
            //    opResult.StatusText = ex.Message;
            //}
            return opResult;
        }

        #endregion

        public FileInfo MakeAudioFile(string ssmlToSpeak)
        {

            string wavFileName = DateTime.Now.ToString().Replace(" ", "");
            //AddInHost.Current.MediaCenterEnvironment.Dialog("Doing the playback", "", DialogButtons.Ok, 5, true);
            wavFileName = System.Environment.GetEnvironmentVariable("TEMP") + "\\" + wavFileName.Replace("-", "").Replace("/", "").Replace(":", "") + ".wav";
            //EventLog.WriteEntry("VmcController Client AddIn", "Generating speech file to " + wavFileName);

            ssmlToSpeak = "<speak version='1.0' xml:lang='en-US'>" + ssmlToSpeak + "<break time='2000ms'/></speak>";

            using (SpeechSynthesizer synth = new SpeechSynthesizer())
            {

                synth.SetOutputToWaveFile(wavFileName);

                synth.SpeakSsml(ssmlToSpeak);

            }

            //EventLog.WriteEntry("VmcController Client AddIn", "Speech file written to " + wavFileName);

            FileInfo fi = new FileInfo(wavFileName);

            return fi;
        }
    }
}
