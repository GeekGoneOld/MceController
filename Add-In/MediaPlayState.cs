using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMPLib;

namespace VmcController.AddIn
{
    public class MediaPlayState
    {
        private const int UNDEFINED = 0;
        private const int PAUSED = 1;
        private const int PLAYING = 2;
        private const int STOPPED = 3;
        private int playState;

        public MediaPlayState()
        {
            playState = UNDEFINED;
        }

        public MediaPlayState(WMPPlayState state)
        {
            switch (state)
            {
                case WMPPlayState.wmppsPaused:
                    playState = PAUSED;
                    break;
                case WMPPlayState.wmppsPlaying:
                    playState = PLAYING;
                    break;
                case WMPPlayState.wmppsStopped:
                    playState = STOPPED;
                    break;
                default:
                    playState = UNDEFINED;
                    break;
            }
        }

        public string getState()
        {
            switch (playState)
            {
                case PAUSED:
                    return "Pause";
                case PLAYING:
                    return "Play";
                case STOPPED:
                    return "Stop";
                default:
                    return "Undefined";
            }
        }

        public bool isStopped()
        {
            return playState == STOPPED;
        }

        public bool isPlaying()
        {
            return playState == PLAYING;
        }

        public bool isPaused()
        {
            return playState == PAUSED;
        }

        public bool isUndefined()
        {
            return playState == UNDEFINED;
        }
    }
}
