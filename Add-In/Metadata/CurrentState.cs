using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VmcController.AddIn.Commands;
using VmcController.AddIn.Metadata;
using WMPLib;

namespace VmcController.AddIn
{
    public class CurrentState : OpResultObject
    {
        [JsonConverter(typeof(StringEnumConverter))]
        public TruncatedWMPPlayState play_state = TruncatedWMPPlayState.Undefined;
        public bool shuffle_mode = false;
        public int volume = VolumeCmd.NO_VOLUME_STATE;
        public bool is_muted = false;
        public PlaylistTrack current_track;


        public CurrentState(RemotedWindowsMediaPlayer remotePlayer)
        {
            IWMPMedia current_item = remotePlayer.getCurrentMediaItem();
            IWMPPlaylist playlist = remotePlayer.getNowPlaying();
            int index = -1;
            if (playlist != null && playlist.count > 0) 
            {
                for (int j = 0; j < playlist.count; j++)
                {
                    IWMPMedia item = playlist.get_Item(j);
                    if (item != null && item.get_isIdentical(current_item))
                    {
                        index = j;
                    }
                }
            }
            if (index >= 0)
            {
                current_track = new PlaylistTrack(index, current_item); 
            }
            shuffle_mode = remotePlayer.isShuffleModeEnabled();
            play_state = getTruncatedPlayState(remotePlayer.getPlayState());
            VolumeCmd volumeCmd = new VolumeCmd();
            volume = volumeCmd.getVolume();
            is_muted = volumeCmd.isMuted();
        }

        public static TruncatedWMPPlayState getTruncatedPlayState(WMPPlayState state)
        {
            switch (state)
            {
                case WMPPlayState.wmppsBuffering:
                    return TruncatedWMPPlayState.Buffering;
                case WMPPlayState.wmppsMediaEnded:
                    return TruncatedWMPPlayState.MediaEnded;
                case WMPPlayState.wmppsReconnecting:
                    return TruncatedWMPPlayState.Reconnecting;
                case WMPPlayState.wmppsScanForward:
                    return TruncatedWMPPlayState.ScanForward;
                case WMPPlayState.wmppsScanReverse:
                    return TruncatedWMPPlayState.ScanReverse;
                case WMPPlayState.wmppsTransitioning:
                    return TruncatedWMPPlayState.Transitioning;
                case WMPPlayState.wmppsWaiting:
                    return TruncatedWMPPlayState.Waiting;
                case WMPPlayState.wmppsPaused:
                    return TruncatedWMPPlayState.Paused;
                case WMPPlayState.wmppsPlaying:
                    return TruncatedWMPPlayState.Playing;
                case WMPPlayState.wmppsStopped:
                    return TruncatedWMPPlayState.Stopped;
                case WMPPlayState.wmppsUndefined:
                    return TruncatedWMPPlayState.Undefined;
                default:
                    return TruncatedWMPPlayState.Undefined;
            }
        }
    }
}
