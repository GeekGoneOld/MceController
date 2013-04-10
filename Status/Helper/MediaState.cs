/*
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 */
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;

namespace VmcController.MceState {
	/// <summary>
	/// Tracks state information
	/// </summary>
	public static class MediaState {
		private static MEDIASTATUSPROPERTYTAG m_mediaMode = MEDIASTATUSPROPERTYTAG.Unknown;
		private static MEDIASTATUSPROPERTYTAG m_playRate = MEDIASTATUSPROPERTYTAG.Unknown;
		private static MEDIASTATUSPROPERTYTAG m_page = MEDIASTATUSPROPERTYTAG.Unknown;
		private static Dictionary<string, object> m_metaData = new Dictionary<string, object>();
		private static string m_volume = string.Empty;
		private static string m_mute = string.Empty;

		[ComVisible(false)]
		public enum MEDIASTATUSPROPERTYTAG {
			Application = 0xf001,
			ArtistName = 0x2018,
			CallingPartyName = 0x2029,
			CallingPartyNumber = 0x2028,
			CD = 0x2002,
			CurrentPicture = 0x201b,
			DiscWriter_ProgressPercentageChanged = 0x2030,
			DiscWriter_ProgressTimeChanged = 0x202f,
			DiscWriter_SelectedFormat = 0x202e,
			DiscWriter_Start = 0x202d,
			DiscWriter_Stop = 0x2031,
			DVD = 0x2001,
			Ejecting = 0x1010,
			DialogVisible = 0x100f, // Was Error
			FF1 = 0x100a,
			FF2 = 0x100b,
			FF3 = 0x100c,
			FS_DVD = 0x2010,
			FS_Extensibility = 0x202c,
			FS_Guide = 0x2011,
			FS_Home = 0x200e,
			FS_Music = 0x2012,
			FS_Photos = 0x2013,
			FS_Radio = 0x2025,
			FS_RecordedShows = 0x2015,
			FS_TV = 0x200f,
			FS_Unknown = 0x2016,
			FS_Videos = 0x2014,
			GuideLoaded = 0x201d,
			MediaName = 0x2017,
			MediaTime = 0x2007,
			MediaTypes = 0x2000,
			MSASPrivateTags = 0xf000,
			Mute = 0x1000,
			Next = 0x100d,
			NextFrame = 0x2021,
			ParentalAdvisoryRating = 0x202a,
			Pause = 0x1002,
			PhoneCall = 0x2027,
			Photos = 0x201a,
			Play = 0x1001,
			Prev = 0x100e,
			PrevFrame = 0x2022,
			PVR = 0x2003,
			Radio = 0x2023,
			RadioFrequency = 0x2024,
			Recording = 0x1006,
			RepeatSet = 0x1005,
			RequestForTuner = 0x202b,
			Rewind1 = 0x1007,
			Rewind2 = 0x1008,
			Rewind3 = 0x1009,
			Shuffle = 0x1004,
			SlowMotion1 = 0x201e,
			SlowMotion2 = 0x201f,
			SlowMotion3 = 0x2020,
			Stop = 0x1003,
			StreamingContentAudio = 0x2004,
			StreamingContentVideo = 0x2005,
			TitleNumber = 0x200c,
			TotalTracks = 0x2009,
			TrackDuration = 0x200a,
			TrackName = 0x2019,
			TrackNumber = 0x2008,
			TrackTime = 0x200b,
			TransitionTime = 0x201c,
			TVTuner = 0x2006,
			Unknown = 0,
			Visualization = 0x2026,
			Volume = 0x200d
		}

		/// <summary>
		/// Gets the volume.
		/// </summary>
		/// <value>The volume.</value>
		public static string Volume {
			get { return MediaState.m_volume; }
		}

		/// <summary>
		/// Gets the mute.
		/// </summary>
		/// <value>The mute.</value>
		public static string Mute {
			get { return MediaState.m_mute; }
		}

		/// <summary>
		/// Gets the meta data.
		/// </summary>
		/// <value>The meta data.</value>
		public static Dictionary<string, object> MetaData {
			get { return MediaState.m_metaData; }
		}

		/// <summary>
		/// Gets the current page.
		/// </summary>
		/// <value>The page.</value>
		public static MEDIASTATUSPROPERTYTAG Page {
			get { return MediaState.m_page; }
		}

		/// <summary>
		/// Gets the current play rate.
		/// </summary>
		/// <value>The play rate.</value>
		public static MEDIASTATUSPROPERTYTAG PlayRate {
			get { return MediaState.m_playRate; }
		}

		/// <summary>
		/// Gets the current media mode.
		/// </summary>
		/// <value>The media mode.</value>
		public static MEDIASTATUSPROPERTYTAG MediaMode {
			get { return MediaState.m_mediaMode; }
		}

		/// <summary>
		/// Updates the state.
		/// </summary>
		/// <param name="tag">The tag.</param>
		/// <param name="property">The property.</param>
		public static void UpdateState(MEDIASTATUSPROPERTYTAG tag, object property) {
			if (property == null) {
				return;
			}

			switch (tag) {
					//  Volume
				case MEDIASTATUSPROPERTYTAG.Volume:
					m_volume = property.ToString();
					break;

					//  Mute
				case MEDIASTATUSPROPERTYTAG.Mute:
					m_mute = property.ToString();
					break;

					//  Current Navigation Page
				case MEDIASTATUSPROPERTYTAG.FS_DVD:
				case MEDIASTATUSPROPERTYTAG.FS_Guide:
				case MEDIASTATUSPROPERTYTAG.FS_Home:
				case MEDIASTATUSPROPERTYTAG.FS_Music:
				case MEDIASTATUSPROPERTYTAG.FS_Photos:
				case MEDIASTATUSPROPERTYTAG.FS_Radio:
				case MEDIASTATUSPROPERTYTAG.FS_RecordedShows:
				case MEDIASTATUSPROPERTYTAG.FS_TV:
				case MEDIASTATUSPROPERTYTAG.FS_Unknown:
				case MEDIASTATUSPROPERTYTAG.FS_Videos:
					if ((bool) property == true) {
						m_page = tag;
					}
					break;

					//  Play Rates
				case MEDIASTATUSPROPERTYTAG.Play:
				case MEDIASTATUSPROPERTYTAG.Stop:
				case MEDIASTATUSPROPERTYTAG.Pause:
				case MEDIASTATUSPROPERTYTAG.FF1:
				case MEDIASTATUSPROPERTYTAG.FF2:
				case MEDIASTATUSPROPERTYTAG.FF3:
				case MEDIASTATUSPROPERTYTAG.Rewind1:
				case MEDIASTATUSPROPERTYTAG.Rewind2:
				case MEDIASTATUSPROPERTYTAG.Rewind3:
				case MEDIASTATUSPROPERTYTAG.SlowMotion1:
				case MEDIASTATUSPROPERTYTAG.SlowMotion2:
				case MEDIASTATUSPROPERTYTAG.SlowMotion3:
					if ((bool) property == true) {
						m_playRate = tag;
					}
					break;

					//  Current Media Mode
				case MEDIASTATUSPROPERTYTAG.StreamingContentAudio:
				case MEDIASTATUSPROPERTYTAG.StreamingContentVideo:
				case MEDIASTATUSPROPERTYTAG.PVR:
				case MEDIASTATUSPROPERTYTAG.TVTuner:
				case MEDIASTATUSPROPERTYTAG.CD:
				case MEDIASTATUSPROPERTYTAG.DVD:
					if ((bool) property == true) {
						m_metaData.Clear();
						m_mediaMode = tag;
					}
					break;

				default:
					m_metaData[tag.ToString()] = property;
					break;
			}
		}
	}
}