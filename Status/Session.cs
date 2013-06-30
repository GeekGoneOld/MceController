/*
 * New MediaStatusSession implementation. 
 * 
 * Based on Jonathan Bradshaw's
 * 
 */

using System;
using System.Diagnostics;
using System.Globalization;
using VmcController.MceState;

namespace VmcController.Status {
	class Session : IMediaStatusSession {
		private readonly int sessionCounter;
		private MediaState.MEDIASTATUSPROPERTYTAG prevCheckTag;
		private object prevCheckProp;

		public Session(int count) {
			sessionCounter = count;

			if (Sink.SocketServer != null) {
				Sink.SocketServer.SendMessage(
					string.Format(CultureInfo.InvariantCulture, "StartSession={0}\r\n", sessionCounter)
				);
			}
			else {
				Trace.TraceWarning("SocketServer reference is null");
			}
		}


		#region IMediaStatusSession Members

		public void MediaStatusChange(MediaState.MEDIASTATUSPROPERTYTAG[] Tags, object[] Properties) {
			Trace.TraceInformation("Session.MediaStatusChange called");

			//Check to see if the socket server is valid
			if (Sink.SocketServer == null) {
			  Trace.TraceError("SocketServer is null");
			  return;
			}

			try {
				for (int i = 0; i < Tags.Length; i++) {
					var tag = (MediaState.MEDIASTATUSPROPERTYTAG)Tags.GetValue(i);
					var value = Properties.GetValue(i);

					//  Remove duplicate notifications
					if ((tag == prevCheckTag) && (value.ToString() == prevCheckProp.ToString())) {
						return;
					}

					prevCheckTag = tag;
					prevCheckProp = value;

					Trace.TraceInformation("Tag  {0}={1}", tag, value);

					Sink.SocketServer.SendMessage(
						string.Format(CultureInfo.InvariantCulture, "{0}={1}\r\n", tag, value)
						);

					MediaState.UpdateState(tag, value);
				}
			}
			catch (Exception ex) {
				Trace.TraceError("Exception: {0}", ex.ToString());
			}
		}

		public void Close() {
			Trace.TraceInformation("Closing media session #{0}", sessionCounter);

			if (Sink.SocketServer != null) {
				Sink.SocketServer.SendMessage(
					string.Format(CultureInfo.InvariantCulture, "EndSession={0}\r\n", sessionCounter)
				);
                Sink.SocketServer.closeClients();
			}
		}
		
		#endregion
	}
}
