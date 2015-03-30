/*
 * New MediaStatusSink implementation. 
 * 
 * Based on Jonathan Bradshaw's implementation
 * 
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using VmcController.MceState;
using System.IO;

namespace VmcController.Status {
	[ComVisible(true)]
	[GuidAttribute("38751110-99ac-4e25-b2f3-6f5e48428ebf")]
	public class Sink : IMediaStatusSink {
		private const int BASE_PORT = 40400;
		private static int listeningPortNumber;
		private static KeyboardHook keyboardHook;

		public Sink() {
			var debugFile = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\vmccdebug.txt";

			if(File.Exists(debugFile)) {
				Trace.Listeners.Add(new TextWriterTraceListener(debugFile));
			} 

			Trace.AutoFlush = true;
			Trace.TraceInformation("Sink started");

			try {
				listeningPortNumber = GetPortNumberForUserType(BASE_PORT);
				keyboardHook = new KeyboardHook();
				SocketServer = new TcpSocketServer();

                //  Setup HTTP socket listener
                m_httpServer = new HttpSocketServer();
            }
			catch (Exception ex) {
				Trace.TraceError(ex.ToString());
				throw;
			}
			finally {
				Trace.Unindent();
				Trace.TraceInformation("MsasSink() End");
			}
		}

		#region IMediaStatusSink Members

		public void Initialize() {
			Trace.TraceInformation("Sink.Initialize called");

			try {
				SocketServer.StartListening(listeningPortNumber);
				SocketServer.Connected += SocketConnected;
				keyboardHook.KeyPress += KeyPressEvent;
                m_httpServer.InitServer();
                m_httpServer.StartListening(listeningPortNumber + 10);
            }
			catch (Exception ex) {
				Trace.TraceError(ex.ToString());
				throw;
			}
		}


		public IMediaStatusSession CreateSession() {
			Trace.TraceInformation("Sink.CreateSession called");
			try {
                SessionCount++;
                return new Session(SessionCount);
			}
			catch (Exception ex) {
				Trace.TraceError(ex.ToString());
				throw;
			}
			finally {
				Trace.Unindent();
				Trace.TraceInformation("CreateSession() End");
			}
		}
		#endregion

		private void SocketConnected(object sender, SocketEventArgs e) {
			var sb = new StringBuilder();
            string prefix;

            try
            {
                foreach (KeyValuePair<int, MediaState> mediaState in MediaStateDict.mediaStates)
                {
                    sb.AppendFormat(
                        "204 Connected (Build: {0} Clients: {1})\r\n",
                        GetVersionInfo, SocketServer.Count);

                    if (mediaState.Value.MediaMode == MediaState.MEDIASTATUSPROPERTYTAG.Recording)
                        prefix = "BG_";
                    else
                        prefix = "";

                    //  Provide current state information to the client
                    if (!string.IsNullOrEmpty(mediaState.Value.Volume))
                    {
                        sb.AppendFormat(prefix + "Volume={0}\r\n", mediaState.Value.Volume);
                    }
                    if (!string.IsNullOrEmpty(mediaState.Value.Mute))
                    {
                        sb.AppendFormat(prefix + "Mute={0}\r\n", mediaState.Value.Mute);
                    }
                    if (mediaState.Value.Page != MediaState.MEDIASTATUSPROPERTYTAG.Unknown)
                    {
                        sb.AppendFormat(prefix + "{0}=True\r\n", mediaState.Value.Page);
                        sb.AppendFormat(prefix + "m_page={0}\r\n", mediaState.Value.Page);
                    }
                    if (mediaState.Value.MediaMode != MediaState.MEDIASTATUSPROPERTYTAG.Unknown)
                    {
                        sb.AppendFormat(prefix + "{0}=True\r\n", mediaState.Value.MediaMode);
                        sb.AppendFormat(prefix + "m_mediaMode={0}\r\n", mediaState.Value.MediaMode);
                    }
                    if (mediaState.Value.PlayRate != MediaState.MEDIASTATUSPROPERTYTAG.Unknown)
                    {
                        sb.AppendFormat(prefix + "{0}=True\r\n", mediaState.Value.PlayRate);
                        sb.AppendFormat(prefix + "m_playRate={0}\r\n", mediaState.Value.PlayRate);
                    }
                    foreach (KeyValuePair<string, object> item in mediaState.Value.MetaData)
                        sb.AppendFormat(prefix + "{0}={1}\r\n", item.Key, item.Value);

                    //  Send the data to the connected client
                    Trace.TraceInformation(sb.ToString());
                    SocketServer.SendMessage(sb.ToString(), e.TcpClient);
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
            }
            finally
            {
                Trace.Unindent();
                Trace.TraceInformation("socketServer_Connected() End");
            }
		}

		private void KeyPressEvent(object sender, KeyboardHookEventArgs e) {
			try {
				//  Send the data to the clients
				SocketServer.SendMessage(
					String.Format(CultureInfo.InvariantCulture, "KeyPress={0}\r\n", e.vkCode)
				);
			}
			catch (Exception ex) {
				Trace.TraceError(ex.ToString());
			}
			finally {
				Trace.Unindent();
				Trace.TraceInformation("keyboardHook_KeyPress() End");
			}
		}

		/// <summary>
		/// Gets the TCP socket server.
		/// </summary>
		/// <value>The TCP socket server.</value>
		public static TcpSocketServer SocketServer { get; private set; }

        /// <summary>
        /// HTTP Socket Server
        /// </summary>
        private HttpSocketServer m_httpServer;

		/// <summary>
		/// Gets the session count.
		/// </summary>
		/// <value>The session count.</value>
		public static int SessionCount { get; private set; }

        /// <summary>
		/// Determine what TCP port to listen on. Current Session listens on BASE_PORT
		/// Extender Sessions listens on BASE_PORT + Extender ID
		/// </summary>
		/// <param name="basePort">The base port.</param>
		/// <returns>port number</returns>
		private static int GetPortNumberForUserType(int basePort) {
			var curPrincipal = System.Security.Principal.WindowsIdentity.GetCurrent();
			int port = basePort;

			try {
				if (curPrincipal != null) {
					var principalName = curPrincipal.Name;
					var sessionId = Process.GetCurrentProcess().SessionId;

					Trace.TraceInformation("Windows Session #{0} Identity: {1}", sessionId, principalName);

					if (principalName.IndexOf("Mcx") > 0 && sessionId != 1) {
						//Max. of 5 extenders for MC. read the first digit after Mcx
						var extenderString = principalName.Substring(principalName.LastIndexOf("Mcx") + 3, 1);

						port = basePort + int.Parse(extenderString, CultureInfo.InvariantCulture);
					}

					if (sessionId == 1) {
						port = basePort;
					}
				}
				else {
					Trace.TraceError("Principal is null! Exit.");
				}				
			}
			catch (Exception ex) {
				Trace.TraceInformation(ex.Message);
				throw new InvalidOperationException("Unable to determine correct port number");
			}

			return port;
		}

		public static string GetVersionInfo {
			get { return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(3); }
		}
	}
}
