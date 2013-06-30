/*
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 */
using System;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Collections.Generic;
using System.Threading;
using System.IO;

namespace VmcController.MceState {
	/// <summary>
	/// Provides a generic socket server
	/// </summary>
	public class TcpSocketServer {
		#region Member Variables

		private TcpListener m_tcpListener;
		private List<TcpClient> m_tcpClients = new List<TcpClient>();
		private int m_listenPort;
		private object m_threadLock = new object();

		#endregion

		#region Public Events

		/// <summary>
		/// Fired when a client first connects
		/// </summary>
		public event EventHandler<SocketEventArgs> Connected;

		/// <summary>
		/// Fired each time a client sends a message
		/// </summary>
		public event EventHandler<SocketEventArgs> NewMessage;

		/// <summary>
		/// Fired when a client disconnects
		/// </summary>
		public event EventHandler<SocketEventArgs> Disconnected;

		#endregion

		#region Constructor

		/// <summary>
		/// Initializes a new instance of the <see cref="SocketServer"/> class.
		/// </summary>
		public TcpSocketServer() {}

		#endregion

		#region Public Methods

		/// <summary>
		/// Create and starts the listening socket
		/// </summary>
		/// <param name="portNum">The port num.</param>
		public void StartListening(int portNum) {
			// Create the listening socket...
			m_tcpListener = new TcpListener(IPAddress.Any, portNum);
			// Start listening...
			m_tcpListener.Start(4);
			// Create the call back for any client connections...
			m_tcpListener.BeginAcceptTcpClient(new AsyncCallback(OnClientConnection), null);
			// Store the port number
			m_listenPort = portNum;
		}

		/// <summary>
		/// Returns the port the server is listening on
		/// </summary>
		public int PortNumber {
			get { return m_listenPort; }
		}

		/// <summary>
		/// Gets the count of connections.
		/// </summary>
		/// <value>The count.</value>
		public int Count {
			get { return m_tcpClients.Count; }
		}

		/// <summary>
		/// Sends the message to all the sockets
		/// </summary>
		/// <param name="text">The text.</param>
		public void SendMessage(string text) {
			byte[] byData = System.Text.Encoding.ASCII.GetBytes(text);
			foreach (TcpClient tcpClient in m_tcpClients) {
				try {
					if (tcpClient != null && tcpClient.Client.Connected) {
						tcpClient.GetStream().Write(byData, 0, byData.Length);
					}
				}
				catch (IOException) {
					CloseClient(tcpClient);
				}
			}
		}

		/// <summary>
		/// Sends the message to a specific client
		/// </summary>
		/// <param name="text">The text.</param>
		public void SendMessage(string text, TcpClient tcpClient) {
			if (tcpClient == null) {
				throw new ArgumentNullException("tcpClient");
			}

			byte[] byData = System.Text.Encoding.ASCII.GetBytes(text);
			try {
				if (tcpClient.Client.Connected) {
					tcpClient.GetStream().Write(byData, 0, byData.Length);
				}
			}
			catch (IOException) {
				CloseClient(tcpClient);
			}
		}

		/// <summary>
		/// Closes all sockets
		/// </summary>
		public void StopListening() {
			if (m_tcpListener != null) {
				m_tcpListener.Stop();
			}

            closeClients();
		}

        public void closeClients()
        {
            foreach (TcpClient tcpClient in m_tcpClients)
            {
                if (tcpClient != null)
                {
                    CloseClient(tcpClient);
                }
            }
        }

		#endregion

		#region Private Methods

		/// <summary>
		/// This is the call back function, which will be invoked when a client is connected
		/// </summary>
		/// <param name="asyn">The asyn.</param>
		private void OnClientConnection(IAsyncResult asyn) {
			try {
				// Here we complete/end the BeginAccept() asynchronous call
				TcpClient tcpClient = m_tcpListener.EndAcceptTcpClient(asyn);

				// Add the new connection to our list
				lock (m_threadLock)
					m_tcpClients.Add(tcpClient);

				// Spawn a new thread to listen for data from this client connection
				ThreadPool.QueueUserWorkItem(new WaitCallback(ProcessClient), tcpClient);

				// Since the main Socket is now free, it can go back and wait for
				// other clients who are attempting to connect
				m_tcpListener.BeginAcceptTcpClient(new AsyncCallback(OnClientConnection), null);
			}
			catch (ObjectDisposedException) {
				System.Diagnostics.Debugger.Log(0, "1", "\n OnClientConnection: Socket has been closed\n");
			}
			catch (SocketException) {
				System.Diagnostics.Debugger.Log(0, "1", "\n OnClientConnection: Socket has been closed\n");
			}
		}

		/// <summary>
		/// Processes the client.
		/// </summary>
		/// <param name="asyn">The asyn.</param>
		private void ProcessClient(object tcpClient) {
			TcpClient newClient = (TcpClient) tcpClient;

			try {
				if (Connected != null) {
					Connected(this, new SocketEventArgs(string.Empty, newClient));
				}
				using (StreamReader streamReader = new StreamReader(newClient.GetStream())) {
					while (newClient.Client.Connected) {
						string line = streamReader.ReadLine();
						if (line == null) {
							break;
						}
						else if (NewMessage != null) {
							NewMessage(this, new SocketEventArgs(line.TrimEnd(), newClient));
						}
					}
				}
			}
			catch (IOException) {}
			if (Disconnected != null) {
				Disconnected(this, new SocketEventArgs(string.Empty, newClient));
			}
			CloseClient(newClient);
		}

		/// <summary>
		/// Closes the network stream and client socket.
		/// </summary>
		/// <param name="client">The client.</param>
		private void CloseClient(TcpClient tcpClient) {
			if (tcpClient != null) {
				lock (m_threadLock)
					m_tcpClients.Remove(tcpClient);

				if (tcpClient.Connected) {
					tcpClient.GetStream().Close();
				}
				tcpClient.Close();
			}
		}

		#endregion
	}

	/// <summary>
	/// EventHandler Class
	/// </summary>
	public class SocketEventArgs : EventArgs {
		private string m_message;
		private TcpClient m_tcpClient;

		/// <summary>
		/// Initializes a new instance of the <see cref="SocketEventArgs"/> class.
		/// </summary>
		/// <param name="message">The message.</param>
		/// <param name="client">The client.</param>
		public SocketEventArgs(string message, TcpClient client) {
			m_message = message;
			m_tcpClient = client;
		}

		/// <summary>
		/// Gets the message.
		/// </summary>
		/// <value>The message.</value>
		public string Message {
			get { return m_message; }
		}

		/// <summary>
		/// Gets the TCP client.
		/// </summary>
		/// <value>The TCP client.</value>
		public TcpClient TcpClient {
			get { return m_tcpClient; }
		}
	}
}