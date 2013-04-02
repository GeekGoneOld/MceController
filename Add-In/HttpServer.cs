/*
 * This module is loosly based on J.Bradshaw's TcpSocketServer.cs
 * Implements a simple http server on m_basePortNumber+10
 * 
 * Copyright (c) 2009 Anthony Jones
 * 
 * Portions copyright (c) 2007 Jonathan Bradshaw
 * 
 * This software code module is provided 'as-is', without any express or implied warranty. 
 * In no event will the authors be held liable for any damages arising from the use of this software.
 * Permission is granted to anyone to use this software for any purpose, including commercial 
 * applications, and to alter it and redistribute it freely.
 * 
 * History:
 * 2009-01-21 Created Anthony Jones
 * 
 */
using System;	
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Threading ;
using System.Web;


namespace VmcController.AddIn
{
    class HttpSocketServer 
	{

        #region Member Variables
        private TcpListener m_httpListener;
        private int m_listenPort = 40510; // Select any free port you wish
        private string m_sHttpVersion;
        private object m_threadLock = new object();
        private List<Socket> m_httpSocket = new List<Socket>();
        #endregion

        #region Public Events
        /// <summary>
        /// Fired when a http request is made
        /// </summary>
        public event EventHandler<HttpEventArgs> NewRequest;
        #endregion


        public void StartListening (int portNum)
        {
            // Create the listening socket...
            m_httpListener = new TcpListener(IPAddress.Any, portNum);
            // Start listening...
            m_httpListener.Start(4);
            // Create the call back for any client connections...
            m_httpListener.BeginAcceptSocket(new AsyncCallback(OnClientConnection), null);
            // Store the port number
            m_listenPort = portNum;
        }

        /// <summary>
        /// This is the call back function, which will be invoked when a client is connected
        /// </summary>
        /// <param name="asyn">The asyn.</param>
        private void OnClientConnection (IAsyncResult asyn)
        {
            int iStartPos = 0;
            String sRequest;

            try
            {
                // Here we complete/end the BeginAccept() asynchronous call
                Socket mySocket = m_httpListener.EndAcceptSocket(asyn);

                //Accept a new connection
                //Socket mySocket = m_httpListener.AcceptSocket();

                //make a byte array and receive data from the client 
                Byte[] bReceive = new Byte[1024];
                int i = mySocket.Receive(bReceive, bReceive.Length, 0);

                //Convert Byte to String
                //string sBuffer = Encoding.ASCII.GetString(bReceive);
                string sBuffer = Encoding.UTF8.GetString(bReceive);

                //At present we will only deal with GET type
                if (sBuffer.Substring(0, 3) != "GET")
                {
                    mySocket.Close();
                }
                else
                {

                    // Look for HTTP request
                    iStartPos = sBuffer.IndexOf("HTTP", 1);

                    // Get the HTTP text and version e.g. it will return "HTTP/1.1"
                    m_sHttpVersion = sBuffer.Substring(iStartPos, 8);

                    // Extract the Requested Type and Requested file/directory
                    sRequest = sBuffer.Substring(5, iStartPos - 1 - 5);

                    sRequest = HttpUtility.UrlDecode(sRequest);
                    if (NewRequest != null)
                        NewRequest(this, new HttpEventArgs(sRequest.TrimEnd(), mySocket));
                }

                // Since the main Socket is now free, it can go back and wait for
                // other clients who are attempting to connect
                m_httpListener.BeginAcceptSocket(new AsyncCallback(OnClientConnection), null);
            }
            catch (ObjectDisposedException)
            {
                System.Diagnostics.Debugger.Log(0, "1", "\n OnClientConnection: Socket has been closed\n");
            }
            catch (SocketException)
            {
                System.Diagnostics.Debugger.Log(0, "1", "\n OnClientConnection: Socket has been closed\n");
            }
        }




        /// <summary>
        /// Sends the generated page
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="httpSocket">Socket reference</param>
        public void SendPage (string text, Socket httpSocket)
        {
            if (httpSocket == null)
                throw new ArgumentNullException("httpSocket");

            //byte[] byData = System.Text.Encoding.ASCII.GetBytes(text);
            byte[] byData = System.Text.Encoding.UTF8.GetBytes(text);
            try
            {
                if (httpSocket.Connected)
                {
                    SendHeader(m_sHttpVersion, "text/html;charset=utf-8", byData.Length, " 200 OK", ref httpSocket);
                    SendToBrowser(byData, ref httpSocket);
                    httpSocket.Close();
                }
            }
            catch (IOException)
            {
                httpSocket.Close();
            }
        }


        /// <summary>
        /// Sends the served image
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="mime_ext">The mime extention.</param>
        /// <param name="httpSocket">Socket reference</param>
        public void SendImage(string text, string mime_ext, Socket httpSocket)
        {
            if (httpSocket == null)
                throw new ArgumentNullException("httpSocket");

            byte[] byData = Convert.FromBase64String(text);
            try
            {
                if (httpSocket.Connected)
                {
                    if (mime_ext == "jpg") mime_ext = "jpeg";
                    SendHeader(m_sHttpVersion, "image/" + mime_ext, byData.Length, " 200 OK", ref httpSocket);
                    SendToBrowser(byData, ref httpSocket);
                    httpSocket.Close();
                }
            }
            catch (IOException)
            {
                httpSocket.Close();
            }
        }


        /// <summary>
        /// This function sends the basic header Information to the client browser
        /// </summary>
        /// <param name="sHttpVersion">HTTP version</param>
        /// <param name="sMIMEHeader">Mime type</param>
        /// <param name="iTotBytes">Total bytes that will be sent</param>
        /// <param name="mySocket">Socket reference</param>
        /// <returns></returns>
        public void SendHeader (string sHttpVer, string sMimeType, int iBytesCount, string sStatusCode, ref Socket httpSocket)
        {

            String sBuffer = "";

            // if Mime type is not provided set default to text/html
            if (sMimeType.Length == 0)
            {
                sMimeType = "text/html";  // Default Mime Type is text/html
            }

            sBuffer = sBuffer + sHttpVer + sStatusCode + "\r\n";
            sBuffer = sBuffer + "Server: cx1193719-b\r\n";
            sBuffer = sBuffer + "Content-Type: " + sMimeType + "\r\n";
            sBuffer = sBuffer + "Accept-Ranges: bytes\r\n";
            sBuffer = sBuffer + "Content-Length: " + iBytesCount + "\r\n\r\n";

            Byte[] bSendData = Encoding.ASCII.GetBytes(sBuffer);

            SendToBrowser(bSendData, ref httpSocket);

        }

        /// <summary>
        /// Overloaded Function, takes string, convert to bytes and calls 
        /// overloaded sendToBrowserFunction.
        /// </summary>
        /// <param name="sData">The data to be sent to the browser(client)</param>
        /// <param name="mySocket">Socket reference</param>
        public void SendToBrowser (String sData, ref Socket httpSocket)
        {
            SendToBrowser(Encoding.ASCII.GetBytes(sData), ref httpSocket);
        }



        /// <summary>
        /// Sends data to the browser (client)
        /// </summary>
        /// <param name="bSendData">Byte Array</param>
        /// <param name="mySocket">Socket reference</param>
        public void SendToBrowser (Byte[] bSendData, ref Socket mySocket)
        {
            int numBytes = 0;

            try
            {
                if (mySocket.Connected)
                {
                    numBytes = mySocket.Send(bSendData, bSendData.Length, 0);
                    //if (numBytes == -1)
                    //    EventLog.WriteEntry("VmcController Client httpServer", "Socket Error cannot Send Packet", EventLogEntryType.Information);
                    //else
                    //    EventLog.WriteEntry("VmcController Client httpServer", String.Format("Number of bytes sent: {0}", numBytes), EventLogEntryType.Information);
                }
                else
                    EventLog.WriteEntry("VmcController Client httpServer", "Connection Dropped....", EventLogEntryType.Information);
            }
            catch (Exception e)
            {
                EventLog.WriteEntry("VmcController Client httpServer", String.Format("Error Occurred : {0} ", e), EventLogEntryType.Information);

            }
        }

	}


    /// <summary>
    /// EventHandler Class
    /// </summary>
    public class HttpEventArgs : EventArgs
    {
        private string m_httpRequest;
        private Socket m_httpSocket;

        /// <summary>
        /// Initializes a new instance of the <see cref="SocketEventArgs"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="client">The client.</param>
        public HttpEventArgs (string message, Socket httpSocket)
        {
            m_httpRequest = message;
            m_httpSocket = httpSocket;
        }

        /// <summary>
        /// Gets the message.
        /// </summary>
        /// <value>The message.</value>
        public string Request
        {
            get { return m_httpRequest; }
        }

        /// <summary>
        /// Gets the http Socket.
        /// </summary>
        /// <value>The http socket.</value>
        public Socket HttpSocket
        {
            get { return m_httpSocket; }
        }
    }
}
