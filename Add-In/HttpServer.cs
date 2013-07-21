/*
 * This module is loosly based on J.Bradshaw's TcpSocketServer.cs
 * Implements a simple http server on m_basePortNumber+10
 * 
 * Copyright (c) 2013 Skip Mercier
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
 * 2013 Added code to use remoted WMP instance 
 * 2013-06-07 Converted to AsyncSocketEvenArgs based server with preallocated buffer
 * 
 */
using System;	
using System.IO;
using System.Timers;
using System.Net;
using System.Net.Sockets;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Threading ;
using System.Web;
using VmcController.AddIn.Commands;
using System.Xml;
using System.Reflection;
using System.Runtime.InteropServices;
using WMPLib;
using Newtonsoft.Json;
using VmcController.AddIn.Metadata;


namespace VmcController.AddIn
{
    public class HttpSocketServer : IDisposable
	{
        #region Member Variables

        public const int MIN_SEND_BUFFER_SIZE = 3000000;

        // the maximum number of connections the sample is designed to handle simultaneously 
        private int maxConnections = 4;

        // this variable allows us to create some extra SAEA objects for the pool,
        // if we wish.
        private int numberOfEventArgsForRecSend = 6;

        // max # of pending connections the listener can hold in queue
        private int backlog = 3;

        // tells us how many objects to put in pool for accept operations
        private int maxSimultaneousAcceptOps = 3;

        // buffer size to use for each socket receive operation
        private int receiveBufferSize = 500;

        // the total number of clients connected to the server 
        int numConnectedSockets;

        private string sHttpVersion;

        RemoteCommands remoteCommands = new RemoteCommands();
        ServerSettings settings = new ServerSettings();

        //Writes logs to AddInModule.DATA_DIR
        Logger logger;

        // represents a large reusable set of buffers for all socket operations
        BufferManager bufferManager;

        // the socket used to listen for incoming connection requests
        Socket listenSocket;

        // pool of reusable SocketAsyncEventArgs objects for accept operations
        SocketAsyncEventArgsPool poolOfAcceptEventArgs;
        
        // pool of reusable SocketAsyncEventArgs objects for receive and send socket operations
        SocketAsyncEventArgsPool poolOfRecSendEventArgs;

        //A Semaphore has two parameters, the initial number of available slots
        // and the maximum number of slots. We'll make them the same. 
        //This Semaphore is used to keep from going over max connection #. (It is not about 
        //controlling threading really here.)  
        Semaphore maxNumberAcceptedClients;

        AutoResetEvent waitHandle = new AutoResetEvent(true);

        //Variables below represent those related to regulating access to cache and the creation schedule
        System.Timers.Timer cacheTimer;

        protected bool _disposed;

                                  
        #endregion


        public HttpSocketServer()
        {
            _disposed = false;

            logger = new Logger("HttpServer", false);
            logger.Write("Loading settings...");

            InitSettings();

            numConnectedSockets = 0;                                   

            bufferManager = new BufferManager((receiveBufferSize + settings.send_buffer_size) * numberOfEventArgsForRecSend, receiveBufferSize + settings.send_buffer_size);

            poolOfAcceptEventArgs = new SocketAsyncEventArgsPool(maxSimultaneousAcceptOps);
            poolOfRecSendEventArgs = new SocketAsyncEventArgsPool(numberOfEventArgsForRecSend);
            maxNumberAcceptedClients = new Semaphore(maxConnections, maxConnections); 
        }

        private XmlDocument getSettingsDoc()
        {
            try
            {
                if (!Directory.Exists(AddInModule.DATA_DIR))
                {
                    Directory.CreateDirectory(AddInModule.DATA_DIR);
                }
                XmlDocument doc = new XmlDocument();
                if (System.IO.File.Exists(AddInModule.DATA_DIR + "\\settings.xml"))
                {
                    //Settings found, load if exist and add if not
                    doc.Load(AddInModule.DATA_DIR + "\\settings.xml");
                    return doc;
                }
                else
                {
                    //Create settings for the first time
                    doc.LoadXml("<settings><sendBuffer>" + Convert.ToString(settings.send_buffer_size) + "</sendBuffer><lastCacheTime>" + DateTime.Now + "</lastCacheTime>"
                        + "<cacheStartHour>4</cacheStartHour></settings>");
                    doc.PreserveWhitespace = true;
                    doc.Save(AddInModule.DATA_DIR + "\\settings.xml");
                    return doc;
                }
            }
            catch (Exception e)
            {
                logger.Write("Error loading settings.xml, " + e);                
            }
            return null;
        }

        private void InitSettings()
        {            
            XmlDocument doc = getSettingsDoc();
            if (doc != null)
            {
                XmlNode sendBufferNode = doc.DocumentElement.SelectSingleNode("sendBuffer");
                if (sendBufferNode != null)
                {
                    int buffer = Convert.ToInt32(sendBufferNode.InnerText);
                    if (buffer > settings.send_buffer_size) settings.send_buffer_size = buffer;
                    else sendBufferNode.InnerText = Convert.ToString(settings.send_buffer_size);
                }
                int startHour = 4;
                XmlNode cacheStartHourNode = doc.DocumentElement.SelectSingleNode("cacheStartHour");
                if (cacheStartHourNode != null)
                {
                    startHour = Convert.ToInt32(cacheStartHourNode.InnerText);
                }
                else
                {
                    //Set 4AM as the cache start time
                    XmlElement elem = doc.CreateElement("cacheStartHour");
                    XmlText text = doc.CreateTextNode("4");
                    doc.DocumentElement.AppendChild(elem);
                    doc.DocumentElement.LastChild.AppendChild(text);
                    //Save settings
                    doc.PreserveWhitespace = true;
                    doc.Save(AddInModule.DATA_DIR + "\\settings.xml");
                }

                logger.Write("Cache building is set for hour " + startHour);
                setCacheBuildTimer(startHour, doc);

                startCacheBuildConditionally(doc, false); 
            }
        }

        private OpResult setSendBufferSize(int size)
        {
            OpResult opResult = new OpResult();
            XmlDocument doc = getSettingsDoc();
            if (doc != null)
            {
                XmlNode sendBuffer = doc.DocumentElement.SelectSingleNode("sendBuffer");
                if (sendBuffer != null)
                {
                    //Save settings
                    sendBuffer.InnerText = Convert.ToString(size);
                    doc.PreserveWhitespace = true;
                    doc.Save(AddInModule.DATA_DIR + "\\settings.xml");

                    opResult.StatusCode = OpStatusCode.Success;
                    opResult.StatusText = "Buffer size set to " + size + ". Media Center must be restarted for settings to take effect.";
                    return opResult;
                }
            }
            opResult.StatusCode = OpStatusCode.BadRequest;
            opResult.StatusText = "Error saving buffer setting!";           
            return opResult;
        }

        private OpResult setCacheBuildTimer(int startHour, XmlDocument doc)
        {
            OpResult opResult = new OpResult();
            if (doc != null)
            {
                XmlNode cacheStartHourNode = doc.DocumentElement.SelectSingleNode("cacheStartHour");
                if (cacheStartHourNode != null)
                {
                    settings.cache_build_hour = startHour;
                    //Save settings
                    cacheStartHourNode.InnerText = Convert.ToString(startHour);                    
                    doc.PreserveWhitespace = true;
                    doc.Save(AddInModule.DATA_DIR + "\\settings.xml");

                    //Set cache timer
                    DateTime now = DateTime.Now;
                    int hour;
                    int minutes;
                    if (startHour > now.Hour)
                    {
                        hour = startHour - now.Hour;
                        minutes = 60 - now.Minute;
                    }
                    else
                    {
                        //Round now hour up one then subtract to get to 24 hour, then add to cacheStartHour
                        hour = 24 - now.Hour + 1 + startHour;
                        minutes = 60 - now.Minute;
                    }
                    if (cacheTimer == null) cacheTimer = new System.Timers.Timer();
                    else cacheTimer.Stop();
                    cacheTimer.AutoReset = false;
                    cacheTimer.Interval = 1000 * 60 * (60 * hour + minutes);
                    cacheTimer.Enabled = true;
                    cacheTimer.Elapsed += new ElapsedEventHandler(TimerElapsedEvent);

                    logger.Write("Cache timer set for " + hour + " hours and " + minutes + " minutes");

                    opResult.StatusCode = OpStatusCode.Success;
                    opResult.StatusText = "Cache start hour set to " + startHour;
                    return opResult;
                }                
            }
            opResult.StatusCode = OpStatusCode.BadRequest;
            opResult.StatusText = "Error saving cache start hour setting!";
            return opResult;
        }

        public OpResult ExecuteSettingsCommand(string command, string sParam)
        {
            OpResult opResult = new OpResult();
            try
                {
                    //Check if set-hour param is present
                    int value = GetServerSetting(sParam);
                    if (value == -1)
                    {
                        return remoteCommands.ExecuteServerSettings(command, settings);
                    }
                    else if (sParam.Contains("set-build-hour:"))
                    {
                        return setCacheBuildTimer(value, getSettingsDoc());
                    }
                    else if (sParam.Contains("set-send-buffer:"))
                    {
                        return setSendBufferSize(value);
                    }
                }
                catch (Exception ex)
                {
                    opResult.StatusCode = OpStatusCode.Exception;
                    opResult.StatusText = "Value out of range: " + ex.Message;
                }
            return opResult;
        }

        /// <summary>
        /// Checks to see if there is a settings parameter and
        /// returns -1 if none
        /// </summary>
        public int GetServerSetting(string param)
        {
            if (param != null)
            {
                string setter = null;

                if (param.Contains("set-build-hour:"))
                {
                    setter = "set-build-hour:";
                }
                else if (param.Contains("set-send-buffer:"))
                {
                    setter = "set-send-buffer:";
                }

                if (setter != null)
                {
                    string preString = param.Substring(param.IndexOf(setter) + setter.Length);
                    string rawString = MusicCmd.trim_parameter(preString);
                    int setting = Convert.ToInt32(rawString);

                    if (param.Contains("set-build-hour:") && setting >= 0 && setting < 24)
                    {
                        return setting;
                    }
                    else if (param.Contains("set-send-buffer:") && setting > HttpSocketServer.MIN_SEND_BUFFER_SIZE)
                    {
                        return setting;
                    }
                    else
                    {
                        throw new Exception();
                    }
                }
                return -1;
            }
            return -1;
        }

        /// <summary>
        /// Initializes the server by preallocating reusable buffers and context objects.  These objects do not 
        /// need to be preallocated or reused, by is done this way to illustrate how the API can easily be used
        /// to create reusable objects to increase server performance.
        /// </summary>
        public void InitServer()
        {
            // Allocates one large byte buffer which all I/O operations use a piece of.  This gaurds 
            // against memory fragmentation
            bufferManager.InitBuffer();

            //Allocate the pool of accept operation EventArgs objects
            //which do not need a buffer
            for (int i = 0; i < maxSimultaneousAcceptOps; i++)
            {
                SocketAsyncEventArgs acceptEventArgs = new SocketAsyncEventArgs();
                acceptEventArgs.Completed += new EventHandler<SocketAsyncEventArgs>(AcceptEventArg_Completed);
                
                poolOfAcceptEventArgs.Push(acceptEventArgs);
            }

            //Allocate the pool of receive/send operation EventArgs objects
            //which DO need a buffer
            for (int i = 0; i < numberOfEventArgsForRecSend; i++)
            {
                SocketAsyncEventArgs sendReceiveEventArgs = new SocketAsyncEventArgs();
                bufferManager.SetBuffer(sendReceiveEventArgs);

                sendReceiveEventArgs.Completed += new EventHandler<SocketAsyncEventArgs>(IO_Completed);

                DataHolderToken token = new DataHolderToken(sendReceiveEventArgs.Offset, receiveBufferSize);
                sendReceiveEventArgs.UserToken = token;

                poolOfRecSendEventArgs.Push(sendReceiveEventArgs);                
            }
        }

        /// <summary>
        /// Starts the server such that it is listening for incoming connection requests.    
        /// </summary>
        /// <param name="localEndPoint">The endpoint which the server will listening for conenction requests on</param>
        public void StartListening(int port)
        {
            //ipv4 endpoint
            IPEndPoint localEndPoint = new IPEndPoint(IPAddress.Any, port);

            // create the socket which listens for incoming connections
            listenSocket = new Socket(localEndPoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
            listenSocket.Bind(localEndPoint);

            // start the server with a listen backlog of 100 connections
            listenSocket.Listen(backlog);

            // post accepts on the listening socket
            StartAccept();
        }

        /// <summary>
        /// Begins an operation to accept a connection request from the client 
        /// </summary>
        public void StartAccept()
        {
            SocketAsyncEventArgs acceptEventArgs;

            //Get a SocketAsyncEventArgs object to accept the connection.                        
            //Get it from the pool if there is more than one in the pool.
            //We could use zero as bottom, but one is a little safer.            
            if (this.poolOfAcceptEventArgs.Count > 1)
            {
                try
                {
                    acceptEventArgs = this.poolOfAcceptEventArgs.Pop();
                }                
                catch
                {
                    //or make a new one.
                    acceptEventArgs = new SocketAsyncEventArgs();
                    acceptEventArgs.Completed += new EventHandler<SocketAsyncEventArgs>(AcceptEventArg_Completed);
                }
            }
            else
            {
                //or make a new one.
                acceptEventArgs = new SocketAsyncEventArgs();
                acceptEventArgs.Completed += new EventHandler<SocketAsyncEventArgs>(AcceptEventArg_Completed);
            }

            maxNumberAcceptedClients.WaitOne();

            bool willRaiseEvent = listenSocket.AcceptAsync(acceptEventArgs);
            if (!willRaiseEvent)
            {
                //Only called when connection will happen synchronously
                ProcessAccept(acceptEventArgs);
            }
        }

        /// <summary>
        /// This method is the callback method associated with Socket.AcceptAsync operations and is invoked
        /// when an accept operation is complete
        /// </summary>
        void AcceptEventArg_Completed(object sender, SocketAsyncEventArgs acceptEventArgs)
        {
            ProcessAccept(acceptEventArgs);
        }

        private void ProcessAccept(SocketAsyncEventArgs acceptEventArgs)
        {            
            // This is when there was an error with the accept operation which could
            // indicate a socket issue. This check prevents infinite loop if trying to use
            // same socket.
            if (acceptEventArgs.SocketError != SocketError.Success)
            {
                // Loop back to post another accept op. Notice that we are NOT
                // passing the SAEA object here.
                StartAccept();

                //Let's destroy this socket, since it could be bad.          
                acceptEventArgs.AcceptSocket.Close();

                //Put the SAEA back in the pool.
                poolOfAcceptEventArgs.Push(acceptEventArgs);

                maxNumberAcceptedClients.Release();
                //Jump out of the method.
                return;
            }
           
            Interlocked.Increment(ref numConnectedSockets);

            if (numConnectedSockets == maxConnections) logger.Write("Number of connected sockets: " + numConnectedSockets + " / " + maxConnections);

            // Loop back to post another accept op. Notice that we are NOT
            // passing the SAEA object here.
            StartAccept();

            // Get a SocketAsyncEventArgs object from the pool of receive/send objects
            SocketAsyncEventArgs receiveSendEventArgs = poolOfRecSendEventArgs.Pop();

            //Transfer the socket created by the AcceptAsync method to the 
            //args object that will do receive/send operations
            receiveSendEventArgs.AcceptSocket = acceptEventArgs.AcceptSocket;

            //Before putting the accept args back in the pool, set the socket to null
            acceptEventArgs.AcceptSocket = null;
            poolOfAcceptEventArgs.Push(acceptEventArgs);

            StartReceive(receiveSendEventArgs);
        }

        private void StartReceive(SocketAsyncEventArgs receiveSendEventArgs)
        {
            DataHolderToken token = (DataHolderToken)receiveSendEventArgs.UserToken;
            try
            {
                //Set the buffer for the receive operation.
                receiveSendEventArgs.SetBuffer(token.receiveBufferOffset, receiveBufferSize);

                // Post async receive operation on the socket.
                bool willRaiseEvent = receiveSendEventArgs.AcceptSocket.ReceiveAsync(receiveSendEventArgs);
                if (!willRaiseEvent)
                {
                    ProcessReceive(receiveSendEventArgs);
                }
            }
            catch (Exception e)
            {
                logger.Write("Exception in SetBuffer in StartReceive method of HttpServer.cs, offset=" + token.receiveBufferOffset +
                    ", count=" + receiveBufferSize + ", Buffer=" + receiveSendEventArgs.Buffer + ": " + e.Message); 
                token.Reset();
                CloseClientSocket(receiveSendEventArgs);
            }
        }

        // This method is called whenever a receive or send operation completes.
        void IO_Completed(object sender, SocketAsyncEventArgs e)
        {
            // determine which type of operation just completed and call the associated handler
            switch (e.LastOperation)
            {
                case SocketAsyncOperation.Receive:
                    ProcessReceive(e);
                    break;
                case SocketAsyncOperation.Send:
                    ProcessSend(e);
                    break;
                default:
                    //This exception will occur if you code the Completed event of some
                    //operation to come to this method, by mistake.
                    throw new ArgumentException("The last operation completed on the socket was not a receive or send");
            }
        }

        /// <summary>
        /// This method is invoked when an asynchronous receive operation completes. If the 
        /// remote host closed the connection, then the socket is closed.  If data was received then
        /// the data is echoed back to the client.
        /// </summary>
        private void ProcessReceive(SocketAsyncEventArgs receiveSendEventArgs)
        {
            DataHolderToken token = (DataHolderToken)receiveSendEventArgs.UserToken;

            //If there was a socket error, close the connection
            if (receiveSendEventArgs.SocketError != SocketError.Success)
            {
                token.Reset();
                CloseClientSocket(receiveSendEventArgs);

                //Jump out of the ProcessReceive method.
                return;
            }

            // If no data was received, close the connection. This is a NORMAL
            // situation that shows when the client has finished sending data.
            if (receiveSendEventArgs.BytesTransferred == 0)
            {
                token.Reset();
                CloseClientSocket(receiveSendEventArgs);
                return;
            }

            //Assume that all of the bytes transfered is the received message
            token.dataMessageReceived = new Byte[receiveSendEventArgs.BytesTransferred];
            Buffer.BlockCopy(receiveSendEventArgs.Buffer, token.receiveBufferOffset, token.dataMessageReceived, 0, receiveSendEventArgs.BytesTransferred);

            // Decode the byte array received in the token                 
            //string sBuffer = Encoding.ASCII.GetString(bReceive);
            string sBuffer = Encoding.UTF8.GetString(token.dataMessageReceived);

            //At present we will only deal with GET type
            if (sBuffer.Substring(0, 3) != "GET")
            {
                token.Reset();
                CloseClientSocket(receiveSendEventArgs);
            }
            else
            {
                // Look for HTTP request
                int iStartPos = sBuffer.IndexOf("HTTP", 1);

                // Get the HTTP text and version e.g. it will return "HTTP/1.1"
                sHttpVersion = sBuffer.Substring(iStartPos, 8);

                // Extract the Requested Type and Requested file/directory
                String sRequest = sBuffer.Substring(5, iStartPos - 1 - 5);

                sRequest = HttpUtility.UrlDecode(sRequest);

                token.httpRequest = sRequest.TrimEnd();

                StartSendRequest(receiveSendEventArgs);
            }         
        }

        // This method is called by I/O Completed() when an asynchronous send completes.  
        // If all of the data has been sent, then this method calls StartReceive
        //to start another receive op on the socket to read any additional 
        // data sent from the client. If all of the data has NOT been sent, then it 
        //calls StartSend to send more data.        
        private void ProcessSend(SocketAsyncEventArgs receiveSendEventArgs)
        {
            DataHolderToken token = (DataHolderToken)receiveSendEventArgs.UserToken;

            if (receiveSendEventArgs.SocketError == SocketError.Success)
            {
                token.sendBytesRemainingCount = token.sendBytesRemainingCount - receiveSendEventArgs.BytesTransferred;

                if (token.sendBytesRemainingCount == 0)
                {
                    // If we are within this if-statement, then all the bytes in
                    // the message have been sent. 
                    token.Reset();
                    CloseClientSocket(receiveSendEventArgs);
                }
                else
                {
                    // If some of the bytes in the message have NOT been sent,
                    // then we will need to post another send operation, after we store
                    // a count of how many bytes that we sent in this send op.                    
                    token.bytesSentAlreadyCount += receiveSendEventArgs.BytesTransferred;
                    // So let's loop back to StartSend().
                    StartSend(receiveSendEventArgs);
                }
            }
            else
            {
                // We'll just close the socket if there was a
                // socket error when receiving data from the client.
                token.Reset();
                CloseClientSocket(receiveSendEventArgs);
            }
        }

        //Post a send.    
        private void StartSend(SocketAsyncEventArgs receiveSendEventArgs)
        {
            DataHolderToken token = (DataHolderToken)receiveSendEventArgs.UserToken;
            int count = 0;
            try
            {
                //The number of bytes to send depends on whether the message is larger than
                //the buffer or not. If it is larger than the buffer, then we will have
                //to post more than one send operation. If it is less than or equal to the
                //size of the send buffer, then we can accomplish it in one send op.                
                if (token.sendBytesRemainingCount <= (receiveBufferSize + settings.send_buffer_size))
                {
                    count = token.sendBytesRemainingCount;
                    receiveSendEventArgs.SetBuffer(token.sendBufferOffset, count);
                    //Copy the bytes to the buffer associated with this SAEA object.
                    Buffer.BlockCopy(token.dataToSend, token.bytesSentAlreadyCount, receiveSendEventArgs.Buffer, token.sendBufferOffset, token.sendBytesRemainingCount);
                }
                else
                {
                    logger.Write("Exceeded buffer size in StartSend method, total size = " + token.sendBytesRemainingCount);
                    //We cannot try to set the buffer any larger than its size.
                    //So since receiveSendToken.sendBytesRemainingCount > BufferSize, we just
                    //set it to the maximum size, to send the most data possible.
                    count = receiveBufferSize + settings.send_buffer_size;
                    receiveSendEventArgs.SetBuffer(token.sendBufferOffset, count);
                    //Copy the bytes to the buffer associated with this SAEA object.
                    Buffer.BlockCopy(token.dataToSend, token.bytesSentAlreadyCount, receiveSendEventArgs.Buffer, token.sendBufferOffset, receiveBufferSize);                     
                }

                //post asynchronous send operation
                bool willRaiseEvent = receiveSendEventArgs.AcceptSocket.SendAsync(receiveSendEventArgs);

                if (!willRaiseEvent)
                {
                    ProcessSend(receiveSendEventArgs);
                }
            }
            catch (Exception e) 
            {
                logger.Write("Exception in SetBuffer in StartSend method of HttpServer.cs, offset=" + token.sendBufferOffset +
                    ", count=" + count + ", Buffer=" + receiveSendEventArgs.Buffer + ": " + e.Message);                  
                token.Reset();
                CloseClientSocket(receiveSendEventArgs);
            }            
        }

        string[] getCommandsFromArgs(SocketAsyncEventArgs e)
        {
            DataHolderToken token = (DataHolderToken)e.UserToken;
            string[] req = token.httpRequest.Split(new char[] { '?' }, 2); //Strip off "?"
            string[] cmd_stack = req[0].Split(new char[] { '/' });
            string[] command = cmd_stack[0].Split(new char[] { ' ' }, 2);
            return command;
        }

        private void TimerElapsedEvent(object source, ElapsedEventArgs e)
        {
            cacheTimer.Interval = 1000 * 60 * (60 * 24);
            cacheTimer.Enabled = true;
            startCacheBuildConditionally();
        }

        void startCacheBuildConditionally()
        {
            startCacheBuildConditionally(getSettingsDoc(), false);
        }

        void startCacheBuildNow()
        {
            startCacheBuildConditionally(getSettingsDoc(), true);
        }

        void startCacheBuildConditionally(XmlDocument doc, bool startNow)
        {
            if (doc != null)
            {
                if (!startNow)
                {
                    XmlNode lastCacheTimeNode = doc.DocumentElement.SelectSingleNode("lastCacheTime");
                    if (lastCacheTimeNode != null)
                    {
                        if (lastCacheTimeNode.InnerText != null && !lastCacheTimeNode.InnerText.Equals(""))
                        {
                            startNow = (DateTime.Now - Convert.ToDateTime(lastCacheTimeNode.InnerText)) > new TimeSpan(24, 0, 0);
                        }
                        else
                        {
                            startNow = true;
                        }
                    }
                    else
                    {
                        startNow = true;
                        XmlElement elem = doc.CreateElement("lastCacheTime");
                        XmlText text = doc.CreateTextNode("");
                        doc.DocumentElement.AppendChild(elem);
                        doc.DocumentElement.LastChild.AppendChild(text);
                    }
                    //Save settings
                    doc.PreserveWhitespace = true;
                    doc.Save(AddInModule.DATA_DIR + "\\settings.xml");

                    //Check for missing cache files
                    if (!startNow)
	                {
                        startNow = !remoteCommands.ExecuteCacheCheck();   
	                }                  
                }

                if (startNow)
                {
                    Thread http_thread = new Thread(new ParameterizedThreadStart(CacheBuildThread));
                    http_thread.IsBackground = true;
                    http_thread.SetApartmentState(ApartmentState.MTA);
                    http_thread.Start(doc);
                }
            }                                   
        }

        void CacheBuildThread(object o)
        {
            logger.Write("Cache update thread started");
            XmlDocument doc = (XmlDocument)o;
            remoteCommands.ExecuteCacheBuild(doc, logger);
        }
        
        void StartSendRequest(SocketAsyncEventArgs e)
        {            
            Thread http_thread = new Thread(new ParameterizedThreadStart(StartSendRequestThread));
            http_thread.IsBackground = false;
            http_thread.SetApartmentState(ApartmentState.STA);
            http_thread.Start(e);
        }                

        /// <summary>
        /// Handles the received commands of the m_httpServer control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="args">The <see cref="HttpServer.HttpEventArgs"/> instance containing the event data.</param>
        [STAThread]
        void StartSendRequestThread(Object o)
        {            
            System.Threading.Thread.CurrentThread.CurrentUICulture = System.Threading.Thread.CurrentThread.CurrentCulture;

            SocketAsyncEventArgs e = (SocketAsyncEventArgs)o;
            DataHolderToken token = (DataHolderToken)e.UserToken;
           
            string[] command = getCommandsFromArgs(e);
            string sCommand = command[0];
            string sParam = (command.Length == 2 ? command[1] : string.Empty);

            Thread http_thread = new Thread(new ParameterizedThreadStart(NewRequestThread));
            http_thread.IsBackground = true;
            http_thread.SetApartmentState(ApartmentState.MTA);

            if (sCommand.Equals("music-list-playing") || sCommand.Equals("music-list-current") || sCommand.StartsWith("play") || sCommand.Equals("music-shuffle"))
            {
                RemotedWindowsMediaPlayer remotedPlayer = null;
                try
                {
                    //Only allow one thread at a time to access remoted player
                    waitHandle.WaitOne();

                    remotedPlayer = new RemotedWindowsMediaPlayer();
                    remotedPlayer.CreateControl();

                    token.opResult = new OpResult();
                    if (sCommand.Equals("music-list-playing"))
                    {
                        if (sParam != null && sParam.Length != 0)
                        {
                            string sIndex = sParam.Substring(sParam.IndexOf("index:") + "index:".Length);
                            if (remotedPlayer.setNowPlaying(Int16.Parse(sIndex)))
                            {
                                token.opResult.StatusCode = OpStatusCode.Success;
                                token.opResult.StatusText = "Current media set to index " + sIndex;
                            }
                            else
                            {
                                token.opResult.StatusCode = OpStatusCode.BadRequest;
                                token.opResult.StatusText = "Current playback item not set";
                            }
                        }
                        else
                        {
                            token.opResult.StatusCode = OpStatusCode.Success;
                            NowPlaying nowPlaying = new NowPlaying(remotedPlayer);
                            token.opResult.ContentObject = nowPlaying;
                        }
                    }
                    else if (sCommand.StartsWith("play"))
                    {
                        //For playrate and playstate-get commands
                        token.opResult = remoteCommands.Execute(remotedPlayer, sCommand, sParam);
                    }
                    else if (sCommand.Equals("music-shuffle"))
                    {
                        remotedPlayer.setShuffleMode();
                        token.opResult.StatusCode = OpStatusCode.Success;
                        token.opResult.StatusText = "Shuffle mode set to true";
                    }
                    else
                    {
                        //"music-list-current" command
                        token.opResult.StatusCode = OpStatusCode.Success;
                        CurrentState state = new CurrentState(remotedPlayer);
                        token.opResult.ContentObject = state;
                    }
                }
                catch (Exception c)
                {
                    logger.Write("Exception in StartSendRequestThread: " + c.Message);
                    token.opResult.StatusCode = OpStatusCode.Exception;
                    token.opResult.StatusText = c.Message;
                }
                finally
                {
                    if (remotedPlayer != null) remotedPlayer.Dispose();
                    waitHandle.Set();
                }                
            }
            
            http_thread.Start(e);
        }      

        void NewRequestThread(Object o)
        {
            SocketAsyncEventArgs e = (SocketAsyncEventArgs)o;
            DataHolderToken token = (DataHolderToken)e.UserToken;

            OpResult opResult = new OpResult(OpStatusCode.BadRequest);
            string sCommand = "";
            string sParam = "";
            string sBody = "";
            try
            {
                // Show error for index
                if (token.httpRequest.Length == 0)
                {
                    sCommand = "<i>No command specified.</i>";
                    sParam = "<i>No parameters specified.</i>";
                }
                else
                {
                    string[] req = token.httpRequest.Split(new char[] { '?' }, 2); //Strip off "?"
                    string[] cmd_stack = req[0].Split(new char[] { '/' });
                    for (int idx = 0; idx < cmd_stack.Length; idx++)
                    {
                        string[] command = cmd_stack[idx].Split(new char[] { ' ' }, 2);
                        if (command.Length == 0)
                            return;
                        sCommand = command[0];
                        sParam = (command.Length == 2 ? command[1] : string.Empty);

                        if (sCommand.Equals("help", StringComparison.InvariantCultureIgnoreCase))
                        {
                            opResult = remoteCommands.CommandListHTML(AddInModule.GetPortNumber(AddInModule.m_basePortNumber));
                        }
                        else if (sCommand.Equals("format", StringComparison.InvariantCultureIgnoreCase))
                        {
                            ICommand formatter = new customCmd(sBody);
                            opResult = formatter.Execute(sParam);
                            sBody = "";
                        }
                        else if (token.opResult != null)
                        {
                            opResult = token.opResult;
                        }
                        else if (sCommand.Equals("server-settings"))
                        {
                            opResult = ExecuteSettingsCommand(sCommand, sParam);
                        }     
                        else if (sCommand.Equals("music-clear-cache"))
                        {
                            startCacheBuildNow();
                            opResult = new OpResult();
                            opResult.StatusCode = OpStatusCode.Success;
                            opResult.StatusText = "Cache cleared";
                        }      
                        else
                        {
                            opResult = remoteCommands.Execute(sCommand, sParam, settings);
                        }
                    }
                }

                //Get bytes to send to browser
                if (opResult.isHelpFormat())
                {
                    string sTempBody = opResult.ToString();
                    if (sParam.Length == 0)
                    {
                        sParam = "<i>No parameters specified.</i>";
                    }
                    if (opResult.StatusCode != OpStatusCode.Ok && opResult.StatusCode != OpStatusCode.Success)
                    {
                        sTempBody = string.Format("<h1>ERROR<hr>Command: {0}<br>Params: {1}<br>Returned: {2} - {3}<hr>See <a href='help'>Help</a></h1>", sCommand, sParam, opResult.StatusCode, opResult.ToString());
                    }
                    else if (opResult.StatusCode != OpStatusCode.OkImage)
                    {
                        if (sTempBody.Length > 0)
                        {
                            if (sTempBody.TrimStart()[0] != '<')
                            {
                                sTempBody = "<pre>" + sTempBody + "</pre>";
                            }
                        }
                        else
                        {
                            sTempBody = string.Format("<h1>Ok<hr>Last Command: '{0}'<br>Params: {1}<br>Returned: {2}<hr>See <a href='help'>Help</a></h1>", sCommand, sParam, opResult.StatusCode);
                        }
                    }
                    if (sBody.Length > 0)
                    {
                        sBody += "<HR>";
                    }
                    sBody += sTempBody;
                    token.dataToSend = GetPageDataToSend(string.Format("{0}\r\n", sBody));
                }
                else if (opResult.StatusCode != OpStatusCode.OkImage)
                {
                    token.dataToSend = GetPageJsonDataToSend(opResult.ToString());
                }
                else
                {
                    token.dataToSend = GetImageDataToSend(opResult);
                }               
            }
            catch (Exception ex)
            {
                logger.Write("Exception in NewRequestThread: " + ex.Message);
                token.dataToSend = GetPageDataToSend(string.Format("<html><body>EXCEPTION: {0}<hr></body></html>", ex.Message));
                Trace.TraceError(ex.ToString());
            }

            //Set send operation variables
            token.sendBytesRemainingCount = token.dataToSend.Length;
            token.bytesSentAlreadyCount = 0;

            StartSend(e);
        }  

        /// <summary>
        /// Sends the generated page
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="httpSocket">Socket reference</param>
        public Byte[] GetPageDataToSend(string text)
        {
            byte[] stringData = Encoding.UTF8.GetBytes(text);
            byte[] headerData = GetHeaderBytes(sHttpVersion, "text/html;charset=utf-8", stringData.Length, " 200 OK");
            byte[] totalData = new byte[stringData.Length + headerData.Length];
            Buffer.BlockCopy(headerData, 0, totalData, 0, headerData.Length);
            Buffer.BlockCopy(stringData, 0, totalData, headerData.Length, stringData.Length);
            return totalData;
        }

        public Byte[] GetPageJsonDataToSend(string text)
        {
            byte[] stringData = Encoding.UTF8.GetBytes(text);
            byte[] headerData = GetHeaderBytes(sHttpVersion, "application/json", stringData.Length, " 200 OK");
            byte[] totalData = new byte[stringData.Length + headerData.Length];
            Buffer.BlockCopy(headerData, 0, totalData, 0, headerData.Length);
            Buffer.BlockCopy(stringData, 0, totalData, headerData.Length, stringData.Length);
            return totalData;
        }

        /// <summary>
        /// Sends the served image
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="mime_ext">The mime extention.</param>
        /// <param name="httpSocket">Socket reference</param>
        public Byte[] GetImageDataToSend(OpResult opResult)
        {
            string mime_ext = opResult.StatusText;
            byte[] imageData = opResult.ImageData;
            if (mime_ext == "jpg") mime_ext = "jpeg";
            byte[] headerData = GetHeaderBytes(sHttpVersion, "image/" + mime_ext, imageData.Length, " 200 OK");
            byte[] totalData = new byte[imageData.Length + headerData.Length];
            Buffer.BlockCopy(headerData, 0, totalData, 0, headerData.Length);
            Buffer.BlockCopy(imageData, 0, totalData, headerData.Length, imageData.Length);
            return totalData;
        }

        /// <summary>
        /// This function sends the basic header Information to the client browser
        /// </summary>
        /// <param name="sHttpVersion">HTTP version</param>
        /// <param name="sMIMEHeader">Mime type</param>
        /// <param name="iTotBytes">Total bytes that will be sent</param>
        /// <param name="mySocket">Socket reference</param>
        /// <returns></returns>
        public Byte[] GetHeaderBytes(string sHttpVer, string sMimeType, int iBytesCount, string sStatusCode)
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
            return Encoding.UTF8.GetBytes(sBuffer);
        }


        // Does the normal destroying of sockets after 
        // we finish receiving and sending on a connection.        
        private void CloseClientSocket(SocketAsyncEventArgs e)
        {
            DataHolderToken token = (DataHolderToken)e.UserToken;

            // do a shutdown before you close the socket
            try
            {
                e.AcceptSocket.Shutdown(SocketShutdown.Both);
            }
            // throws if socket was already closed
            catch (Exception)
            {
            }

            //This method closes the socket and releases all resources, both
            //managed and unmanaged. It internally calls Dispose.
            e.AcceptSocket.Close();

            //Make sure the new DataHolder has been created for the next connection.
            //If it has, then dataMessageReceived should be null.
            if (token.dataMessageReceived != null)
            {
                token.Reset();
            }

            // Put the SocketAsyncEventArg back into the pool,
            // to be used by another client. This 
            poolOfRecSendEventArgs.Push(e);

            // decrement the counter keeping track of the total number of clients 
            //connected to the server, for testing
            Interlocked.Decrement(ref numConnectedSockets);

            if (numConnectedSockets + 1 == maxConnections) logger.Write("Number of connected sockets: " + numConnectedSockets + " / " + maxConnections);

            //Release Semaphore so that its connection counter will be decremented.
            //This must be done AFTER putting the SocketAsyncEventArg back into the pool,
            //or you can run into problems.
            maxNumberAcceptedClients.Release();
        }

        internal void CleanUpOnExit()
        {
            logger.Write("Shutting down server...");

            DisposeAllSaeaObjects();
        }

        private void DisposeAllSaeaObjects()
        {
            SocketAsyncEventArgs eventArgs;
            while (this.poolOfAcceptEventArgs.Count > 0)
            {
                eventArgs = poolOfAcceptEventArgs.Pop();
                eventArgs.Dispose();
            }
            while (this.poolOfRecSendEventArgs.Count > 0)
            {
                eventArgs = poolOfRecSendEventArgs.Pop();
                eventArgs.Dispose();
            }
        }

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                // Need to dispose managed resources if being called manually
                if (disposing)
                {
                    if (remoteCommands != null) remoteCommands.Dispose();
                    if (maxNumberAcceptedClients != null) maxNumberAcceptedClients.Close();
                    if (waitHandle != null) waitHandle.Close();
                    if (listenSocket != null) listenSocket.Close();
                    if (cacheTimer != null) cacheTimer.Dispose();
                    if (logger != null) logger.Dispose();
                    _disposed = true;
                }
            }
        }

        #endregion
	}
}
