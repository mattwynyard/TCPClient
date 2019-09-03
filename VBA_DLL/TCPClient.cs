using System;
using System.Collections.Concurrent;
using System.Text;
using System.Net.Sockets;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Drawing;
using stdole;
using System.Windows.Forms;

namespace TCPClient
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("01A31113-9353-44CC-A1F4-C6F1210E4B30")]

    public interface IClient
    {
        String GetBuffer();
        object getPicture();
        Int32 GetBufferLength();
        object GetPhotoBuffer();
        Boolean IsConnected();
        int SendCommand(String command);
        Boolean CloseAll();
        int Connect(String path, Boolean mode, int camera);
    }
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("E2F07CD4-CE73-4102-B35D-119362624C47")]
    [ProgId("TCPClient.Client")]

    /// <summary>
    /// The class creates a TCP client which reads from local host port 38200 on a seperate thread acting as a client.
    /// Data received from the server is put into a concurrent queue via a locking mechanism for thread safety.
    /// The VBA program can call the getBuffer() method which puts the data in the queue into a string and returns it to VBA,
    /// then set the queue to null, thus acting like a buffer.
    /// </summary>
    public class Client : IClient
    {
        private ConcurrentQueue<string> dataQueue = new ConcurrentQueue<string>();
        //private ConcurrentQueue<StdPicture> photoQueue = new ConcurrentQueue<StdPicture>();
        private ConcurrentQueue<Image> photoQueue = new ConcurrentQueue<Image>();
        //private Boolean mLock = false;
        private Boolean mConnected = false;
        private TcpClient client;
        private TcpClient photoClient;
        private NetworkStream stream;
        private NetworkStream photoStream;
        private Thread readThread;
        private Thread photoReadThread;
        private Process clientProcess;
        //private Process node;
        private readonly object bufferLock = new Object();
        private readonly object photoLock = new Object();

        private byte[] data;
        private Int32 dataLength;
        private Image image;
        private StdPicture pict;

        /// <summary>
        /// The constructor for the Client which starts the Read method running on a new thread
        /// </summary>
        public Client() 
        {
            readThread = new Thread(Read);
            readThread.Start();
            //photoReadThread = new Thread(ReadPhoto);
            //photoReadThread.Start();
        }

        internal sealed class IPictureDispHost : AxHost
        {
            /// <summary>
            /// Default Constructor, required by the framework.
            /// </summary>
            private IPictureDispHost() : base(string.Empty) { }
            /// <summary>
            /// Convert the image to an ipicturedisp.
            /// </summary>
            /// <param name="image">The image instance</param>
            /// <returns>The picture dispatch object.</returns>
            public new static object GetIPictureDispFromPicture(Image image)
            {
                return (IPictureDisp)GetIPictureDispFromPicture(image);
            }
            /// <summary>
            /// Convert the dispatch interface into an image object.
            /// </summary>
            /// <param name="picture">The picture interface</param>
            /// <returns>An image instance.</returns>
            public new static Image GetPictureFromIPicture(object picture)
            {
                return AxHost.GetPictureFromIPicture(picture);
            }
        }

        public void ReadPhoto()
        {
            while (photoClient == null)
            {
                try
                { 
                    photoClient = new TcpClient("localhost", 38300);
                    if (photoClient.Connected == true)
                    {
                        photoStream = photoClient.GetStream();
                        //mConnected = true;

                    }
                }
                catch (SocketException e) //not connected to server
                {
                    //Console.WriteLine("server not ready...");
                    dataQueue = new ConcurrentQueue<string>();
                    dataQueue.Enqueue("server not ready...");
                    Thread.Sleep(100);
                }
            }
            while (true)
            {
                if (photoStream.CanRead)
                {
                    try
                    {
                        if (photoStream.DataAvailable)
                        {
                            byte[] buffer = new byte[photoClient.ReceiveBufferSize];

                            dataLength = photoStream.Read(buffer, 0, buffer.Length);
                            //sendToQueue(dataLength + " bytes read");
                            data = new byte[dataLength];
                            Buffer.BlockCopy(buffer, 0, data, 0, dataLength);

                            MemoryStream ms = new MemoryStream(data, 0, data.Length);
                           
                            ms.Position = 0;
                            ms.Write(data, 0, data.Length);
                            image = Image.FromStream(ms, true);// this line giving exception parameter not valid
                                                                     //Image bitmap = new Im(ms);



                            pict = (StdPicture)IPictureDispHost.GetIPictureDispFromPicture(image);


                            //bitmap.Save(@"C:\Road Inspection\Thumbnails\MyImage2.bmp");
                            if (photoQueue == null)
                            {
                                photoQueue = new ConcurrentQueue<Image>();

                                photoQueue.Enqueue(image);
                            }
                            else
                            {
                                photoQueue.Enqueue(image);
                            }



                        }
                    }
                    catch (IOException e)
                    {
                        //Console.WriteLine(e.Message);
                        dataQueue.Enqueue(e.Message);
                    }
                }
                Thread.Sleep(100);
            }
        }

        private void sendToQueue(String message)
        {
            if (dataQueue == null)
            {
                dataQueue = new ConcurrentQueue<string>();

                dataQueue.Enqueue(message);
            }
            else
            {
                dataQueue.Enqueue(message);
            }
        }
        /// <summary>
        /// Creates a new TCP Client and intialises input stream. When the server connects client reads incoming
        /// data from the buffer. New data is added to a thread safe concurrent queue. This method runs on its own thread
        /// and reads data in a blocking while loop.
        /// </summary>
        public void Read()
        {
            while (client == null)
            {
                try
                {
                    client = new TcpClient("localhost", 38200);
                    stream = client.GetStream();
                    //if (stream == null)
                    //{
                    //    dataQueue = new ConcurrentQueue<string>();
                    //    dataQueue.Enqueue("stream null...");
                    //}
                    mConnected = true;
                    //}

                }
                catch (SocketException e) //not connected to server
                {
                    dataQueue = new ConcurrentQueue<string>();
                    dataQueue.Enqueue("server not ready...");
                    Thread.Sleep(100);
                }
            }
            while (true)
            {
                string received = null;
                if (stream.CanRead)
                {
                    try
                    {
                        if (stream.DataAvailable)
                        {
                            byte[] buffer = new byte[client.ReceiveBufferSize];
                            Int32 receiveCount = stream.Read(buffer, 0, buffer.Length);
                            received = new ASCIIEncoding().GetString(buffer, 0, receiveCount);
                            if (dataQueue == null)
                            {
                                dataQueue = new ConcurrentQueue<string>();
                               
                                dataQueue.Enqueue(received);
                            }
                            else
                            {
                                dataQueue.Enqueue(received);
                            }
                        }
                    }
                    catch (IOException e)
                    {

                        if (dataQueue == null)
                        {
                            dataQueue = new ConcurrentQueue<string>();

                            dataQueue.Enqueue(e.Message);
                        }
                        else
                        {
                            dataQueue.Enqueue(e.Message);
                        }
                    }
                }
                Thread.Sleep(100);
            }
        }
        /// <summary>
        /// Called by VBA to start java process which intialises a connection to android phone.
        /// </summary>
        /// <param name="jarPath">The path of the java executable.</param>
        /// <returns>
        /// Int - the process id or -1 if process didnt start or has stopped.
        /// </returns>
        public int Connect(String jarPath, Boolean mode, int cameras)
        {
            clientProcess = new Process();
            if (mode)
            {
                clientProcess.StartInfo.FileName = "java";
                
            } else
            {
                clientProcess.StartInfo.FileName = "javaw";
            }
            clientProcess.StartInfo.Arguments = @"-jar " + jarPath + " " + cameras;
            clientProcess.Start();


            //node = new Process();
            //node.StartInfo.FileName = "node";
            //node.StartInfo.Arguments = "C:\\html\\server.js";
            //node.Start();


            if (clientProcess.HasExited || clientProcess == null)
            {
                return -1;
            }
            else
            {
                return clientProcess.Id;
            }

        }

        public object getPicture()
        {

                return pict;

        
        }

        /// <summary>
        /// Called by VBA read messages in the queue. Returns messages and empties buffer.
        /// </summary>
        /// <returns>
        /// String - the messages in the queue.
        /// </returns>
        public String GetBuffer()
        {
            string s = "";
            //while (!mLock)
            //lock(bufferLock)
            //{
                //mLock = true;
                if (dataQueue == null)
                {
                    return "";
                }
                String[] data = dataQueue.ToArray();
                dataQueue = null;
                s = string.Join("", data);
                data = null;
            //}
            //mLock = false;
            return s;
        }


        /// <summary>
        /// Called by VBA read messages in the queue. Returns messages and empties buffer.
        /// </summary>
        /// <returns>
        /// String - the messages in the queue.
        /// </returns>
        public object GetPhotoBuffer()
        {
            object[] data;
            //lock (photoLock)
            //{
            data = photoQueue.ToArray();
            //}
            //sendToQueue("Ärray: " + data.Length.ToString() + ",");
                if (data != null)
                {
                    photoQueue = null;
                    return data[0];
                }
                else
                {
                    photoQueue = null;
                    return null;
                }
            //}
                
            
        }

        public Int32 GetBufferLength()
        {
            //lock (photoLock)
            if (data != null)
            {
                return dataLength;
            } else
            {
                return -1;
            }
           
        }

        /// <summary>
        /// Called by VBA to send command to the sever.
        /// </summary>
        /// <param name="command">The string sent to the server .</param>
        /// <returns>
        /// Integer - Either 0 for success or -1 for failure
        /// </returns>
        /// <exception cref="System.IO.IOException">Thrown when the socket is closed.
        /// and the other is greater than zero.</exception>
        public int SendCommand(String command)
        {
            if (stream.CanWrite)
            {
                try
                {
                    byte[] buffer = Encoding.ASCII.GetBytes(command);
                    stream.Write(buffer, 0, buffer.Length);
                    return 0; //success
                }
                catch (IOException e)
                {
                    return -1;
                }
            }
            else
            {
                return -1;//failed
            }
        }
        /// <summary>
        /// Called by VBA to check if connected to server.
        /// </summary>
        /// <returns>
        /// Boolean - true if connected, false if not connected.
        /// </returns>
        public Boolean IsConnected()
        {
            return mConnected;
        }

        /// <summary>
        /// Called by VBA to close connection to server and kill process
        /// </summary>
        public Boolean CloseAll()
        {
            if (client != null)
            {
                mConnected = false;
                stream.Dispose();
                photoStream.Dispose();
                client = null;
                photoClient = null;
            }
            readThread.Abort();
            readThread = null;
            photoReadThread.Abort();
            photoReadThread = null;
            clientProcess.Kill();
            clientProcess.WaitForExit(1000);
            return clientProcess.HasExited;

        }
    }
}


