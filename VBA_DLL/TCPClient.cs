using System;
using System.Collections.Concurrent;
using System.Text;
using System.Net.Sockets;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
//using stdole;

namespace TCPClient
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("EAA4976A-45C3-4BC5-BC0B-E474F4C3C83F")]
    public interface IClient
    {
        string GetBuffer();

        bool IsConnected();
        int SendCommand(String command);
        Boolean CloseAll();
        int Connect(string jarPath, bool mode, string camera, string path);
        
    }


    /// <summary>
    /// The class creates a TCP client which reads from local host port 38200 on a seperate thread acting as a client.
    /// Data received from the server is put into a concurrent queue via a locking mechanism for thread safety.
    /// The VBA program can call the getBuffer() method which puts the data in the queue into a string and returns it to VBA,
    /// then set the queue to null, thus acting like a buffer.
    /// </summary>
    [Guid("0D53A3E8-E51A-49C7-944E-E72A2064F938"),
        ClassInterface(ClassInterfaceType.None)]
    [ProgId("TCPClient.mClient")]
    public class Client : IClient
    {
        private ConcurrentQueue<string> dataQueue = new ConcurrentQueue<string>();
        private Boolean mConnected = false;
        private TcpClient client;
        private NetworkStream stream;
        private Thread readThread;
        private Process clientProcess;
        private readonly object bufferLock = new Object();

        /// <summary>
        /// The constructor for the Client which starts the Read method running on a new thread
        /// </summary>       
        public Client()
        {
            readThread = new Thread(Read);
            readThread.Start();
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
                    if (client.Connected)
                    {
                        stream = client.GetStream();
                        mConnected = true;                   
                    }
                }
                catch (SocketException e) //not connected to server
                {
                    dataQueue = new ConcurrentQueue<string>();
                    dataQueue.Enqueue("server waiting...");
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
                        dataQueue = new ConcurrentQueue<string>();
                        dataQueue.Enqueue(e.Message);
                    }
                }
                Thread.Sleep(10);
            }
        }


        /// <summary>
        /// Called by VBA to start java process which intialises a connection to android phone.
        /// </summary>
        /// <param name="jarPath">The path of the java executable.</param>
        /// <returns>
        /// Int - the process id or -1 if process didnt start or has stopped.
        /// </returns>
        [DispId(2)]
        public int Connect(string jarPath, bool mode, string camera, string path)
        {
            clientProcess = new Process();
            if (mode)
            {
                clientProcess.StartInfo.FileName = "java";
            } else
            {
                clientProcess.StartInfo.FileName = "javaw";
            }
            clientProcess.StartInfo.Arguments = @"-jar " + jarPath + " " + camera + " " + path;
            clientProcess.Start();
            if (clientProcess.HasExited || clientProcess == null)
            {
                return -1;
            } else
            {
                return clientProcess.Id;
            }
            
        }

        /// <summary>
        /// Called by VBA read messages in the queue. Returns messages and empties buffer.
        /// </summary>
        /// <returns>
        /// String - the messages in the queue.
        /// </returns>
        [DispId(4)]
        public String GetBuffer()
        {
            string s = "";
            lock(bufferLock)
            {
                if (dataQueue == null)
                {
                    return "";
                }
                String[] data = dataQueue.ToArray();
                dataQueue = null;
                s = string.Join("", data);
            }
            return s;
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
        [DispId(5)]
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
        [DispId(6)]
        public Boolean IsConnected()
        {
            return mConnected;
        }

        /// <summary>
        /// Called by VBA to close connection to server and kill process
        /// </summary>
        [DispId(7)]
        public Boolean CloseAll()
        {
            if (client != null)
            {
                mConnected = false;
                stream.Dispose();
                client = null;
            }
            readThread.Abort();
            readThread = null;
            clientProcess.Kill();
            clientProcess.WaitForExit(1000);
            return clientProcess.HasExited;

        }
    }
}





