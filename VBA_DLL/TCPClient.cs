using System;
using System.Collections.Concurrent;
using System.Text;
using System.Net.Sockets;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace TCPClient
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("01A31113-9353-44CC-A1F4-C6F1210E4B30")]

    public interface IClient
    {
        string getBuffer();
        Boolean isConnected();
        int sendCommand(String command);
        void closeAll();
        int connect(String path);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("E2F07CD4-CE73-4102-B35D-119362624C47")]
    [ProgId("TCPClient.dll")]

    /// <summary>
    /// The class creates a TCP client which reads from local host port 38200 on a seperate thread acting as a client.
    /// Data received from the server is put into a concurrent queue via a locking mechanism for thread safety.
    /// The VBA program can call the getBuffer() method which puts the data in the queue into a string and returns it to VBA,
    /// then set the queue to null, thus acting like a buffer.
    /// </summary>
    public class Client : IClient
    {
        private ConcurrentQueue<string> dataQueue = new ConcurrentQueue<string>();
        private Boolean mLock = false;
        private Boolean mConnected = false;
        private TcpClient client;
        private NetworkStream stream;
        private Thread readThread;
        private Process clientProcess;

        /// <summary>
        /// The constructor for the Client which starts the Read method running on a new thread
        /// </summary>
        public Client()
        {
            readThread = new Thread(Read);
            readThread.Start();
        }
        /// <summary>
        /// The constructor for the Client which starts the Read method running on a new thread
        /// </summary>
        public void Read()
        {
            while (client == null)
            {
                try
                {
                    client = new TcpClient("localhost", 38200);
                    if (client.Connected == true)
                    {
                        stream = client.GetStream();
                        mConnected = true;
                    }
                }
                catch (SocketException e) //not connected to server
                {
                    //Console.WriteLine("server not ready...");
                    dataQueue = new ConcurrentQueue<string>();
                    dataQueue.Enqueue(e.Message + '\n');
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
                            while (!mLock)
                            {
                                mLock = true;
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
                            mLock = false;
                        }
                    }
                    catch (IOException e)
                    {
                        //Console.WriteLine(e.Message);
                        dataQueue.Enqueue(e.Message + '\n');
                    }
                }
            }
        }

        public int connect(String jarPath)
        {
            clientProcess = new Process();
            clientProcess.StartInfo.FileName = "java";
            clientProcess.StartInfo.Arguments = @"-jar " + jarPath;
            clientProcess.Start();
            if (clientProcess.HasExited || clientProcess == null)
            {
                return -1;
            }
            return clientProcess.Id;
        }

        //Called by VBA
        public string getBuffer()
        {
            string s = "";
            while (!mLock)
            {
                mLock = true;
                if (dataQueue == null)
                {
                    return "";
                }
                String[] data = dataQueue.ToArray();
                dataQueue = null;
                s = string.Join(",", data);
                data = null;
            }
            mLock = false;
            return s;
        }

        //Called by VBA
        public int sendCommand(String command)
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

        public Boolean isConnected()
        {
            return mConnected;
        }

        public void closeAll()
        {
            if (client != null)
            {
                mConnected = false;
                stream.Close();
                client.Close();
            }
            readThread = null;
            clientProcess.Kill();
        }
    }
}

//        static void Main(string[] args)
//        {
//            Client c = new Client();
//            c.connect("C:\\IntelliJ_CameraApp.jar");
//            int count = 0;
//            while (true)
//            {
//                if (count == 200)
//                {
//                    Console.WriteLine("closing...");
//                    try
//                    {
//                        c.closeAll();
//                    }
//                    catch (Exception e)
//                    {
//                        Console.WriteLine(e.Message);
//                    }
//                    break;
//                }
//                else
//                {
//                    String s = c.getBuffer();
//                    Console.Write(s);
//                    count++;
//                    Thread.Sleep(1000);
//                }

//            }
//        }
//    }
//}
