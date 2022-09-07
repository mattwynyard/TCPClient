using System;
using System.Collections.Concurrent;
using System.Text;
using System.Net.Sockets;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Management;

namespace TCPClient
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("EAA4976A-45C3-4BC5-BC0B-E474F4C3C83F")]
    public interface IClient
    {
        string GetBuffer();
        bool IsConnected();
        bool IsJavaRunning();
        int SendCommand(String command);
        void CloseAll();
        int ConnectPhone(string jarPath, bool mode, string inspector, string path, bool hasPhone, bool hasMap);
        void KillProcessAndChildren(int pid);
        int StartNodeServer(string parentDirectory, bool debug);
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
        private TcpClient client;
        private NetworkStream stream;
        private Thread readThread;
        private Process clientProcess;
        private Process nodeProcess;
        private readonly object bufferLock = new Object();
        private bool mConnected;

        /// <summary>
        /// The constructor for the Client which starts the Read method running on a new thread
        /// </summary>       
        public Client()
        {
            readThread = new Thread(Read);
            readThread.Start();
            mConnected = false;
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
                        
                    }
                }
                catch (SocketException e) //not connected to server
                {
                    dataQueue = new ConcurrentQueue<string>();
                    dataQueue.Enqueue("server waiting...");
                    Thread.Sleep(100);
                }
            }
            while (stream.CanRead)
            {
                mConnected = true;
                string received = null;
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
                    catch (Exception e)
                    {
                        dataQueue = new ConcurrentQueue<string>();
                        dataQueue.Enqueue(e.Message);
                        mConnected = false;
                }
            }
            KillProcessAndChildren(clientProcess.Id);
        }

        /// <summary>
        /// Called by VBA to start java process which intialises a connection to android phone.
        /// </summary>
        /// <param name="jarPath">The path of the java executable.</param>
        /// <returns>
        /// Int - the process id or -1 if process didnt start or has stopped.
        /// </returns>
        [DispId(2)]
        public int ConnectPhone(string jarPath, bool mode, string inspector, string path, bool hasPhone, bool hasMap)
        {
            clientProcess = new Process();
            if (mode)
            {
                clientProcess.StartInfo.FileName = "java";
            } else
            {
                clientProcess.StartInfo.FileName = "javaw";
            }
            clientProcess.StartInfo.Arguments = @"-jar " + jarPath + " " + hasPhone + " " + hasMap + " " + inspector + " " + path;
            try
            {
                clientProcess.Start();
                if (clientProcess.HasExited || clientProcess == null)
                {
                    return 0;
                }
                else
                {
                    return clientProcess.Id;
                }
            } catch (Exception err)
            {
                return 0;
            }
        }

        /// <summary>
        /// Called by VBA read messages in the queue. Returns messages and empties buffer.
        /// </summary>
        /// <returns>
        /// String - the messages in the queue.
        /// </returns>
        [DispId(4)]
        public string GetBuffer()
        {
            string line = "";
            lock (bufferLock)
            {
                if (dataQueue == null)
                {
                    return "";
                }
                string[] data = dataQueue.ToArray();
               
                dataQueue = null;
                line = string.Join("", data);
            }
            return line;
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
        public int SendCommand(string command)
        {
            if (stream.CanWrite)
            {
                try
                {
                    byte[] buffer = Encoding.ASCII.GetBytes(command);
                    stream.Write(buffer, 0, buffer.Length);
                    return 0; 
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

        [DispId(6)]
        public bool IsJavaRunning()
        {
            try
            {
                if (clientProcess != null && !clientProcess.HasExited)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            } catch
            {
                return false;
            }  
        }

        /// <summary>
        /// Called by VBA to close connection to server and kill process
        /// </summary>
        [DispId(7)]
        public void CloseAll()
        {
            try
            {
                stream.Close();
                client = null;
                readThread = null;
                mConnected = false;
            } catch (Exception err)
            {
                //dataQueue.Enqueue(err.Message);
            }
        }

        [DispId(8)]
        public bool IsConnected()
        {
            return mConnected;
        }

        [DispId(9)]
        public void KillProcessAndChildren(int pid)
        {
            CloseAll();
            ManagementObjectSearcher processSearcher = new ManagementObjectSearcher
              ("Select * From Win32_Process Where ParentProcessID=" + pid);
            ManagementObjectCollection processCollection = processSearcher.Get();
            try
            {
                Process proc = Process.GetProcessById(pid);
                if (!proc.HasExited) proc.Kill();
            }
            catch (ArgumentException)
            {
                // Process already exited.
            }

            if (processCollection != null)
            {
                foreach (ManagementObject mo in processCollection)
                {
                    KillProcessAndChildren(Convert.ToInt32(mo["ProcessID"]));
                }
            }
        }

        [DispId(10)]
        public int StartNodeServer(string parentDirectory, bool debug)
        {
            nodeProcess = new Process();
            nodeProcess.StartInfo.WorkingDirectory = parentDirectory;
            nodeProcess.StartInfo.FileName = "cmd.exe";
            nodeProcess.StartInfo.Arguments = "/c node src/app.js";
            if (debug)
            {
                nodeProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
            } 
            else 
            {
                nodeProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            }
            nodeProcess.Start();
            return clientProcess.Id;
        }
    }
}





