using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Net.Sockets;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;

namespace TCPClient
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("01A31113-9353-44cc-A1F4-C6F1210E4B30")]

    public interface _Client
    {
        string getBuffer();
        Boolean isConnected();
        int sendCommand(String command);
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
    public class Client : _Client
    {
        private ConcurrentQueue<string> dataQueue = new ConcurrentQueue<string>();
        private Boolean mLock = false;
        private Boolean mConnected = false;
        private TcpClient client;
        private NetworkStream stream;

     /// <summary>
    /// The constructor for the Client which starts the Read method running on a new thread
    /// </summary>
        public Client()
        {

            Thread readThread = new Thread(Read);
            readThread.Start();
            //startTimer();
        }

        public void Read()
        {
            while (client == null)
            {
                try
                {
                    client = new TcpClient("localhost", 38200);
                    //console.writeline("waiting for bluetooth...");
                    if (client.Connected == true)
                    {
                        stream = client.GetStream();
                        mConnected = true;

                    }
                }
                catch (SocketException e)
                {
                    Console.WriteLine("server not ready...");
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
                            //Console.WriteLine(received);
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
                        Console.WriteLine(e.Message);
                    }
                }
            }
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
                    return "null";
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
                byte[] buffer = Encoding.ASCII.GetBytes(command);
                stream.Write(buffer, 0, buffer.Length);
                return 0; //success
            }
            else
            {
                //failed
                return -1;
            }
        }

        public Boolean isConnected()
        {
            return mConnected;
        }
    }
}