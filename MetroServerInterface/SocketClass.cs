using System;
using System.Text;
using System.Net;
using System.Net.Sockets;
using System.Xml;
using System.IO;

namespace MetroServerInterface
{
    class SocketClass
    {
        private mainUI mainForm;
        private int port;
        private Socket listener;
        private Socket handler;

        private string data = null;
        private IPAddress ipAddress;

        public SocketClass(mainUI form, int port)
        {
            mainForm = form;
            this.port = port;
        }

        public void StartListening()
        {
            byte[] bytes = new Byte[10240];
            ipAddress = new IPAddress(new byte[] { 127, 0, 0, 1 });
            IPEndPoint localEndPoint = new IPEndPoint(ipAddress, port);
            LingerOption lingerOption = new LingerOption(false, 0);

            listener = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            listener.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.Linger, lingerOption);

            try
            {
                listener.Bind(localEndPoint);
                listener.Listen(10);

                while (true)
                {
                    handler = listener.Accept();
                    data = null;

                    while (true)
                    {
                        byte[] msg;
                        bytes = new byte[10240];
                        int bytesRec = handler.Receive(bytes);
                        data += Encoding.ASCII.GetString(bytes, 0, bytesRec);

                        if (mainForm.accountVerify(data))
                        {
                            string incomming = data.Substring(0, data.IndexOf('@'));
                            Console.WriteLine("IN : "+incomming);
                            if (incomming.Equals("read") || incomming.Equals("clear") || incomming.Equals("init") || incomming.Equals("restart"))
                            {
                                msg = Encoding.ASCII.GetBytes("PERM_DENIED");
                                mainForm.logError(msg.ToString());
                            }
                            else if (incomming.StartsWith("setTask"))
                            {
                                try
                                {
                                    string xml = incomming.Substring(incomming.IndexOf(":") + 1);
                                    bool state = false;
                                    if (mainForm.InvokeRequired)
                                    {
                                        mainForm.Invoke(new Action(() => state = mainForm.createNewTaskFromXML(xml)));
                                    }
                                    if (state)
                                    {
                                        msg = Encoding.ASCII.GetBytes("TASK_ADDED");
                                        mainForm.logError(msg.ToString());
                                    } else
                                    {
                                        msg = Encoding.ASCII.GetBytes("TASK_ADD_FAILED");
                                        mainForm.logError(msg.ToString());
                                    }
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.StackTrace);
                                    mainForm.logError(e);
                                    msg = Encoding.ASCII.GetBytes("TASK_ADD_FAILED");
                                    mainForm.logError(msg.ToString());
                                }
                            }
                            else if (incomming.StartsWith("getTaskList", StringComparison.OrdinalIgnoreCase))
                            {
                                string xml = "";
                                if (mainForm.InvokeRequired)
                                {
                                    mainForm.Invoke(new Action(() => xml = mainForm.getTaskList()));
                                }
                                if (!xml.Equals(""))
                                {
                                    msg = Encoding.ASCII.GetBytes(xml);
                                }
                                else
                                {
                                    msg = Encoding.ASCII.GetBytes("GET_TASK_LIST_FAILED");
                                    mainForm.logError(msg.ToString());
                                }
                            }
                            else if (incomming.StartsWith("removeTask", StringComparison.OrdinalIgnoreCase))
                            {
                                string name = incomming.Substring(incomming.IndexOf(":")+1);
                                if (mainForm.deleteTask(name))
                                {
                                    msg = Encoding.ASCII.GetBytes("TASK_DELETED");
                                }
                                else
                                {
                                    msg = Encoding.ASCII.GetBytes("FAILED_TO_DELETE_TASK");
                                    mainForm.logError(msg.ToString());
                                }
                            }
                            else
                            {
                                mainForm.setLabelText(incomming);
                                msg = Encoding.ASCII.GetBytes(mainForm.lblOut.Text);
                            }
                        }
                        else
                        {
                            msg = Encoding.ASCII.GetBytes("Incorrect");
                            mainForm.logError(msg.ToString());
                        }
                        
                        handler.Send(msg);
                        handler.Shutdown(SocketShutdown.Both);
                        handler.Close();
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
                Console.WriteLine(e.Data);
                mainForm.logError(e);
            }
            Console.WriteLine("SOCKET_THREAD_CLOSED");
            mainForm.logError("SOCKET_THREAD_CLOSED");
        }

        public void stopConnection()
        {
            try {
                handler.Dispose();
                listener.Dispose();
            } catch (Exception e) { mainForm.logError(e); }
        }
    }
}