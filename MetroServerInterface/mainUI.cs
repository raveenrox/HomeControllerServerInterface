using MetroFramework.Forms;
using System;
using System.Windows.Forms;
using MetroFramework.Controls;
using System.Threading;
using System.IO.Ports;
using Microsoft.Win32.TaskScheduler;
using System.Linq;
using System.Data;
using System.Xml;
using System.IO;
using System.Xml.XPath;
using System.Collections.Generic;
using System.Collections;

namespace MetroServerInterface
{
    public partial class mainUI : MetroForm
    {
        private SocketClass socketClass;
        private Thread socketThread;

        private String username;
        private String password;

        private string webPath="";

        private TaskService taskService;
        private System.Globalization.CultureInfo cultureInfo;

        bool[] customArr = new bool[] {false, false, false, false, false};

        public mainUI()
        {
            InitializeComponent();
            try
            {
                taskService = new TaskService();
                taskService.RootFolder.CreateFolder("HAS");
            }
            catch (Exception e) { Console.WriteLine(e.StackTrace); logError(e); }

            gridTS.DataSource = dataSetTaskList.Tables["child"].DefaultView;
            loadXML();
            loadPrefs();
            loadUserDetails();
            cultureInfo = System.Globalization.CultureInfo.InvariantCulture;
        }

        public string getTaskList()
        {
            try
            {
                string xml = "<tasklist>";
                
                TaskFolder taskFolder = taskService.GetFolder("HAS");

                var tasks = taskFolder.Tasks.Where(t => t.Name.StartsWith("", StringComparison.OrdinalIgnoreCase));
                foreach (Microsoft.Win32.TaskScheduler.Task task in tasks)
                {
                    xml += "<task>";
                    xml += "<name>";
                    xml += task.Name;
                    xml += "</name>";
                    xml += "<date>";
                    xml += task.NextRunTime.ToString("yyyy-MM-dd", cultureInfo);
                    xml += "</date>";
                    xml += "<time>";
                    xml += task.NextRunTime.TimeOfDay;
                    xml += "</time>";

                    xml += "<commands>";
                    
                    string line = task.Definition.Actions.ToString();
                    line = line.Substring(line.LastIndexOf(' ') + 1);
                    string[] lineList = line.Split('X');
                    foreach (string str in lineList)
                    {
                        if (!str.Equals(""))
                        {
                            xml += "<command>";

                            string command = "";
                            string state = "";
                            command = str.Substring(0, str.IndexOf('Y'));
                            state = str.Substring(str.IndexOf('Y') + 1, 1);

                            xml += "<no>" + command + "</no>";
                            xml += "<state>";
                            if (state.Equals("1"))
                            {
                                xml += "true";
                            }
                            else if (state.Equals("0"))
                            {
                                xml += "false";
                            }
                            xml += "</state>";
                            xml += "</command>";
                        }
                    }
                    xml += "</commands>";
                    xml += "</task>";
                }


                xml += "</tasklist>";
                Console.Write(xml);
                return xml;
            }
            catch (Exception e) { logError(e); }
            return "";
        }

        public bool createNewTaskFromXML(string xml)
        {
            try
            {
                XmlDocument xmlTask = new XmlDocument();
                xmlTask.LoadXml(xml);

                XmlNode name_node = xmlTask.GetElementsByTagName("name").Item(0);
                XmlNode time_node = xmlTask.GetElementsByTagName("time").Item(0);
                XmlNode date_node = xmlTask.GetElementsByTagName("date").Item(0);
                XmlNode msg_node = xmlTask.GetElementsByTagName("msg").Item(0);
                Console.WriteLine(name_node.InnerText + "," + date_node.InnerText + "," + time_node.InnerText + "," + msg_node.InnerText);
                if (setTask(name_node.InnerText, date_node.InnerText, time_node.InnerText, msg_node.InnerText))
                    return true;
                else
                    return false;
            }
            catch (Exception e) { Console.WriteLine(e.StackTrace); logError(e); return false; }
        }

        public void saveXML(string line)
        {
            if (line.StartsWith("R"))
            {
                line = line.Substring(line.IndexOf('R')+1, (line.IndexOf('N') - line.IndexOf('R')-1));
                Console.WriteLine(line);
                XmlDocument xmlStatus = new XmlDocument();
                XmlElement root = (XmlElement) xmlStatus.AppendChild(xmlStatus.CreateElement("status"));
                XmlReader xmlReader = null;
                try
                {
                    xmlReader = XmlReader.Create("http://127.0.0.1/hc/db.xml");
                }
                catch (Exception e) { Console.WriteLine(e.StackTrace); logError(e); }
                List<String> commands = new List<String>();
                while (xmlReader.Read())
                {
                    if(xmlReader.NodeType == XmlNodeType.Element)
                    {
                        if (xmlReader.Name.Equals("command"))
                        {
                            commands.Add(xmlReader.ReadElementContentAsString());
                        }
                    }
                }
                int tempCount = 0;
                foreach (String str in commands)
                {
                    int o;
                    XmlElement child = (XmlElement)root.AppendChild(xmlStatus.CreateElement("child"));

                    if (Int32.TryParse(str, out o))
                    {
                        child.AppendChild(xmlStatus.CreateElement("command")).InnerText = o.ToString();
                        child.AppendChild(xmlStatus.CreateElement("state")).InnerText = line.ToCharArray().GetValue(o).ToString();
                    } else
                    {
                        child.AppendChild(xmlStatus.CreateElement("command")).InnerText = str;
                        /*String cval = "";
                        Console.WriteLine(str);
                        if (customArr[tempCount])
                        {
                            cval = "1";
                        } else
                        {
                            cval = "0";
                        }*/
                        child.AppendChild(xmlStatus.CreateElement("state")).InnerText = "1";
                        //TODO fix output
                    }
                    tempCount++;

                }
                File.WriteAllText(webPath+"/status.xml", xmlStatus.OuterXml);
            }
        }
        //test
        public void loadPrefs()
        {
            XmlDocument xmlSettings = new XmlDocument();
            try {
                xmlSettings.Load("preferences.xml");

                XmlNode _webPath = xmlSettings.GetElementsByTagName("webPath").Item(0);
                webPath = _webPath.InnerText;

            } catch (Exception e) { logError(e); }
        }

        public void loadUserDetails()
        {
            string xml="";
            try {
                xml = Base64Decode(File.ReadAllText(webPath+"/user.bin"));
            
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xml);
            
                XmlNodeList userList = xmlDoc.GetElementsByTagName("username");
                XmlNodeList passwordList = xmlDoc.GetElementsByTagName("password");
            
                txtUsername.Text = userList[0].InnerText;
                txtPassword.Text = Base64Decode(passwordList[0].InnerText);

            }
            catch (Exception e) { logError(e); }
        }

        public void saveUserDetails()
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement element = (XmlElement)xmlDoc.AppendChild(xmlDoc.CreateElement("details"));
            element.AppendChild(xmlDoc.CreateElement("username")).InnerText = txtUsername.Text;
            element.AppendChild(xmlDoc.CreateElement("password")).InnerText = Base64Encode(txtPassword.Text);
            File.WriteAllText(webPath+"/user.bin", Base64Encode(xmlDoc.OuterXml));
        }

        public void loadXML()
        {
            XmlDocument xmlDoc = new XmlDocument();
            try
            {
                xmlDoc.Load("http://127.0.0.1/hc/db.xml");

                XmlNodeList childName = xmlDoc.GetElementsByTagName("childname");
                XmlNodeList command = xmlDoc.GetElementsByTagName("command");

                int count = 0;
                dataSetTaskList.Tables[0].Rows.Clear();
                foreach (XmlNode xmlNode in childName)
                {
                    DataRow dataRow = dataSetTaskList.Tables[0].NewRow();
                    dataRow["childName"] = xmlNode.InnerText;
                    dataRow["childNo"] = command[count].InnerText;
                    dataSetTaskList.Tables[0].Rows.Add(dataRow);
                    count++;
                }
            }
            catch (Exception e) { Console.WriteLine(e.StackTrace); logError(e); }
        }

        public bool setTask(string name, string date, string time, string msg)
        {
            try
            {
                DateTime timeDT = DateTime.ParseExact(time, "HH:mm:ss", cultureInfo);
                DateTime dateDT = DateTime.ParseExact(date, "yyyy-MM-dd", cultureInfo);
                if (dateDT.CompareTo(DateTime.Today) == -1)
                {
                    lblStatusBar.Text = "Try a date greater than " + DateTime.Today.Date.ToString("yyyy-MM-dd");
                    return false;
                }
                else
                {
                    TaskDefinition taskDefinition = taskService.NewTask();
                    
                    taskDefinition.Triggers.Add(new TimeTrigger
                    {
                        StartBoundary = new DateTime(
                                    Convert.ToInt32(dateDT.Year.ToString()),
                                    Convert.ToInt32(dateDT.Month.ToString()),
                                    Convert.ToInt32(dateDT.Day.ToString()),
                                    Convert.ToInt32(timeDT.Hour.ToString()),
                                    Convert.ToInt32(timeDT.Minute.ToString()),
                                    Convert.ToInt32(timeDT.Second.ToString()))
                    });
                    taskDefinition.Actions.Add(new ExecAction("TSMsg", "/hc/android.php " + txtUsername.Text + " " + Base64Encode(txtPassword.Text)+ " "+msg, null));

                    try { taskService.RootFolder.CreateFolder("HAS"); } catch (Exception e) { logError(e); }

                    taskService.GetFolder("HAS").RegisterTaskDefinition(@name, taskDefinition);
                    loadTasks();

                    return true;
                }

            }
            catch (FormatException ex)
            {

                lblStatusBar.Text = "Invalid Date/Time";
                logError(ex);
                return false;
            }
        }

        public bool deleteTask(string name)
        {
            try
            {
                taskService.GetFolder("HAS").DeleteTask(name, false);
                loadTasks();
                return true;
            } catch(System.Runtime.InteropServices.COMException ex) { Console.Write("1"); return true; }
            catch(System.InvalidOperationException) { return true; }
            catch (Exception e) { logError(e); Console.Write("2"); return false; }
        }

        public void loadTasks()
        {
            comboBoxTSTaskList.Items.Clear();
            try
            {
                TaskFolder taskFolder = taskService.GetFolder("HAS");

                var tasks = taskFolder.Tasks.Where(t => t.Name.StartsWith("", StringComparison.OrdinalIgnoreCase));
                foreach (Microsoft.Win32.TaskScheduler.Task task in tasks)
                {
                    comboBoxTSTaskList.Items.Add(task.Name);
                }
                comboBoxTSTaskList.SelectedItem = comboBoxTSTaskList.Text;
            }
            catch (Exception e) { logError(e); }
        }

        public void setLabelText(string text)
        {
            if(text == "Garage" || text=="XGarageY1Z" || text == "XGarageY0Z")
            {
                writeSerial2("Garage");
            } else if (text == "Gate" || text == "XGateY1Z" || text == "XGateY0Z")
            {
                writeSerial2("Gate");
            } else if (text == "Door" || text == "XDoorY1Z" || text == "XDoorY0Z")
            {
                writeSerial2("Door");
            } else if (text == "ClothLine" || text == "XClothLineY1Z" || text == "XClothLineY0Z")
            {
                writeSerial2("ClothLine");
            }
            else if (text == "Fan" || text == "XFanY1Z" || text == "XFanY0Z")
            {
                writeSerial2("Fan");
            } else if(text == "Temp")
            {
                writeSerial2("Temp");
            } else if (text == "IRPowerOn" || text == "XIRPowerOnY0Z" || text == "XIRPowerOnY1Z") {
                writeSerial("powerOn");
            }
            else if (text == "IRPowerOff" || text == "XIRPowerOffY0Z" || text == "XIRPowerOffY1Z")
            {
                writeSerial("powerOff");
            }
            else if (text == "IROn" || text == "XIROnY0Z" || text == "XIROnY1Z")
            {
                writeSerial("pioneerOn");
            }
            else if (text == "IREject" || text == "XIREjectY0Z" || text == "XIREjectY1Z")
            {
                writeSerial("pioneerOpen");
            }
            else if (text == "GetIR" || text == "XGetIRY0Z" || text == "XGetIRY1Z")
            {
                writeSerial("getIR");
            }
            else if(text != "ping")
            {
                changeTextBox(txtMsg, text);
                writeSerial(txtMsg.Text);
            }
        }

        public void writeSerial(string str)
        {
            try
            {
                if (serialPortMain.IsOpen)
                {
                    serialPortMain.Write(str);
                }
                else
                {
                    changeLabel(lblStatusBar, "Pease open a port");
                }
            }
            catch (Exception ex)
            {
                changeLabel(lblStatusBar, ex.Message);
                logError(ex);
            }
        }

        public void changeTextBox(MetroTextBox txtbx, string str)
        {
            if (txtbx.InvokeRequired)
            {
                String name = str;
                txtbx.Invoke(new MethodInvoker(delegate { txtbx.Text = name; }));
            }
        }

        public void changeLabel(Label lbl, string str)
        {
            if (lbl.InvokeRequired)
            {
                String name = str;
                lbl.Invoke(new MethodInvoker(delegate { lbl.Text = name; }));
            }
        }

        public void consoleAddLine(MetroTextBox txtbx, string str)
        {
            if (txtbx.InvokeRequired)
            {
                String name = str;
                txtbx.Invoke(new MethodInvoker(delegate { txtbx.AppendText(name + "\n"); }));
            }
        }


        public void openPort()
        {
            try {
                if (comboBoxPorts.SelectedItem.ToString() == "Please Select a Port" || comboBoxPorts.SelectedItem.ToString().Equals(""))
                {
                    lblStatusBar.Text = "Please select a port";
                }
                else
                {
                    try
                    {
                        serialPortMain.PortName = comboBoxPorts.SelectedItem.ToString();
                        serialPortMain.Open();
                        comboBoxPorts.Enabled = false;
                        tileOpenPort.Enabled = false;
                        tileOpenPort.UseCustomBackColor = true;
                        tileColsePort.Enabled = true;
                        tileColsePort.UseCustomBackColor = false;
                        lblStatusBar.Text = serialPortMain.PortName + " Port Opened";
                    }
                    catch (Exception ex)
                    {
                        lblStatusBar.Text = serialPortMain.PortName + " Port is already open";
                        Console.Write(ex.Message);
                        logError(ex);
                    }
                }
            } catch (NullReferenceException e) { lblStatusBar.Text = "Please select a port"; logError(e); }
        }

        public void closePort()
        {
            try
            {
                serialPortMain.Close();
                if (comboBoxPorts.InvokeRequired)
                {
                    comboBoxPorts.Invoke(new MethodInvoker(delegate { comboBoxPorts.Enabled = true; }));
                }
                if(tileOpenPort.InvokeRequired)
                {
                    tileOpenPort.Invoke(new MethodInvoker(delegate
                        {
                            tileOpenPort.Enabled = true;
                            tileOpenPort.UseCustomBackColor = false;
                        }));
                }
                if (tileColsePort.InvokeRequired)
                {
                    tileColsePort.Invoke(new MethodInvoker(delegate
                    {
                        tileColsePort.Enabled = false;
                        tileColsePort.UseCustomBackColor = true;
                    }));
                }
                if (tileListenPort.InvokeRequired)
                {
                    tileListenPort.Invoke(new MethodInvoker(delegate
                    {
                        tileListenPort.Enabled = true;
                        tileListenPort.UseCustomBackColor = false;
                    }));
                }
                changeLabel(lblStatusBar, serialPortMain.PortName + " Port Closed");
            }
            catch (Exception ex)
            {
                lblStatusBar.Text = ex.Message;
                logError(ex);
            }
        }

        public void listenPort()
        {
            new Thread(() =>
            {
                String prevLine = "";
                String curLine = "";

                changeLabel(lblStatusBar, "Listening...");

                while (true)
                {
                    try
                    {
                        if (serialPortMain.IsOpen)
                        {
                            curLine = serialPortMain.ReadLine();
                            if (curLine.StartsWith("Fan")) {
                                writeSerial2("Fan");
                                Console.WriteLine("Fan");
                            } else if (curLine.StartsWith("Garage")) {
                                writeSerial2("Garage");
                                Console.WriteLine("Garage");
                            } else if (curLine.StartsWith("Gate")) {
                                writeSerial2("Gate");
                                Console.WriteLine("Gate");
                            } else if (curLine.StartsWith("ClothLine")) {
                                writeSerial2("ClothLine");
                                Console.WriteLine("ClothLine");
                            } else if(curLine.StartsWith("Door")) {
                                writeSerial2("Door");
                                Console.WriteLine("Door");
                            }
                            if (prevLine != curLine)
                            {
                                if (curLine[0] == 'R' && curLine[curLine.Length-2] == 'N')
                                {
                                    changeLabel(lblOut, curLine);
                                    saveXML(curLine);
                                }
                                consoleAddLine(txtConsoleMain, curLine);
                                
                            }
                            prevLine = curLine;
                        }
                        else
                        {
                            closePort();
                            Thread.CurrentThread.Abort();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.StackTrace);
                        changeLabel(lblStatusBar, ex.Message);
                        changeLabel(lblStatusBar, "Please open a port");
                        logError(ex);
                    }
                }
            }).Start();
        }

        public void startSocketServer()
        {
            socketClass = new SocketClass(this, Int32.Parse(txtSocket.Text));
            socketThread = new Thread(socketClass.StartListening);
            socketThread.Start();
            tileOpenSocket.Enabled = false;
            tileOpenSocket.UseCustomBackColor = true;
            tileColseSocket.Enabled = true;
            tileColseSocket.UseCustomBackColor = false;
            txtSocket.Enabled = false;
            
        }

        public Boolean accountVerify(String data)
        {
            String usr = "";
            String pwd = "";
            char[] arr;
            arr = data.ToCharArray();
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i] == '@')
                {
                    for (int j = i; j < arr.Length; j++)
                    {
                        if (arr[j] == ':')
                        {
                            for (int k = i + 1; k < j; k++)
                            {
                                usr += arr[k].ToString();
                            }
                            for (int k = j + 1; k < arr.Length; k++)
                            {
                                pwd += arr[k].ToString();
                            }
                            break;
                        }

                    }
                }
            }
            if (usr == this.username)
            {
                if (pwd == this.password)
                {
                    return true;
                }
            }
            return false;
        }

        public static string Base64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }
        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }

        public void logError(Exception e)
        {
            File.AppendAllText("ErrorLog.log", "\n" + DateTime.Now + "\t" + e.ToString() + e.StackTrace);
        }

        public void logError(string line)
        {
            File.AppendAllText("ErrorLog.log", "\n" + DateTime.Now + "\t" + line);
        }

        private void mainUI_Load(object sender, EventArgs e)
        {
            username = txtUsername.Text;
            password = txtPassword.Text;

            loadTasks();

            string[] ports = SerialPort.GetPortNames();
            
            foreach (string port in ports)
            {
                comboBoxPorts.Items.Add(port);
                comboBoxPorts2.Items.Add(port);
            }
            comboBoxPorts.Text = "Please select a port!";
        }

        private void tileOpenPort_Click(object sender, EventArgs e)
        {
            openPort();
        }

        private void tileColsePort_Click(object sender, EventArgs e)
        {
            closePort();
        }

        private void tileOpenSocket_Click(object sender, EventArgs e)
        {
            startSocketServer();
        }

        private void tileColseSocket_Click(object sender, EventArgs e)
        {
            socketClass.stopConnection();
            socketThread.Abort();

            tileOpenSocket.UseCustomBackColor = false;
            tileColseSocket.UseCustomBackColor = true;
            tileOpenSocket.Enabled = true;
            tileColseSocket.Enabled = false;
            txtSocket.Enabled = true;
        }

        private void tileListenPort_Click(object sender, EventArgs e)
        {
            listenPort();
            tileListenPort.Enabled = false;
            tileListenPort.UseCustomBackColor = true;
        }

        private void tileWritePort_Click(object sender, EventArgs e)
        {
            writeSerial(txtMsg.Text);
        }

        private void tileResetAccount_Click(object sender, EventArgs e)
        {
            txtUsername.Text = "admin";
            txtPassword.Text = "admin";
        }

        private void tileSaveAccount_Click(object sender, EventArgs e)
        {
            username = txtUsername.Text;
            password = txtPassword.Text;

            saveUserDetails();
            lblStatusBar.Text = "Password Saved";
        }

        private void tileConsoleRead_Click(object sender, EventArgs e)
        {
            writeSerial("read");
        }

        private void tileConsoleClear_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Do you really want to clear the EEPROM?", "Continue?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                writeSerial("clear");
            }
        }

        private void tileConsoleInit_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Do you really want to initialize the EEPROM?\n(Initialization will clear the EEPROM)", "Continue?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                writeSerial("init");
            }
        }

        private void tileConsoleSend_Click(object sender, EventArgs e)
        {
            writeSerial(txtConsoleWrite.Text);
            txtConsoleWrite.Text = "";
            txtConsoleWrite.Focus();
        }

        private void tileTSReset_Click(object sender, EventArgs e)
        {
            txtTSName.Text = "Task";
            dateTimeTSDate.Value = DateTime.Now;
            dateTimeTSTime.Value = DateTime.Now;
            loadXML();
        }

        private void gridTS_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataRow dataRow = dataSetTaskList.Tables[0].Rows[e.RowIndex];

                if (dataRow["on"].ToString().Equals("True"))
                {
                    dataRow["off"] = false;
                }
                else if (dataRow["off"].ToString().Equals("True"))
                {
                    dataRow["on"] = false;
                }
            }
            catch (Exception ex) { logError(ex); }
        }

        private void tileTSAdd_Click(object sender, EventArgs e)
        {
            string line = "";
            foreach (DataRow dataRow in dataSetTaskList.Tables[0].Rows)
            {
                if (dataRow["on"].ToString().Equals("True"))
                {
                    line += "X" + dataRow["childno"].ToString() + "Y1Z";
                }
                else if (dataRow["off"].ToString().Equals("True"))
                {
                    line += "X" + dataRow["childno"].ToString() + "Y0Z";
                }
            }

            if (setTask(txtTSName.Text, dateTimeTSDate.Value.ToString("yyyy-MM-dd"), dateTimeTSTime.Value.ToString("HH:mm:ss"), line))
            {
                lblStatusBar.Text = "Task Added Successfully";
            }
            else
            {
                lblStatusBar.Text = "Failed to Add the Task!";
            }
        }

        private void tileTSSave_Click(object sender, EventArgs e)
        {
            string line = "";
            foreach (DataRow dataRow in dataSetTaskList.Tables[0].Rows)
            {
                if (dataRow["on"].ToString().Equals("True"))
                {
                    line += "X" + dataRow["childno"].ToString() + "Y1Z";
                }
                else if (dataRow["off"].ToString().Equals("True"))
                {
                    line += "X" + dataRow["childno"].ToString() + "Y0Z";
                }
            }

            if (setTask(comboBoxTSTaskList.SelectedItem.ToString(), dateTimeTSDate.Value.ToString("yyyy-MM-dd"), dateTimeTSTime.Value.ToString("HH:mm:ss"), line))
            {
                lblStatusBar.Text = "Task Saved Successfully";
            }
            else
            {
                lblStatusBar.Text = "Failed to Save the Task!";
            }
        }

        private void tileTSDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (deleteTask(comboBoxTSTaskList.SelectedItem.ToString()))
                {
                    loadTasks();
                    lblStatusBar.Text = "Task Deleted Successfully";
                }
                else
                {
                    lblStatusBar.Text = "Failed to Delete Task!";
                }
            }
            catch (Exception ex) { logError(ex); }
        }

        private void tileTSReload_Click(object sender, EventArgs e)
        {
            loadTasks();
        }

        private void comboBoxTSTaskList_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                comboBoxTSTaskList.Text = comboBoxTSTaskList.SelectedItem.ToString();
                TaskFolder taskFolder = taskService.GetFolder("HAS");
                var tasks = taskFolder.Tasks.Where(s => s.Name.Equals(comboBoxTSTaskList.SelectedItem.ToString()));
                DateTime dateTime;
                foreach (Microsoft.Win32.TaskScheduler.Task task in tasks)
                {
                    dateTime = task.NextRunTime;
                    if (dateTime.Date.Equals(DateTime.ParseExact("0001-01-01", "yyyy-MM-dd", cultureInfo)))
                    {

                        DialogResult result = MessageBox.Show("Task has expired do you want to delete the task?", "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Hand);
                        if (result == System.Windows.Forms.DialogResult.Yes)
                        {
                            deleteTask(comboBoxTSTaskList.SelectedItem.ToString());
                            loadTasks();
                        }
                    }
                    else
                    {
                        loadXML();
                        dateTimeTSDate.Value = dateTime.Date;
                        dateTimeTSTime.Value = DateTime.Parse(dateTime.Hour + ":" + dateTime.Minute + ":" + dateTime.Second, cultureInfo);

                        TaskDefinition taskDefinition = task.Definition;
                        txtTSName.Text = comboBoxTSTaskList.SelectedItem.ToString();
                        string line = taskDefinition.Actions.ToString();
                        line = line.Substring(line.LastIndexOf(' ') + 1);

                        string[] lineList = line.Split('X');
                        foreach (string str in lineList)
                        {
                            if (!str.Equals(""))
                            {
                                string command = "";
                                string state = "";
                                command = str.Substring(0, str.IndexOf('Y'));
                                state = str.Substring(str.IndexOf('Y') + 1, 1);
                                foreach (DataRow dataRow in dataSetTaskList.Tables[0].Rows)
                                {
                                    if (dataRow["childno"].ToString().Equals(command))
                                    {
                                        if (state.Equals("1"))
                                        {
                                            dataRow["on"] = true;
                                        }
                                        else if (state.Equals("0"))
                                        {
                                            dataRow["off"] = true;
                                        }

                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.StackTrace); logError(ex); }
        }

        private void txtConsoleWrite_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return)
            {
                writeSerial(txtConsoleWrite.Text);
                txtConsoleWrite.Text = "";
                txtConsoleWrite.Focus();
                e.Handled = true;
            }
        }

        private void mainUI_FormClosing(object sender, FormClosingEventArgs e)
        {
            Environment.Exit(0);
        }

        public void openPort2()
        {
            try
            {
                if (comboBoxPorts2.SelectedItem.ToString() == "Please Select a Port" || comboBoxPorts2.SelectedItem.ToString().Equals(""))
                {
                    lblStatusBar.Text = "Please select a port 2";
                }
                else
                {
                    try
                    {
                        serialPortSecond.PortName = comboBoxPorts2.SelectedItem.ToString();
                        serialPortSecond.Open();
                        comboBoxPorts2.Enabled = false;
                        tileOpenPort2.Enabled = false;
                        tileOpenPort2.UseCustomBackColor = true;
                        tileClosePort2.Enabled = true;
                        tileClosePort2.UseCustomBackColor = false;
                        lblStatusBar.Text = serialPortSecond.PortName + " Port Opened";
                    }
                    catch (Exception ex)
                    {
                        lblStatusBar.Text = serialPortSecond.PortName + " Port is already open";
                        Console.Write(ex.Message);
                        logError(ex);
                    }
                }
            }
            catch (NullReferenceException e) { lblStatusBar.Text = "Please select a port"; logError(e); }
        }

        public void closePort2()
        {
            try
            {
                serialPortSecond.Close();
                if (comboBoxPorts2.InvokeRequired)
                {
                    comboBoxPorts2.Invoke(new MethodInvoker(delegate { comboBoxPorts2.Enabled = true; }));
                }
                if (tileOpenPort2.InvokeRequired)
                {
                    tileOpenPort2.Invoke(new MethodInvoker(delegate
                    {
                        tileOpenPort2.Enabled = true;
                        tileOpenPort2.UseCustomBackColor = false;
                    }));
                }
                if (tileClosePort2.InvokeRequired)
                {
                    tileClosePort2.Invoke(new MethodInvoker(delegate
                    {
                        tileClosePort2.Enabled = false;
                        tileClosePort2.UseCustomBackColor = true;
                    }));
                }
                if (tileListenPort2.InvokeRequired)
                {
                    tileListenPort2.Invoke(new MethodInvoker(delegate
                    {
                        tileListenPort2.Enabled = true;
                        tileListenPort2.UseCustomBackColor = false;
                    }));
                }
                changeLabel(lblStatusBar, serialPortSecond.PortName + " Port Closed");
            }
            catch (Exception ex)
            {
                lblStatusBar.Text = ex.Message;
                logError(ex);
            }
        }

        public void listenPort2()
        {
            new Thread(() =>
            {
                String prevLine = "";
                String curLine = "";

                changeLabel(lblStatusBar, "Listening...");

                while (true)
                {
                    try
                    {
                        if (serialPortSecond.IsOpen)
                        {
                            curLine = serialPortSecond.ReadLine();
                            if (prevLine != curLine)
                            {
                                if (curLine[0] == 'R' && curLine[curLine.Length - 2] == 'N')
                                {
                                    changeLabel(lblOut, curLine);
                                    saveXML(curLine);
                                }
                                if(curLine.StartsWith("Fan On"))
                                {
                                    customArr[0] = true;
                                }else if (curLine.StartsWith("Fan Off"))
                                {
                                    customArr[0] = false;
                                }
                                else if(curLine.StartsWith("Closing Garage Door"))
                                {
                                    customArr[1] = true;
                                }
                                else if (curLine.StartsWith("Opening Garage Door"))
                                {
                                    customArr[1] = true;
                                }
                                else if (curLine.StartsWith("Opening Gate"))
                                {
                                    customArr[2] = true;
                                }
                                else if (curLine.StartsWith("Closing Gate"))
                                {
                                    customArr[2] = true;
                                }
                                else if (curLine.StartsWith("Unlocking Door"))
                                {
                                    customArr[3] = true;
                                }
                                else if (curLine.StartsWith("Locking Door"))
                                {
                                    customArr[3] = true;
                                }
                                else if (curLine.StartsWith("Opening Cloth Line"))
                                {
                                    customArr[4] = true;
                                }
                                else if (curLine.StartsWith("Closing Cloth Line"))
                                {
                                    customArr[4] = true;
                                }
                                else
                                    consoleAddLine(txtConsoleMain, curLine);
                            }
                            prevLine = curLine;
                        }
                        else
                        {
                            closePort2();
                            Thread.CurrentThread.Abort();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.StackTrace);
                        changeLabel(lblStatusBar, ex.Message);
                        changeLabel(lblStatusBar, "Please open a port");
                        logError(ex);
                    }
                }
            }).Start();
        }

        public void writeSerial2(string str)
        {
            try
            {
                if (serialPortSecond.IsOpen)
                {
                    serialPortSecond.Write(str);
                }
                else
                {
                    changeLabel(lblStatusBar, "Pease open a port");
                }
            }
            catch (Exception ex)
            {
                changeLabel(lblStatusBar, ex.Message);
                logError(ex);
            }
        }

        private void tileOpenPort2_Click(object sender, EventArgs e)
        {
            openPort2();
        }

        private void tileClosePort2_Click(object sender, EventArgs e)
        {
            closePort2();
        }

        private void tileListenPort2_Click(object sender, EventArgs e)
        {
            listenPort2();
            tileListenPort2.Enabled = false;
            tileListenPort2.UseCustomBackColor = true;
        }

        private void tileWritePort2_Click(object sender, EventArgs e)
        {
            writeSerial2(txtMsg2.Text);
        }
    }
}
