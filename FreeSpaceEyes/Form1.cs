using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using Telerik.WinControls.UI;
using SnmpSharpNet;
using System.Net;
using System.Collections;
using System.Data.OleDb;
using System.Threading;
using System.Management;
using System.Media;


namespace FreeSpaceEyes
{
    public partial class Form1 : RadForm
    {
        public Form1()
        {
            InitializeComponent();
        }
        //宣告最大thread and IO and Timer
        //讀取systemSetting
        int MaxThread = 200;
        int MaxIOThread = 200;
        int LimitNonPage = 100;
        int LimitTimeRange = 10;
        int timerCheck = 1000000;//預設值
        private void Form1_Load(object sender, EventArgs e)
        {
            //讀取systemSetting
            //====開始讀取系統設定====
            ReloadSystemSetting();
        }

        private void ReloadSystemSetting()
        {
            String strSQLSystem = "SELECT * FROM [SystemSetting] WHERE id=1";
            System.Data.OleDb.OleDbConnection oleConnSystem =
                 new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=FreeSpaceEyesDB.mdb");
            oleConnSystem.Open();

            System.Data.OleDb.OleDbCommand oleCmdSystem = new System.Data.OleDb.OleDbCommand(strSQLSystem, oleConnSystem);
            OleDbDataReader thisReaderSystem = oleCmdSystem.ExecuteReader();//建立OleDbDataReader用來存放從Access裡面讀取出來的
            while (thisReaderSystem.Read())
            {
                timerCheck = int.Parse(thisReaderSystem["Timer"].ToString());
                MaxThread = int.Parse(thisReaderSystem["MaxThread"].ToString());
                MaxIOThread = int.Parse(thisReaderSystem["MaxIOThread"].ToString());
                LimitNonPage = int.Parse(thisReaderSystem["LimitNonPage"].ToString());
                LimitTimeRange = int.Parse(thisReaderSystem["LimitTimeRange"].ToString());
            }
            oleConnSystem.Close();
            TimerForCheck.Interval = timerCheck*1000*60*60;//更改間隔 小時轉毫秒
        }

        private void ToolMenuCheckBox_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            //是否顯是工具列
            if (ToolMenuCheckBox.Checked==true)
            {
                Menupanel.Visible = true;
                Menupanel.BringToFront();
            }
            else
            {
                Menupanel.Visible = false;
            }
        }
        //======檢查程式開始=======
            //檢查按鈕
        private void CheckNow_btn_Click(object sender, EventArgs e)
        {
            //啟動檢察
            DoCheck();
            CheckNow_btn.Enabled = false;

        }

        private void DoCheck()
        {
            CheckNow_btn.Text = "執行中";
            //重置報表
            CheckReset();
            //一次性讀取access內資料
            //重讀一次設定
            ReloadSystemSetting();
            //===開始讀取Windows===
            String strSQL = "SELECT * FROM [WindowsTarget]";
            System.Data.OleDb.OleDbConnection oleConn =
                 new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=FreeSpaceEyesDB.mdb");
            oleConn.Open();

            System.Data.OleDb.OleDbCommand oleCmd = new System.Data.OleDb.OleDbCommand(strSQL, oleConn);
            OleDbDataReader thisReader = oleCmd.ExecuteReader();//建立OleDbDataReader用來存放從Access裡面讀取出來的
            while (thisReader.Read())
            {
                //存放
                string[] AskList = new string[10];
                //多執行緒 天嘉進asllist
                AskList[0] = thisReader["TargetIP"].ToString();
                AskList[1] = thisReader["TargetName"].ToString();
                AskList[2] = thisReader["TargetDomain"].ToString();
                AskList[3] = thisReader["TargetUser"].ToString();
                AskList[4] = thisReader["TargetPassword"].ToString();
                AskList[5] = thisReader["TargetDeviceID"].ToString();
                AskList[6] = thisReader["TargetAlert"].ToString();
                AskList[7] = thisReader["TargetSpaceLog"].ToString();
                AskList[8] = thisReader["id"].ToString();
                AskList[9] = thisReader["TargetWarning"].ToString();
                //多執行緒設定 並啟動檢查硬碟
                ThreadPool.SetMaxThreads(MaxThread, MaxIOThread);
                ThreadPool.SetMinThreads(MaxThread, MaxIOThread);
                ThreadPool.QueueUserWorkItem(new WaitCallback(CheckWindowsDiskThread), AskList);
            }
            //==讀取Windows結束==




            //===開始讀取Linux===
            strSQL = "SELECT * FROM [LinuxTarget]";
            System.Data.OleDb.OleDbCommand oleCmd2 = new System.Data.OleDb.OleDbCommand(strSQL, oleConn);
            OleDbDataReader thisReader2 = oleCmd2.ExecuteReader();//建立OleDbDataReader用來存放從Access裡面讀取出來的
            while (thisReader2.Read())
            {
                //存放
                string[] AskList = new string[6];
                //多執行緒 天嘉進asllist
                AskList[0] = thisReader2["TargetName"].ToString();
                AskList[1] = thisReader2["TargetIP"].ToString();
                AskList[2] = thisReader2["TargetAlert"].ToString();
                AskList[3] = thisReader2["TargetSpaceLog"].ToString();
                AskList[4] = thisReader2["id"].ToString();
                AskList[5] = thisReader2["TargetWarning"].ToString();
                //多執行緒設定 並啟動
                ThreadPool.QueueUserWorkItem(new WaitCallback(CheckLinuxDiskThread), AskList);
            }

            //==讀取Linux結束==
            //==讀取Windows NonPage Check List==
            strSQL = "SELECT * FROM [WindowsTarget] WHERE [TargetNonPageCheck] = true";
            System.Data.OleDb.OleDbCommand oleCmd3 = new System.Data.OleDb.OleDbCommand(strSQL, oleConn);
            OleDbDataReader thisReader3 = oleCmd3.ExecuteReader();//建立OleDbDataReader用來存放從Access裡面讀取出來的
            while (thisReader3.Read())
            {
                //存放
                string[] AskList = new string[10];
                //多執行緒 天嘉進asllist
                AskList[0] = thisReader3["TargetIP"].ToString();
                AskList[1] = thisReader3["TargetName"].ToString();
                AskList[2] = thisReader3["TargetDomain"].ToString();
                AskList[3] = thisReader3["TargetUser"].ToString();
                AskList[4] = thisReader3["TargetPassword"].ToString();
                AskList[5] = thisReader3["TargetDeviceID"].ToString();
                AskList[6] = thisReader3["TargetAlert"].ToString();
                AskList[7] = thisReader3["TargetSpaceLog"].ToString();
                AskList[8] = thisReader3["id"].ToString();
                AskList[9] = thisReader3["TargetWarning"].ToString();

                if (NonPageCheckBox.Checked) //檢查ui上的選項是否勾選
                {
                    //啟動多執行續檢查nonpage
                    ThreadPool.QueueUserWorkItem(new WaitCallback(CheckWindowsNonPageThread), AskList);
                }
            }
            //==讀取Windows NonPage Check List結束==

            //==讀取Windows Time Check List==
            strSQL = "SELECT * FROM [WindowsTarget] WHERE [TargetTimeCheck] = true";
            System.Data.OleDb.OleDbCommand oleCmd4 = new System.Data.OleDb.OleDbCommand(strSQL, oleConn);
            OleDbDataReader thisReader4 = oleCmd4.ExecuteReader();//建立OleDbDataReader用來存放從Access裡面讀取出來的
            while (thisReader4.Read())
            {
                //存放
                string[] AskList = new string[10];
                //多執行緒 天嘉進asllist
                AskList[0] = thisReader4["TargetIP"].ToString();
                AskList[1] = thisReader4["TargetName"].ToString();
                AskList[2] = thisReader4["TargetDomain"].ToString();
                AskList[3] = thisReader4["TargetUser"].ToString();
                AskList[4] = thisReader4["TargetPassword"].ToString();
                AskList[5] = thisReader4["TargetDeviceID"].ToString();
                AskList[6] = thisReader4["TargetAlert"].ToString();
                AskList[7] = thisReader4["TargetSpaceLog"].ToString();
                AskList[8] = thisReader4["id"].ToString();
                AskList[9] = thisReader4["TargetWarning"].ToString();

                if (TimeCheckBox.Checked) //檢查ui上的選項是否勾選
                {
                    //啟動多執行續檢查nonpage
                    ThreadPool.QueueUserWorkItem(new WaitCallback(CheckWindowsTimeThread), AskList);
                }
            }
            //==讀取Windows Time Check List結束==
            //// 關閉資料庫連接。  
            thisReader.Close();
            oleConn.Close();
        }

        //用來存放顯示目前進度的int
        int NowCount = 0;
        int MaxCount = 0;
        private void CheckReset()
        {
            //將音效取消
            warn_vol.Checked = false;
            warn_vol.Enabled = false;
            //將音效check紐關閉
            AlreadyGridView.Rows.Clear();
            AlertGridView.Rows.Clear();
            ErrorList.Items.Clear();
            NonpageGridView.Rows.Clear();
            TimeGridView.Rows.Clear();
            //光棒歸零
            ProgressBar.Value1 = 0;
            int max = 0; //宣告最大值
            //取出總數
            //將資料拉出來到ui上
            // === 對 Access 資料庫下SQL語法 ===  
            // Transact-SQL 陳述式  
            String strSQL = "SELECT count(*) FROM [WindowsTarget]";
            System.Data.OleDb.OleDbConnection oleConn =
                 new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=FreeSpaceEyesDB.mdb");
            //// 開啟資料庫連接。  
            oleConn.Open();
            System.Data.OleDb.OleDbCommand oleCmd =
                new System.Data.OleDb.OleDbCommand(strSQL, oleConn);
            OleDbDataReader thisReader = oleCmd.ExecuteReader();//建立OleDbDataReader用來存放從Access裡面讀取出來的
            //拉到ui
            while (thisReader.Read())
            {
                max = int.Parse(thisReader[0].ToString());
            }
            strSQL = "SELECT count(*) FROM [LinuxTarget]";
            System.Data.OleDb.OleDbCommand oleCmd2 =
                            new System.Data.OleDb.OleDbCommand(strSQL, oleConn);
            thisReader = oleCmd2.ExecuteReader();//建立OleDbDataReader用來存放從Access裡面讀取出來的
            while (thisReader.Read())
            {
                max += int.Parse(thisReader[0].ToString());
            }
            //計算nonpage count
            if (NonPageCheckBox.Checked)
            {
                strSQL = "SELECT count(*) FROM [WindowsTarget] WHERE [TargetNonPageCheck] = true";
                System.Data.OleDb.OleDbCommand oleCmd3 =
                                new System.Data.OleDb.OleDbCommand(strSQL, oleConn);
                thisReader = oleCmd3.ExecuteReader();//建立OleDbDataReader用來存放從Access裡面讀取出來的
                while (thisReader.Read())
                {
                    max += int.Parse(thisReader[0].ToString());
                }
            }
            //計算TIME count
            if (TimeCheckBox.Checked)
            {
                strSQL = "SELECT count(*) FROM [WindowsTarget] WHERE [TargetTimeCheck] = true";
                System.Data.OleDb.OleDbCommand oleCmd4 =
                                new System.Data.OleDb.OleDbCommand(strSQL, oleConn);
                thisReader = oleCmd4.ExecuteReader();//建立OleDbDataReader用來存放從Access裡面讀取出來的
                while (thisReader.Read())
                {
                    max += int.Parse(thisReader[0].ToString());
                }
            }
            //// 關閉資料庫連接。  
            thisReader.Close();
            oleConn.Close();
            ProgressBar.Maximum = max;//將所有長度嘉進光棒
            MaxCount = max;
        }
            //多執行緒檢察windows
        private void CheckWindowsDiskThread(object obj)
        {
            //轉型傳過來的obj
            string[] Askitem = new string[10];
            Askitem = (string[])obj;
            //0=TargetIP , 1 = TargetName, 2= TargetDomain ,3=TargetUser ,4=TargetPassword ,5=TargetDeviceID ,6=TargetAlert ,7=TargetSpaceLog ,8=id , 9=TargetWarning
            //開始檢察
            try
            {
                //這邊進行wmi查詢目標硬碟代號
                System.Management.ConnectionOptions Conn = new ConnectionOptions();
                //設定用於WMI連接操作的用戶名
                string user = Askitem[2] + "\\" + Askitem[3];
                Conn.Username = user;
                //設定用戶的密碼
                Conn.Password = Askitem[4];
                //設定用於執行WMI操作的範圍
                System.Management.ManagementScope Ms = new ManagementScope("\\\\" + Askitem[0] + "\\root\\cimv2", Conn);
                Ms.Connect();
                ObjectQuery Query = new ObjectQuery("select FreeSpace ,Size from Win32_LogicalDisk where DriveType=3 and DeviceID='" + Askitem[5] + "'");
                //WQL語句，設定的WMI查詢內容和WMI的操作範圍，檢索WMI對象集合
                ManagementObjectSearcher Searcher = new ManagementObjectSearcher(Ms, Query);
                //異步調用WMI查詢
                ManagementObjectCollection ReturnCollection = Searcher.Get();
                //設定接取的變數
                double AllMB;
                int AllGB;
                double FreeMB;
                int FreeGB;
                string[] text = new string[10];
                foreach (ManagementObject Return in ReturnCollection)
                {
                    AllMB = Convert.ToInt64(Return["size"]) / 1048576;//1024*1024
                    AllGB = (int)AllMB / 1024;
                    FreeMB = Convert.ToInt64(Return["FreeSpace"]) / 1048576;
                    FreeGB = (int)FreeMB / 1024;
                    text[0] = Askitem[1];//名稱
                    text[1] = Askitem[5];//代號
                    text[2] = AllMB.ToString();//總空間mb
                    text[3] = AllGB.ToString();//總空間gb
                    text[4] = FreeMB.ToString();//剩餘mb
                    text[5] = FreeGB.ToString();//剩餘gb
                    text[6] = Askitem[6];//警戒值
                    text[7] = "Windows";//檢察windows
                    text[8] = Askitem[8];//WindowsTargetID
                    text[9] = Askitem[9];//警報值
                    //如果低於警戒值
                    if ((int)FreeMB <= Convert.ToInt32(Askitem[6]))
                    {
                        this.Invoke(new D_AddAlertGridView(AddAlertGridView), new object[] { text });
                    }
                }
                //委派傳出
                this.Invoke(new D_AddAlreadyGridView(AddAlreadyGridView), new object[] { text });
                //警告訊息

            }
            catch (Exception x)
            {
                this.Invoke(new D_AddErrorList(AddErrorList), Askitem[1] + " >>> " + x.Message.ToString());
            }
            this.Invoke(new D_AddProgress(AddProgress), null);
        }

        //多執行緒檢察windows NonPage Time 3.2版新增
        private void CheckWindowsNonPageThread(object obj)
        {
            //轉型傳過來的obj
            string[] Askitem = new string[10];
            Askitem = (string[])obj;
            //0=TargetIP , 1 = TargetName, 2= TargetDomain ,3=TargetUser ,4=TargetPassword ,5=TargetDeviceID ,6=TargetAlert ,7=TargetSpaceLog ,8=id , 9=TargetWarning
            //開始檢察
            try
            {
                //這邊進行wmi查詢目標硬碟代號
                System.Management.ConnectionOptions Conn = new ConnectionOptions();
                //設定用於WMI連接操作的用戶名
                string user = Askitem[2] + "\\" + Askitem[3];
                Conn.Username = user;
                //設定用戶的密碼
                Conn.Password = Askitem[4];
                //設定用於執行WMI操作的範圍
                System.Management.ManagementScope Ms = new ManagementScope("\\\\" + Askitem[0] + "\\root\\cimv2", Conn);
                Ms.Connect();
                ObjectQuery Query = new ObjectQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Memory");
                //WQL語句，設定的WMI查詢內容和WMI的操作範圍，檢索WMI對象集合
                ManagementObjectSearcher Searcher = new ManagementObjectSearcher(Ms, Query);
                //異步調用WMI查詢
                ManagementObjectCollection ReturnCollection = Searcher.Get();
                //設定接收的變數
                int NonBytesInt;
                string[] text = new string[4];
                foreach (ManagementObject Return in ReturnCollection)
                {
                    NonBytesInt = (int.Parse(Return["PoolNonpagedBytes"].ToString()) / 1024 / 1024);
                    text[0] = Askitem[1];//名稱
                    text[1] = NonBytesInt.ToString();//NonPageBytesInt
                    text[2] = "Windows";//檢察windows
                    text[3] = Askitem[8];//WindowsTargetID
                    if (NonBytesInt >= LimitNonPage)
                    {
                        this.Invoke(new D_AddNonPageGridView(AddNonPageGridView), new object[] { text });
                    }
                }
            }
            catch (Exception x)
            {
                this.Invoke(new D_AddErrorList(AddErrorList), Askitem[1] + " >>NonPageCheckError>> " + x.Message.ToString());
            }
            this.Invoke(new D_AddProgress(AddProgress), null);
        }

        //多執行緒檢察windows Time Time 3.2版新增
        private void CheckWindowsTimeThread(object obj)
        {
            //轉型傳過來的obj
            string[] Askitem = new string[10];
            Askitem = (string[])obj;
            //0=TargetIP , 1 = TargetName, 2= TargetDomain ,3=TargetUser ,4=TargetPassword ,5=TargetDeviceID ,6=TargetAlert ,7=TargetSpaceLog ,8=id , 9=TargetWarning
            //開始檢察
            try
            {
                //這邊進行wmi查詢目標硬碟代號
                System.Management.ConnectionOptions Conn = new ConnectionOptions();
                //設定用於WMI連接操作的用戶名
                string user = Askitem[2] + "\\" + Askitem[3];
                Conn.Username = user;
                //設定用戶的密碼
                Conn.Password = Askitem[4];
                //設定用於執行WMI操作的範圍
                System.Management.ManagementScope Ms = new ManagementScope("\\\\" + Askitem[0] + "\\root\\cimv2", Conn);
                Ms.Connect();
                ObjectQuery Query = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
                //WQL語句，設定的WMI查詢內容和WMI的操作範圍，檢索WMI對象集合
                ManagementObjectSearcher Searcher = new ManagementObjectSearcher(Ms, Query);
                //異步調用WMI查詢
                ManagementObjectCollection ReturnCollection = Searcher.Get();
                DateTime NowCheckTime = DateTime.Now;
                //設定接收的變數
                string[] text = new string[6];
                foreach (ManagementObject Return in ReturnCollection)
                {
                    
                    string UTCTime = Return["LocalDateTime"].ToString();
                    string[] sub = UTCTime.Split('+', '.');
                    DateTime NowBackTime = DateTime.ParseExact(sub[0], "yyyyMMddHHmmss", null);
                    //計算相差時間
                    TimeSpan ts = NowCheckTime - NowBackTime;
                    ts = ts.Duration();//絕對值

                    text[0] = Askitem[1];//名稱
                    text[1] = NowBackTime.ToString("yyyy MM/dd HH:mm:ss");//TimeInt
                    text[2] = ts.Seconds.ToString();
                    //判定是否正常
                    if (ts.Seconds >= LimitTimeRange)
                    {
                        text[3] = "異常";
                    }
                    else
                    {
                        text[3] = "ok";
                    }
                    text[4] = "Windows";//檢察windows
                    text[5] = Askitem[8];//WindowsTargetID
                    this.Invoke(new D_AddTimeGridView(AddTimeGridView), new object[] { text });
                }
            }
            catch (Exception x)
            {
                this.Invoke(new D_AddErrorList(AddErrorList), Askitem[1] + " >>TimeCheckError>> " + x.Message.ToString());
            }
            this.Invoke(new D_AddProgress(AddProgress), null);
        }



            //多執行緒檢察linux
        private void CheckLinuxDiskThread(object obj)
        {
            //0 = TargetName, 1= TargetIP ,2=TargetAlert ,3=TargetSpaceLog ,4=id ,5=TargetWarning
            string[] Askitem = new string[6];
            Askitem = (string[])obj;
            try
            {
                // SNMP community name
                OctetString community = new OctetString("public");
                AgentParameters param = new AgentParameters(community);
                param.Version = SnmpVersion.Ver1;
                IpAddress agent = new IpAddress(Askitem[1]);//TargetIP
                // Construct target
                UdpTarget target = new UdpTarget((IPAddress)agent, 161, 2000, 1);
                // Pdu class used for all requests
                Pdu pdu = new Pdu(PduType.Get);
                pdu.VbList.Add(".1.3.6.1.4.1.2021.9.1.2.1");//代號
                pdu.VbList.Add(".1.3.6.1.4.1.2021.9.1.6.1");//總空間
                pdu.VbList.Add(".1.3.6.1.4.1.2021.9.1.8.1");//已用空間
                // Make SNMP request
                SnmpV1Packet result = (SnmpV1Packet)target.Request(pdu, param);
                //取得 總空間kb 已用空間kb 分割區帶號
                int AllKB = int.Parse(result.Pdu.VbList[1].Value.ToString());
                int UseKB = int.Parse(result.Pdu.VbList[2].Value.ToString());
                int FreeKB = AllKB - UseKB;
                int AllMB = AllKB / 1024;
                int FreeMB = FreeKB / 1024;
                int AllGB = AllMB / 1024;
                int FreeGB = FreeMB / 1024;
                string ID = result.Pdu.VbList[0].Value.ToString();
                //委派傳出
                string[] text = new string[10];
                text[0] = Askitem[0];//名稱
                text[1] = ID;//代號
                text[2] = AllMB.ToString();//總空間mb
                text[3] = AllGB.ToString();//總空間gb
                text[4] = FreeMB.ToString();//剩餘mb
                text[5] = FreeGB.ToString();//剩餘gb
                text[6] = Askitem[2];//警戒值
                text[7] = "Linux";//檢察linux
                text[8] = Askitem[4];//WindowsTargetID
                text[9] = Askitem[5];//TargetWaning
                this.Invoke(new D_AddAlreadyGridView(AddAlreadyGridView), new object[] { text });
                //如果低於警戒值
                if (FreeMB<=Convert.ToInt32(Askitem[2]))
                {
                    this.Invoke(new D_AddAlertGridView(AddAlertGridView), new object[] { text });
                }
            }
            catch (Exception x)
            {
                this.Invoke(new D_AddErrorList(AddErrorList), Askitem[0]+" >>> "+x.Message.ToString());
            }
            this.Invoke(new D_AddProgress(AddProgress), null);
        }
        //委派用addAlertGridView
        delegate void D_AddAlertGridView(object[] obj);
        private void AddAlertGridView(object[] obj)
        {
            string[] item = (string[])obj;
            //檢查是否觸發警報 並寫到ui上
            if (int.Parse(item[4]) < int.Parse(item[9]))
            {
                //聲音啟動
                warn_vol.Checked = true;
                warn_vol.Enabled = true;
                AlertGridView.Rows.Add(item[0], item[1], item[2] + " MB (" + item[3] + " GB)", item[4] + " MB (" + item[5] + " GB)", item[6] + " MB ！", item[7], item[8]);
            }
            else
            {
                AlertGridView.Rows.Add(item[0], item[1], item[2] + " MB (" + item[3] + " GB)", item[4] + " MB (" + item[5] + " GB)", item[6] + " MB", item[7], item[8]);
            }
            
        }
        
        //委派用addNonPageGridView
        delegate void D_AddNonPageGridView(object[] obj);
        private void AddNonPageGridView(object[] obj)
        {
            string[] item = (string[])obj;
            NonpageGridView.Rows.Add(item[0],item[1]+" (MB)",item[2],item[3]);
        }

        //委派用addTimeGridView
        delegate void D_AddTimeGridView(object[] obj);
        private void AddTimeGridView(object[] obj)
        {
            string[] item = (string[])obj;
            TimeGridView.Rows.Add(item[0], item[1], item[2], item[3], item[4], item[5]);
        }

        //委派用addAlreadyGridView
        delegate void D_AddAlreadyGridView(object[] obj);
        private void AddAlreadyGridView(object[] obj)
        {
            string[] item = (string[])obj;
            AlreadyGridView.Rows.Add(item[0], item[1], item[2] + " MB (" + item[3] + " GB)", item[4] + " MB (" + item[5] + " GB)", item[6] + " MB", item[7], item[8]);
        }
        //委派用addErrorList
        delegate void D_AddErrorList(string text);
        private void AddErrorList(string text)
        {
            ErrorList.Items.Add(text);
        }
        //委派用D_AddProgress
        delegate void D_AddProgress();
        private void AddProgress()
        {
            NowCount += 1;
            ProgressBar.Value1 += ProgressBar.Step;
            //所有完成時
            RadPageView.Pages[0].Text = "警戒列表 (" + AlertGridView.Rows.Count().ToString() + ")";
            RadPageView.Pages[1].Text = "過檢列表 (" + AlreadyGridView.Rows.Count().ToString() + ")";
            RadPageView.Pages[2].Text = "錯誤列表 (" + ErrorList.Items.Count + ")";
            RadPageView.Pages[3].Text = "NonPage警戒 (" + NonpageGridView.Rows.Count().ToString() + ")";
            RadPageView.Pages[4].Text = "時間警戒 (" + TimeGridView.Rows.Count().ToString() + ")";
            if (NowCount == MaxCount)
            {
                CheckNow_btn.Enabled = true;
                CheckNow_btn.Text = "啟動檢查";
                NowCount = 0;
                LastCheckTime.Text = "最後檢查時間："+DateTime.Now.ToString();
            }
        }
        //=====檢查程式結束=====
        //=====聲音部分=====
        private SoundPlayer Player = new SoundPlayer();
        private void warn_vol_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            if (warn_vol.Checked==true)
            {
                try
                {
                    this.BackColor = System.Drawing.Color.MistyRose;
                    this.Player.SoundLocation = "ChatBeep.wav";
                    this.Player.PlayLooping();
                }
                catch (Exception)
                {
                } 
            }
            else
            {
                try
                {
                    this.BackColor = System.Drawing.Color.White;
                    this.Player.Stop();
                }
                catch (Exception)
                {
                }
            }
        }

        //=====聲音部分結束=====

        //=====定時偵測開始=====
        private void always_check_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            if (always_check.Checked ==true)
            {
                TimerForCheck.Enabled = true;
            }
            else
            {
                TimerForCheck.Enabled = false;
            }
        }
        private void TimerForCheck_Tick(object sender, EventArgs e)
        {
            //啟動檢察
            DoCheck();
            CheckNow_btn.Enabled = false;
        }

        //====定時偵測結束======


        //=====工具列啟動用=====
        //啟動windowsSetting 起始
        private void WindowsSetting_Click(object sender, EventArgs e)
        {
            Form winsetting = new WinSetting();
            winsetting.Show();
            this.Enabled = false;
            //子視窗關閉程式時 父視窗顯示
            winsetting.FormClosed += new FormClosedEventHandler(winsetting_FormClosed);
        }

        void winsetting_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Enabled = true;
        }
        //啟動linuxSetting 起始
        private void LinuxSetting_Click(object sender, EventArgs e)
        {
            Form linuxsetting = new LinuxSetting();
            linuxsetting.Show();
            this.Enabled = false;
            //子視窗關閉程式時 父視窗顯示
            linuxsetting.FormClosed += new FormClosedEventHandler(linuxsetting_FormClosed);
        }
        void linuxsetting_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Enabled = true;
        }
        //啟動linuxSetting 關閉
        //啟動SystemSetting 起始
        private void SystemSetting_Click(object sender, EventArgs e)
        {
            Form systemsetting = new SystemSetting();
            systemsetting.Show();
            this.Enabled = false;
            //子視窗關閉程式時 父視窗顯示
            systemsetting.FormClosed += new FormClosedEventHandler(systemsetting_FormClosed);
        }
        void systemsetting_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Enabled = true;
            ReloadSystemSetting();
        }
        //啟動SystemSetting 關閉



        //=======關於又見選單=======
            //關閉預設的選單
        private void AlertGridView_ContextMenuOpening(object sender, ContextMenuOpeningEventArgs e)
        {
            e.Cancel = true;
        }

        private void AlertGridView_MouseDown(object sender, MouseEventArgs e)
        {
            // && AlertGridView.Rows.Count()!=0
            if (e.Button == MouseButtons.Right && AlertGridView.Rows.Count() != 0)
            {
                AlertMenu.Show(Cursor.Position.X, Cursor.Position.Y);
               // ChangeResText();
            }

        }

        //複製選項
        private void Copy_Click(object sender, EventArgs e)
        {
            //取index
            int index = AlertGridView.CurrentCell.RowIndex;
            //取值
            string text = AlertGridView.Rows[index].Cells[0].Value.ToString();
            text += "  ";
            text += AlertGridView.Rows[index].Cells[1].Value.ToString();
            text += "\\  剩餘 ";
            text += AlertGridView.Rows[index].Cells[3].Value.ToString();
            // 複製進剪貼簿
            Clipboard.SetData(DataFormats.Text, text);
        }



        //=======關於又見選單結束=======
        //======警告視窗======
        private void NonPageCheckBox_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            if (NonPageCheckBox.Checked)
            {
                MessageBox.Show("開啟NonPage與時間檢查時，可能會降低系統執行效能。\n並於檢查前確保各目標設定值中\nNonPage & Time Check 選項已被勾選。");    
            }
        }

        private void TimeCheckBox_ToggleStateChanged(object sender, StateChangedEventArgs args)
        {
            if (TimeCheckBox.Checked)
            {
                MessageBox.Show("開啟NonPage與時間檢查時，可能會降低系統執行效能。\n並於檢查前確保各目標設定值中\nNonPage & Time Check 選項已被勾選。");
            }
        }
        //======警告視窗結束=====

    }
}
