using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using System.Management;
using Telerik.WinControls.UI;
using System.Collections;
using System.Data.OleDb;
using System.Threading;
using System.Net;

namespace FreeSpaceEyes
{
    public partial class WinSetting : Telerik.WinControls.UI.RadForm
    {
        public WinSetting()
        {
            InitializeComponent();
        }

        private void WinSetting_Load(object sender, EventArgs e)
        {
            TargetListView.Items.Clear();
            Reload_btn.Enabled = false;
            radWaitingBar1.StartWaiting();
            //啟動背景讀取 送出使用者要找的文字
            LoadAccess.RunWorkerAsync(find_text.Text);
        }
        //=======項目設定======== 開始
        //重新整理按鈕
        private void Reload_btn_Click(object sender, EventArgs e)
        {
            TargetListView.Items.Clear();
            Reload_btn.Enabled = false;
            radWaitingBar1.StartWaiting();
            //啟動背景讀取 送出使用者要找的文字
            LoadAccess.RunWorkerAsync(find_text.Text);
        }
        //背景讀取access
        private void LoadAccess_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                string find_text = (string)e.Argument;
                String strSQL = "SELECT id,TargetName,TargetDeviceID FROM [WindowsTarget] WHERE TargetName LIKE '%" + find_text + "%'";
                System.Data.OleDb.OleDbConnection oleConn =
                     new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=FreeSpaceEyesDB.mdb");
                oleConn.Open();
                System.Data.OleDb.OleDbCommand oleCmd =
                    new System.Data.OleDb.OleDbCommand(strSQL, oleConn);
                OleDbDataReader thisReader = oleCmd.ExecuteReader();
                while (thisReader.Read())
                {
                    string[] item = new string[3];
                    item[0] = thisReader["id"].ToString();
                    item[1] = thisReader["TargetName"].ToString();
                    item[2] = thisReader["TargetDeviceID"].ToString();
                    this.Invoke(new D_AddList(addlist), (object)item);
                    Thread.Sleep(5);
                }
                thisReader.Close();
                oleConn.Close();
            }
            catch (Exception)
            {
                
            }
        }
        //背景讀取access委派重新整理用
        delegate void D_AddList(object item);
        private void addlist(object item)
        {
            //寫入targetlistview 0=id,1=targetname,2=targetdeviceID
            string[] Temp = (string[])item;
            TargetListView.Items.Add(Temp[0],Temp[1],Temp[2]);
        }
        //背景執行結束後停止跑條
        private void LoadAccess_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Reload_btn.Enabled = true;
            radWaitingBar1.StopWaiting();
        }
        //帶出選擇的項目
        //當選擇的被典籍時
        private void TargetListView_ItemMouseClick(object sender, ListViewItemEventArgs e)
        {
            ListViewDataItem item = TargetListView.SelectedItem;
            //將資料拉出來到ui上
            // === 對 Access 資料庫下SQL語法 ===  
            //// Transact-SQL 陳述式  
            String strSQL = "SELECT id,TargetName,TargetIP,TargetDomain,TargetUser,TargetPassword,TargetDeviceID,TargetAlert,TargetRes,TargetSpaceLog,TargetWarning,TargetNonPageCheck,TargetTimeCheck FROM [WindowsTarget] WHERE id =" + int.Parse(item[0].ToString());
            System.Data.OleDb.OleDbConnection oleConn =
                 new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=FreeSpaceEyesDB.mdb");
            //// 開啟資料庫連接。  
            oleConn.Open();
            //// OleDbCommand (String, OleDbConnection) : 使用查詢的文字和 OleDbConnection，初始化 OleDbCommand 類別的新執行個體。  
            System.Data.OleDb.OleDbCommand oleCmd =
                new System.Data.OleDb.OleDbCommand(strSQL, oleConn);
            OleDbDataReader thisReader = oleCmd.ExecuteReader();//建立OleDbDataReader用來存放從Access裡面讀取出來的
            //拉到ui
            while (thisReader.Read())
            {
                ShowTargetID.Text = thisReader["id"].ToString();
                ShowTargetName.Text = thisReader["TargetName"].ToString();
                ShowTargetIP.Text = thisReader["TargetIP"].ToString();
                ShowTargetDomain.Text = thisReader["TargetDomain"].ToString();
                ShowTargetUser.Text = thisReader["TargetUser"].ToString();
                ShowTargetPassword.Text = thisReader["TargetPassword"].ToString();
                ShowTargetDeviceID.Text = thisReader["TargetDeviceID"].ToString();
                ShowTargetAlert.Text = thisReader["TargetAlert"].ToString();
                ShowTargetRes.Text = thisReader["TargetRes"].ToString();
                ShowTargetWarning.Text = thisReader["TargetWarning"].ToString();
                ShowNonPageCheckBox.Checked = (Boolean)thisReader["TargetNonPageCheck"];
                ShowTimeCheckBox.Checked = (Boolean)thisReader["TargetTimeCheck"];
            }
            //// 關閉資料庫連接。  
            thisReader.Close();
            oleConn.Close();
            //啟動更新即刪除
            updata_btn.Enabled = true;
            deldata_btn.Enabled = true;
        }

        //更新項目明細
        private void updata_btn_Click(object sender, EventArgs e)
        {
            // === 對 Access 資料庫下SQL語法 ===  
            //// Transact-SQL 陳述式  
            String strSQL = "UPDATE [WindowsTarget] SET TargetName = '" + ShowTargetName.Text + "' ,TargetIP = '" + ShowTargetIP.Text + "' ,TargetDomain = '" + ShowTargetDomain.Text + "' ,TargetUser = '" + ShowTargetUser.Text + "' ,TargetPassword = '" + ShowTargetPassword.Text + "' ,TargetAlert = '" + ShowTargetAlert.Text + "' , TargetRes='" + ShowTargetRes.Text + "' , TargetWarning='" + ShowTargetWarning.Text + "', TargetNonPageCheck = "+ ShowNonPageCheckBox.Checked +", TargetTimeCheck = "+ ShowTimeCheckBox.Checked +" WHERE id=" + int.Parse(ShowTargetID.Text);
            System.Data.OleDb.OleDbConnection oleConn =
                 new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=FreeSpaceEyesDB.mdb");
            //// 開啟資料庫連接。  
            oleConn.Open();
            //// OleDbCommand (String, OleDbConnection) : 使用查詢的文字和 OleDbConnection，初始化 OleDbCommand 類別的新執行個體。  
            System.Data.OleDb.OleDbCommand oleCmd =
                new System.Data.OleDb.OleDbCommand(strSQL, oleConn);
            oleCmd.ExecuteNonQuery();
            //// 關閉資料庫連接。  
            oleConn.Close();
        }
        //刪除項目
        private void deldata_btn_Click(object sender, EventArgs e)
        {
            // === 對 Access 資料庫下SQL語法 ===  
            //// Transact-SQL 陳述式  
            String strSQL = "DELETE FROM [WindowsTarget] WHERE id =" + int.Parse(ShowTargetID.Text);
            System.Data.OleDb.OleDbConnection oleConn =
                 new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=FreeSpaceEyesDB.mdb");
            //// 開啟資料庫連接。  
            oleConn.Open();
            //// OleDbCommand (String, OleDbConnection) : 使用查詢的文字和 OleDbConnection，初始化 OleDbCommand 類別的新執行個體。  
            System.Data.OleDb.OleDbCommand oleCmd =
                new System.Data.OleDb.OleDbCommand(strSQL, oleConn);
            oleCmd.ExecuteNonQuery();
            //// 關閉資料庫連接。  
            oleConn.Close();
            //重設ui 並且青空 list
            ShowTargetID.Text = "_";
            ShowTargetName.Text = "";
            ShowTargetIP.Text = "";
            ShowTargetDomain.Text = "";
            ShowTargetUser.Text = "";
            ShowTargetPassword.Text = "";
            ShowTargetDeviceID.Text = "_";
            ShowTargetAlert.Text = "";
            ShowTargetRes.Text = "";
            ShowTargetWarning.Text = "";
            TargetListView.Items.Clear();
            //關閉更新即刪除
            updata_btn.Enabled = false;
            deldata_btn.Enabled = false;
        }


        //==========項目設定========== 結束





        //=============新增精靈========== 開始
        private void radWizard1_Help(object sender, EventArgs e)
        {
            MessageBox.Show("在這個情況下，我無法給你更多的幫助");
        }

        //下一步時
        private void radWizard1_Next(object sender, Telerik.WinControls.UI.WizardCancelEventArgs e)
        {
            //如果在第二夜，則檢察目標機器的硬碟
            if (radWizard1.SelectedPage == this.radWizard1.Pages[0])
            {
                //清除列表
                TargetDeviceIDView.Items.Clear();
                try
                {
                    //這邊進行wmi查詢目標硬碟代號
                    System.Management.ConnectionOptions Conn = new ConnectionOptions();
                    //設定用於WMI連接操作的用戶名
                    string user = NewTargetDomainName.Text + "\\" + NewTargetUsername.Text;
                    Conn.Username = user;
                    //設定用戶的密碼
                    Conn.Password = NewTargetPassword.Text;
                    //設定用於執行WMI操作的範圍
                    System.Management.ManagementScope Ms = new ManagementScope("\\\\" + NewTargetIP.Text + "\\root\\cimv2", Conn);
                    Ms.Connect();
                    ObjectQuery Query = new ObjectQuery("select DeviceID from Win32_LogicalDisk where DriveType=3");
                    //WQL語句，設定的WMI查詢內容和WMI的操作範圍，檢索WMI對象集合
                    ManagementObjectSearcher Searcher = new ManagementObjectSearcher(Ms, Query);
                    //異步調用WMI查詢
                    ManagementObjectCollection ReturnCollection = Searcher.Get();
                    foreach (ManagementObject Return in ReturnCollection)
                    {
                        TargetDeviceIDView.Items.Add(Return["DeviceID"]);
                    }
                }
                catch (Exception)
                {
                    //發生問題 警告並回到第一夜
                    radWizard1.SelectedPage = this.radWizard1.WelcomePage;
                }
            }
            //如果在第三夜 則寫入資料庫
            if (radWizard1.SelectedPage == this.radWizard1.Pages[1])
            {
                //如果列表是空的 且烈表中有打勾
                if (TargetDeviceIDView.Items.Count!=0 && TargetDeviceIDView.CheckedItems!=null)
                {
                    //取得勾選的deviceID
                    foreach (ListViewDataItem DeviceID in TargetDeviceIDView.CheckedItems)
                    {
                        //對DB座新增動作
                        String strSQL = "INSERT INTO [WindowsTarget](TargetName,TargetIP,TargetDomain,TargetUser,TargetPassword,TargetDeviceID) VALUES ('" + NewTargetName.Text + "','" + NewTargetIP.Text + "','" + NewTargetDomainName.Text + "','" + NewTargetUsername.Text + "','" + NewTargetPassword.Text + "','" + DeviceID.Text + "')";
                        System.Data.OleDb.OleDbConnection oleConn =
                        new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=FreeSpaceEyesDB.mdb");
                        //// 開啟資料庫連接。  
                        oleConn.Open();
                        //// OleDbCommand (String, OleDbConnection) : 使用查詢的文字和 OleDbConnection，初始化 OleDbCommand 類別的新執行個體。  
                        System.Data.OleDb.OleDbCommand oleCmd =
                        new System.Data.OleDb.OleDbCommand(strSQL, oleConn);
                        oleCmd.ExecuteNonQuery();
                        //// 關閉資料庫連接。  
                        oleConn.Close();
                    }

                }
                else
                {
                    MessageBox.Show("無法取得目標DeviceID，請回到上一步");
                }


            }
        }
        //結束時動作
        private void radWizard1_Finish(object sender, EventArgs e)
        {
            reset_all_new_detail();
            radWizard1.SelectedPage = this.radWizard1.WelcomePage;
        }
        //按下cancel時
        private void radWizard1_Cancel(object sender, EventArgs e)
        {
            reset_all_new_detail();
            radWizard1.SelectedPage = this.radWizard1.WelcomePage;
        }
        //重置所有項目含數
        private void reset_all_new_detail() 
        {
            NewTargetIP.Text = "";
            NewTargetDomainName.Text = "";
            NewTargetName.Text = "";
            NewTargetPassword.Text = "";
            NewTargetUsername.Text = "";
            //清空資料 and 清空被選取得資料
            foreach (ListViewDataItem item in TargetDeviceIDView.Items)
            {
                item.CheckState = Telerik.WinControls.Enumerations.ToggleState.Off;
            }
            TargetDeviceIDView.Items.Clear();
        }

        //轉換ip用
        private void TransIP_btn_Click(object sender, EventArgs e)
        {
            try
            {
                IPAddress[] ip = Dns.GetHostAddresses(NewTargetName.Text);
                NewTargetIP.Text = ip[0].ToString();
            }
            catch (Exception)
            {

            }
            
        }



        //========新增精靈=========== 結束
    }
}
