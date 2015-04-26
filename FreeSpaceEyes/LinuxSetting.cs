using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using System.Net;
using System.Data.OleDb;
using System.Threading;
using Telerik.WinControls.UI;

namespace FreeSpaceEyes
{
    public partial class LinuxSetting : Telerik.WinControls.UI.RadForm
    {
        public LinuxSetting()
        {
            InitializeComponent();
        }
        private void LinuxSetting_Load(object sender, EventArgs e)
        {
            TargetListView.Items.Clear();
            Reload_btn.Enabled = false;
            radWaitingBar1.StartWaiting();
            //啟動背景讀取 送出使用者要找的文字
            LoadAccess.RunWorkerAsync(find_text.Text);
        }
        //=======項目設定======== 開始

        private void Reload_btn_Click(object sender, EventArgs e)
        {
            TargetListView.Items.Clear();
            Reload_btn.Enabled = false;
            radWaitingBar1.StartWaiting();
            //啟動背景讀取 送出使用者要找的文字
            LoadAccess.RunWorkerAsync(find_text.Text);
        }
        private void LoadAccess_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                string find_text = (string)e.Argument;
                String strSQL = "SELECT id,TargetName FROM [LinuxTarget] WHERE TargetName LIKE '%" + find_text + "%'";
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
            //寫入targetlistview 0=id,1=targetname
            string[] Temp = (string[])item;
            TargetListView.Items.Add(Temp[0], Temp[1]);
        }
        //背景執行結束後停止跑條
        private void LoadAccess_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Reload_btn.Enabled = true;
            radWaitingBar1.StopWaiting();
        }

        private void TargetListView_ItemMouseClick(object sender, Telerik.WinControls.UI.ListViewItemEventArgs e)
        {
            ListViewDataItem item = TargetListView.SelectedItem;
            //將資料拉出來到ui上
            // === 對 Access 資料庫下SQL語法 ===  
            //// Transact-SQL 陳述式  
            String strSQL = "SELECT id,TargetName,TargetIP,TargetAlert,TargetRes,TargetSpaceLog,TargetWarning FROM [LinuxTarget] WHERE id =" + int.Parse(item[0].ToString());
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
                ShowTargetAlert.Text = thisReader["TargetAlert"].ToString();
                ShowTargetRes.Text = thisReader["TargetRes"].ToString();
                ShowTargetWarning.Text = thisReader["TargetWarning"].ToString();
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
            String strSQL = "UPDATE [LinuxTarget] SET TargetName = '" + ShowTargetName.Text + "' ,TargetIP = '" + ShowTargetIP.Text + "' ,TargetAlert = '" + ShowTargetAlert.Text + "' , TargetRes='" + ShowTargetRes.Text + "' , TargetWarning='" + ShowTargetWarning.Text + "' WHERE id=" + int.Parse(ShowTargetID.Text);
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

        private void deldata_btn_Click(object sender, EventArgs e)
        {
            // === 對 Access 資料庫下SQL語法 ===  
            //// Transact-SQL 陳述式  
            String strSQL = "DELETE FROM [LinuxTarget] WHERE id =" + int.Parse(ShowTargetID.Text);
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
            ShowTargetAlert.Text = "";
            ShowTargetRes.Text = "";
            ShowTargetWarning.Text = "";
            TargetListView.Items.Clear();
            //關閉更新即刪除
            updata_btn.Enabled = false;
            deldata_btn.Enabled = false;
        }



        //=======項目設定======== 結束

        //=============新增精靈========== 開始
        private void radWizard1_Help(object sender, EventArgs e)
        {
            MessageBox.Show("在這個情況下，我無法給你更多的幫助");
        }

        //下一步時
        private void radWizard1_Next(object sender, Telerik.WinControls.UI.WizardCancelEventArgs e)
        {
            //如果列表是空的 且烈表中有打勾
            if (NewTargetName.Text != "" && NewTargetIP.Text != "")
            {
                //對DB座新增動作
                String strSQL = "INSERT INTO [LinuxTarget](TargetName,TargetIP) VALUES ('" + NewTargetName.Text + "','" + NewTargetIP.Text + "')";
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
            else
            {
                MessageBox.Show("欄位不得為空");
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
            NewTargetName.Text = "";

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
