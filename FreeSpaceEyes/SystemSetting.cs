using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls;
using System.Data.OleDb;

namespace FreeSpaceEyes
{
    public partial class SystemSetting : Telerik.WinControls.UI.RadForm
    {
        public SystemSetting()
        {
            InitializeComponent();
        }

        private void SystemSetting_Load(object sender, EventArgs e)
        {
            //===開始讀取SystemSetting===
            String strSQL = "SELECT * FROM [SystemSetting] WHERE id=1";
            System.Data.OleDb.OleDbConnection oleConn =
                 new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=FreeSpaceEyesDB.mdb");
            oleConn.Open();

            System.Data.OleDb.OleDbCommand oleCmd = new System.Data.OleDb.OleDbCommand(strSQL, oleConn);
            OleDbDataReader thisReader = oleCmd.ExecuteReader();//建立OleDbDataReader用來存放從Access裡面讀取出來的
            while (thisReader.Read())
            {
                Timer_num.Value = int.Parse(thisReader["Timer"].ToString());
                Thread_num.Value = int.Parse(thisReader["MaxThread"].ToString());
                ThreadIO_num.Value = int.Parse(thisReader["MaxIOThread"].ToString());
                LimitNonPage_Num.Value = int.Parse(thisReader["LimitNonPage"].ToString());
                LimitTimeRange_Num.Value = int.Parse(thisReader["LimitTimeRange"].ToString());
            }
            oleConn.Close();
        }

        //確定更新按鈕
        private void Update_btn_Click(object sender, EventArgs e)
        {
            // === 對 Access 資料庫下SQL語法 ===  
            //// Transact-SQL 陳述式  
            String strSQL = "UPDATE [SystemSetting] SET Timer = " + Timer_num.Value + " ,MaxThread = " + Thread_num.Value + " ,MaxIOThread = " + ThreadIO_num.Value + " ,LimitNonPage = " + LimitNonPage_Num.Value + " ,LimitTimeRange = " + LimitTimeRange_Num.Value + " WHERE id=1";
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
            this.Close();
        }
    }
}
