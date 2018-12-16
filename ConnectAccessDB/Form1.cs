using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ConnectAccessDB
{
    public partial class Form1 : Form
    {
        private static OleDbConnection _dbconn;
        public Form1()
        {
            InitializeComponent();
            string Con = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source=C:\AccessDB\Database1.mdb";//第二个参数为文件的路径  
            _dbconn = new OleDbConnection(Con);
             
            _dbconn.Open();

        }

        private void ConnectDb()
        {
            //string Con = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source=C:\AccessDB\Database1.mdb";//第二个参数为文件的路径  
            //OleDbConnection dbconn = new OleDbConnection(Con);
            //dbconn.Open();//建立连接
            //return dbconn;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            OleDbDataAdapter inst = new OleDbDataAdapter("SELECT * FROM T_User", _dbconn);//选择全部内容
            DataSet ds = new DataSet();//临时存储
            inst.Fill(ds);//用inst填充ds
            var a = ds.Tables[0];

        }

        private void button5_Click(object sender, EventArgs e)
        {
            _dbconn.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {

            //string Insert = @"INSERT INTO T_User(Name,Sex,Age) values('{0}','{1}','{2}')";
            string insert = String.Format("INSERT INTO T_User(Name,Sex,Age) values('{0}','{1}','{2}')","Bob","male",25);
            OleDbCommand myCommand = new OleDbCommand(insert, _dbconn);//执行命令
            myCommand.ExecuteNonQuery();
        }

        private void button4_Click(object sender, EventArgs e)
        {          
            string delete = "DELETE FROM T_User WHERE Name = '"+"chao"+"'";
            OleDbCommand myCommand = new OleDbCommand(delete, _dbconn);//执行命令
            myCommand.ExecuteNonQuery();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string update = String.Format("UPDATE T_User SET Name = '" + "Sam" + "' WHERE Name ='"+"Bob"+"' ");
            OleDbCommand myCommand = new OleDbCommand(update, _dbconn);//执行命令
            myCommand.ExecuteNonQuery();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string update = String.Format("DELETE FROM T_User");
            OleDbCommand myCommand = new OleDbCommand(update, _dbconn);//执行命令
            myCommand.ExecuteNonQuery();
        }

        private void test()
        {
            //OpenFileDialog dlg = new OpenFileDialog();
            //dlg.Filter = "Excel文件(*.xlsx)|*.xlsx";
            //if (dlg.ShowDialog() == DialogResult.OK)
            //{
            //    string filePath = dlg.FileName;
            //    string a = filePath;
            //}
            OleDbConnectionStringBuilder connectStringBuilder = new OleDbConnectionStringBuilder();
            connectStringBuilder.DataSource = @"C:\AccessDB\test.xlsx";
            connectStringBuilder.Provider = "Microsoft.ACE.OLEDB.16.0";
            connectStringBuilder.Add("Extended Properties", "Excel 8.0");
            using (OleDbConnection cn = new OleDbConnection(connectStringBuilder.ConnectionString))
            {
                DataSet ds = new DataSet();
                string sql = "Select * from [Sheet1$]";
                OleDbCommand cmdLiming = new OleDbCommand(sql, cn);
                cn.Open();
                using (OleDbDataReader drLiming = cmdLiming.ExecuteReader())
                {
                    ds.Load(drLiming, LoadOption.OverwriteChanges, new string[] { "Sheet1" });
                    DataTable dt = ds.Tables["Sheet1"];
                    if (dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            string MySql = "insert into T_User (Name,Sex,Age) values('" + dt.Rows[i]["Name"].ToString() + "','"+ dt.Rows[i]["Sex"].ToString() + "','" 
                                           + dt.Rows[i]["Age"].ToString() + "')";
                            SQLExecute(MySql);
                        }
                    }
                }
                cn.Close();
            }

        }

        public void insertIntoExcel()
        {
            OleDbConnectionStringBuilder connectStringBuilder = new OleDbConnectionStringBuilder();
            connectStringBuilder.DataSource = @"C:\AccessDB\test1.xlsx";
            connectStringBuilder.Provider = "Microsoft.ACE.OLEDB.16.0";
            connectStringBuilder.Add("Extended Properties", "Excel 8.0");
            using (OleDbConnection cn = new OleDbConnection(connectStringBuilder.ConnectionString))
            {
                string sql = "Insert into [Sheet1$] (Name,Sex,Age) values ('"+"chao"+"','"+"male"+"','"+18+"')";
                OleDbCommand cmdLiming = new OleDbCommand(sql, cn);
                cn.Open();
                cmdLiming.ExecuteNonQuery();
                cn.Close();
            }
        }

        public bool SQLExecute(string sql)
        {
            try
            {
                OleDbCommand comm = new OleDbCommand();
                comm.Connection = _dbconn;
                comm.CommandText = sql;
                comm.ExecuteNonQuery();
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            test();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            insertIntoExcel();
        }
    }
}
