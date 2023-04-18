using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.NetworkInformation;
using System.Net;

namespace RunOutNG
{
    public partial class Form1 : Form
    {
        string DelFlag = "1";
        string Date = String.Empty;
        string DatePlus1 = String.Empty;

       // string connstr = ("Data Source=10.243.151.14;Initial Catalog=ASANMOD;Persist Security Info=True;User ID=********;Password=********");
      //OA 망에서 접속시
        string connstr = ("Data Source=10.21.205.152;Initial Catalog=ASANEE;Persist Security Info=True;User ID=sa;Password=********");

        public Form1()
        {
            InitializeComponent();
            

        }
        
        public void Form1_Load(object sender, EventArgs e)

        {
            button1.Enabled = false;
            button1.Location = new Point(this.Width / 2 - button1.Width / 2, button1.Location.Y);
        }


        public void LoadRunOutNG()
        {
            Ping ping = new Ping();

            PingOptions options = new PingOptions();

            options.DontFragment = true;

            string data = "aaaaaaaaaaaaaa";

            byte[] buffer = ASCIIEncoding.ASCII.GetBytes(data);

            int timeout = 120;

            PingReply reply = ping.Send(IPAddress.Parse("10.21.205.152"), timeout, buffer, options);



            if (reply.Status != IPStatus.Success)

            {

                MessageBox.Show("DB 접속 불가", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);


            }

            else
            {
                SqlConnection conn = new SqlConnection(connstr);

                conn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conn;
                //cmd.CommandText = "SELECT * FROM [RunOut].[dbo].[TResult] with(nolock) where InspectDate BETWEEN '" + Date + " 06:00:00' and '" + DatePlus1 + " 00:40:00' and Result in (0,2) order by InspectDate,LineCode";
                // OA 망에서 접속시
                cmd.CommandText = "SELECT * FROM [RemoteDB].[RunOut].[dbo].[TResult] with(nolock) where InspectDate BETWEEN '" + Date + " 06:00:00' and '" + DatePlus1 + " 00:40:00' and Result in (0,2) order by InspectDate,LineCode";
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);

                int rowcount = ds.Tables[0].Rows.Count;

                if (rowcount < 1)
                {
                    dataGridView1.Columns.Clear();
                    MessageBox.Show("런아웃 NG 이력이 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    button1.Enabled = false;
                }

                else

                {
                    dataGridView1.DataSource = ds.Tables[0];
                    dataGridView1.AutoResizeColumns();
                    dataGridView1.ReadOnly = true;
                    button1.Enabled = true;
                }

                conn.Close();
            }
            

        }

        private void button1_Click(object sender, EventArgs e)
        {
           
                string[] filenames = Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "런아웃NG이력_*");
                List<string> filelist = new List<string>();
           
            if (DelFlag == "1")
            {
                foreach (string todelete in filenames)

                {

                    FileInfo file = new FileInfo(todelete);
                    file.Delete();
                    string[] splited = todelete.Split('\\');
                    filelist.Add(splited[4]);

                }
            }

            string[] realfilename = filelist.ToArray();
            
            ExportExcel("런아웃NG이력_" + Date);

            if (filelist.Count > 0 & DelFlag == "1")

            {
                string toDisplay = string.Join(Environment.NewLine, realfilename);
                MessageBox.Show("엑셀 파일 저장 성공,\r삭제한 기존 파일 :\r"+toDisplay, "정보", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            else

            {
                MessageBox.Show("엑셀 파일 저장 성공.", "정보", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        private void ExportExcel(string filename)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            
            object misValue = System.Reflection.Missing.Value;

            Int16 i, j;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            string[] colNames = new string[dataGridView1.Columns.Count];
            int col = 0;

            foreach (DataGridViewColumn dc in dataGridView1.Columns)
                colNames[col++] = dc.HeaderText;

            for (i = 0; i < dataGridView1.RowCount; i++)
            {
                for (j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    xlWorkSheet.Cells[i + 2, j + 1] = dataGridView1[j, i].Value.ToString();
                }
            }
                       
            

            char lastColumn = (char)(65 + dataGridView1.Columns.Count - 1);

            xlWorkSheet.get_Range("A1", lastColumn + "1").Value2 = colNames;
            xlWorkSheet.get_Range("A1", lastColumn + "1").Font.Bold = true;
            xlWorkSheet.get_Range("A1", lastColumn + "1").VerticalAlignment
                        = Excel.XlVAlign.xlVAlignCenter;

            xlWorkSheet.Columns.AutoFit();
            xlWorkBook.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+"\\"+filename +".xls",Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            
        }

        
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
          
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            
            Date = dateTimePicker1.Value.Date.ToShortDateString();
            DatePlus1 = dateTimePicker1.Value.AddDays(1).ToShortDateString();

            LoadRunOutNG();
            GC.Collect();
         
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (DelFlag == "1")
            {
                DelFlag = "0";
            }
            else
            {
                DelFlag = "1";
            }
        }
    }
}
