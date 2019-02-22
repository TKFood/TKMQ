﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;
using FastReport;
using FastReport.Data;
using System.Net.Mail;//<-基本上發mail就用這個class
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace TKMQ
{
    public partial class FrmMQMAIL : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapterMAIL = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAIL = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet dsMAIL = new DataSet();

        string DirectoryNAME = @"C:\MQTEMP\" + DateTime.Now.ToString("yyyyMMdd")+@"\";
        string pathFile = @"C:\MQTEMP\" + DateTime.Now.ToString("yyyyMMdd") + @"\" + "MQ" + DateTime.Now.ToString("yyyyMMdd");
        FileInfo info;
        string[] tempFile;
        string tFileName = "";

        public FrmMQMAIL()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SENDMAIL(DataSet SEND)
        {            
            string MySMTPCONFIG = ConfigurationManager.AppSettings["MySMTP"];
            string NAME = ConfigurationManager.AppSettings["NAME"];
            string PW = ConfigurationManager.AppSettings["PW"];

            System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
            MyMail.From = new System.Net.Mail.MailAddress("tk290@tkfood.com.tw");
           
            //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
            MyMail.Subject = "每日訂單-製令追踨表"+DateTime.Now.ToString("yyyy/MM/dd");
            MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
            MyMail.IsBodyHtml = true; //是否使用html格式

            System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
            MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);
                        

            if (Directory.Exists(DirectoryNAME))
            {
                tempFile = Directory.GetFiles(DirectoryNAME);//取得資料夾下所有檔案

                foreach (string item in tempFile)
                {
                    info = new FileInfo(item);
                    tFileName = info.Name.ToString().Trim();//取得檔名
                    Attachment attch = new Attachment(DirectoryNAME+tFileName);
                    MyMail.Attachments.Add(attch);

                }

            }

            
            try
            {
                int i = 0;
                foreach (DataTable table in SEND.Tables)
                {                    
                    //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                    MyMail.To.Add(table.Rows[i].ItemArray[1].ToString()); //設定收件者Email
                    MySMTP.Send(MyMail);

                    i++;
                }

               
              
                MyMail.Dispose(); //釋放資源

                MessageBox.Show("OK");
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }
    
        public void SETFILE()
        {
            if (Directory.Exists(DirectoryNAME))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(DirectoryNAME);
            }

            // 設定儲存檔名，不用設定副檔名，系統自動判斷 excel 版本，產生 .xls 或 .xlsx 副檔名 
            Excel.Application excelApp;
            Excel._Workbook wBook;
            Excel._Worksheet wSheet;
            Excel.Range wRange;

            // 開啟一個新的應用程式
            excelApp = new Excel.Application();
            // 讓Excel文件可見
            //excelApp.Visible = true;
            // 停用警告訊息
            excelApp.DisplayAlerts = false;
            // 加入新的活頁簿
            excelApp.Workbooks.Add(Type.Missing);
            // 引用第一個活頁簿
            wBook = excelApp.Workbooks[1];
            // 設定活頁簿焦點
            wBook.Activate();

            wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                       

            //關閉Excel
            excelApp.Quit();

            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wSheet = null;
            wRange = null;
            excelApp = null;
            GC.Collect();

            Console.Read();

            //SEARCH()
            SEARCH();
        }

        public void SEARCH()
        {
            DateTime SEARCHDATE = DateTime.Now;
            SEARCHDATE = SEARCHDATE.AddMonths(-1);


            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TC053 AS '客戶',TD013 AS '預計交貨日',TD004 AS '訂單品號',TD005 AS '訂單品名',TD006 AS '規格',TD008 AS '訂單量',TD009 AS '出貨量',TD024 AS '贈品量',TD025 AS '贈品已交量',(TD008-TD009+TD024-TD025) AS '總未出貨量',TD010 AS '品號單位',TD001 AS '訂單單別',TD002 AS '訂單單號',TD003 AS '訂單序號',TD016 AS '訂單狀態',MOCTA.TA001 AS '批次轉製令單別',MOCTA.TA002 AS '批次轉製令單號',MOCTA.TA009 AS '製令預計開工日',MOCTA.TA012 AS '製令實際開工日',MOCTA.TA010 AS '製令預計完工日' ,MOCTA.TA014 AS '製令實際完工日',MOCTA.TA006 AS '生產品號',MOCTA.TA034 AS '生產品名',MOCTA.TA007 AS '生產單位',MOCTA.TA015 AS '製令預計產量',MOCTA.TA017 AS '實際入庫數量'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MOCTA.TA011='Y' THEN '已完工' ELSE CASE WHEN MOCTA.TA011='y' THEN '指定完工' ELSE  CASE WHEN MOCTA.TA011='1' THEN '未生產' ELSE CASE WHEN MOCTA.TA011='2' THEN '已發料' ELSE CASE WHEN MOCTA.TA011='3' THEN '生產中' ELSE '' END END END END END)AS '生產進度'");
                sbSql.AppendFormat(@"  ,(CASE WHEN CONVERT(datetime,MOCTA.TA009)<CONVERT(datetime,MOCTA.TA012) THEN '是' ELSE ''  END ) AS '製令開工異常警示'");
                sbSql.AppendFormat(@"  ,(CASE WHEN CONVERT(datetime,MOCTA.TA010)<CONVERT(datetime,MOCTA.TA014) THEN '是' ELSE ''  END ) AS '製令完工異常警示'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MOCTA.TA017<MOCTA.TA015 THEN '是' ELSE ''  END) AS '產量不足'");
                sbSql.AppendFormat(@"  ,LRPTA.TA001 AS '批次計畫單號'");
                sbSql.AppendFormat(@"  ,(CASE WHEN ISNULL(MOCTA.TA033,'')<>''  THEN '是' ELSE ''  END )  AS '製令發放'");
                sbSql.AppendFormat(@"  ,(CASE WHEN CONVERT(datetime,TD013)<=CONVERT(datetime,MOCTA.TA009) THEN '是' ELSE ''  END )  AS '訂單是否延遲生產'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.MOCTA ON MOCTA.TA026=TD001 AND MOCTA.TA027=TD002 AND MOCTA.TA028=TD003");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.LRPTA ON LRPTA.TA023=TD001 AND LRPTA.TA024=TD002 AND LRPTA.TA025=TD003");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD013>='{0}' ", SEARCHDATE.ToString("yyyyMM")+"01");
                sbSql.AppendFormat(@"  AND TD004 LIKE '4%'");
                sbSql.AppendFormat(@"  AND (TD008-TD009+TD024-TD025)>0");
                sbSql.AppendFormat(@"  AND TD021='Y' ");
                sbSql.AppendFormat(@"  AND TD016='N'");
                sbSql.AppendFormat(@"  AND TC001 IN ('A221', 'A222','A223','A227','A228')");
                sbSql.AppendFormat(@"  ORDER BY TC053,TD013,TD004");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                   
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(ds1, pathFile);
                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void ExportDataSetToExcel(DataSet ds,string TopathFile)
        {
            //Creae an Excel application instance
            Excel.Application excelApp = new Excel.Application();

            //Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(TopathFile);

            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                        // Set the range to fill.
                       
                    }
                }

                //設定為按照內容自動調整欄寬
                excelWorkSheet.Columns.AutoFit();
            }

            

            excelWorkBook.Save();
            excelWorkBook.Close();
            excelApp.Quit();

        }

        public void SERACHMAIL()
        {          
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

              
                sbSql.AppendFormat(@" SELECT [SENDTO],[MAIL] ");
                sbSql.AppendFormat(@" FROM [TKMQ].[dbo].[MQSENDMAIL] ");
                sbSql.AppendFormat(@"  WHERE [SENDTO]='COP'");
                sbSql.AppendFormat(@"  ");

                adapterMAIL = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMAIL = new SqlCommandBuilder(adapterMAIL);
                sqlConn.Open();
                dsMAIL.Clear();
                adapterMAIL.Fill(dsMAIL, "TEMPdsMAIL");
                sqlConn.Close();


                if (dsMAIL.Tables["TEMPdsMAIL"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsMAIL.Tables["TEMPdsMAIL"].Rows.Count >= 1)
                    {
                        
                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }
  
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFILE();
            SERACHMAIL();
            SENDMAIL(dsMAIL);
        }
        private void button2_Click(object sender, EventArgs e)
        {
           
            MessageBox.Show("OK");
        }
        #endregion


    }
}
