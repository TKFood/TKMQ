using System;
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
using System.Diagnostics;
using System.Threading;

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
        SqlDataAdapter adapterCOPTE = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderCOPTE = new SqlCommandBuilder();
        SqlDataAdapter adapterPURTA = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderPURTA = new SqlCommandBuilder();
        SqlDataAdapter adapterMOCTA = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMOCTA = new SqlCommandBuilder();

        SqlDataAdapter adapterMAILCOPTE = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILCOPTE = new SqlCommandBuilder();
        SqlDataAdapter adapterMAILPURTA = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILPURTA = new SqlCommandBuilder();
        SqlDataAdapter adapterMAILMOCTA = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILMOCTA = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;
        DataSet ds1 = new DataSet();
        DataSet dsMAIL = new DataSet();
        DataSet dsCOPTE = new DataSet();
        DataSet dsMAILCOPTE = new DataSet();
        DataSet dsPURTA = new DataSet();
        DataSet dsMAILPURTA = new DataSet();
        DataSet dsMOCTA = new DataSet();
        DataSet dsMAILMOCTA = new DataSet();

        string DATES =null;
        string DirectoryNAME=null;
        string pathFile = null;
        string pathFileCOPTE = null;
        string pathFilePURTA = null;
        string pathFileMOCTA = null;


        FileInfo info;
        string[] tempFile;
        string tFileName = "";

        public FrmMQMAIL()
        {
            InitializeComponent();

            timer1.Enabled = true;
            timer1.Interval = 1000 *60;
            //timer1.Interval = 1000 ;
            timer1.Start();

            CLEAREXCEL();

            SETPATH();
        }

        #region FUNCTION
        public void SETPATH()
        {

            DATES = DateTime.Now.ToString("yyyyMMdd");

            DirectoryNAME = @"C:\MQTEMP\" + DATES.ToString() + @"\";
            pathFile = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日訂單-製令追踨表" + DATES.ToString();
            pathFileCOPTE = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日訂單變更表" + DATES.ToString();
            pathFilePURTA = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日製令-請購表" + DATES.ToString();
            pathFileMOCTA = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日製令-訂單表" + DATES.ToString();

        }

        public void CLEAREXCEL()
        {
            System.Diagnostics.Process[] p = System.Diagnostics.Process.GetProcesses();
            for (int i = 0; i < p.Length; i++)
            {
                if (p[i].ToString().IndexOf("EXCEL") > 0)
                    p[i].Kill();
            }
        }
        public void SENDMAIL(StringBuilder Subject, StringBuilder Body, DataSet SEND,string Attachments)
        {            
            string MySMTPCONFIG = ConfigurationManager.AppSettings["MySMTP"];
            string NAME = ConfigurationManager.AppSettings["NAME"];
            string PW = ConfigurationManager.AppSettings["PW"];

            System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
            MyMail.From = new System.Net.Mail.MailAddress("tk290@tkfood.com.tw");

            //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
            //MyMail.Subject = "每日訂單-製令追踨表"+DateTime.Now.ToString("yyyy/MM/dd");
            MyMail.Subject = Subject.ToString();
            //MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
            MyMail.Body = Body.ToString();
            //MyMail.IsBodyHtml = true; //是否使用html格式

            System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
            MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);

            Attachment attch = new Attachment( Attachments+".xlsx");
            MyMail.Attachments.Add(attch);

            //if (Directory.Exists(DirectoryNAME))
            //{
            //    tempFile = Directory.GetFiles(DirectoryNAME);//取得資料夾下所有檔案

            //    foreach (string item in tempFile)
            //    {
            //        info = new FileInfo(item);
            //        tFileName = info.Name.ToString().Trim();//取得檔名
            //        Attachment attch = new Attachment(DirectoryNAME+tFileName);
            //        MyMail.Attachments.Add(attch);

            //    }

            //}

            
            try
            {
                foreach (DataRow od in SEND.Tables[0].Rows)
                {                    
                   
                    MyMail.To.Add(od["MAIL"].ToString()); //設定收件者Email，多筆mail
                }

                //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email

                MySMTP.Send(MyMail);

                MyMail.Dispose(); //釋放資源

               
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
                //資料夾存在，pathFile
                if (File.Exists(pathFile + ".xlsx"))
                {
                    File.Delete(pathFile + ".xlsx");
                }
                  
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

            if(!File.Exists(pathFile+".xlsx"))
            {
                wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            
                       

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


            SEARCH();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()
                
            //}
                
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

                sbSql.AppendFormat(@"  SELECT TC053 AS '客戶',TD013 AS '預計交貨日',TD004 AS '訂單品號',TD005 AS '訂單品名',TD006 AS '規格',TD008 AS '訂單量',TD009 AS '出貨量',TD024 AS '贈品量',TD025 AS '贈品已交量',(TD008-TD009+TD024-TD025) AS '總未出貨量',TD010 AS '品號單位',TD001 AS '訂單單別',TD002 AS '訂單單號',TD003 AS '訂單序號',TD016 AS '訂單狀態',MOCTA.TA001 AS '批次轉製令單別',MOCTA.TA002 AS '批次轉製令單號',MOCTA.TA009 AS '製令預計開工日',MOCTA.TA012 AS '製令實際開工日',MOCTA.TA010 AS '製令預計完工日' ,MOCTA.TA014 AS '製令實際完工日',MOCTA.TA006 AS '生產品號',MOCTA.TA034 AS '生產品名',MOCTA.TA007 AS '生產單位',MOCTA.TA015 AS '製令預計產量',MOCTA.TA017 AS '實際入庫數量',COMMENT AS '備註'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MOCTA.TA011='Y' THEN '已完工' ELSE CASE WHEN MOCTA.TA011='y' THEN '指定完工' ELSE  CASE WHEN MOCTA.TA011='1' THEN '未生產' ELSE CASE WHEN MOCTA.TA011='2' THEN '已發料' ELSE CASE WHEN MOCTA.TA011='3' THEN '生產中' ELSE '' END END END END END)AS '生產進度'");
                sbSql.AppendFormat(@"  ,(CASE WHEN CONVERT(datetime,MOCTA.TA009)<CONVERT(datetime,MOCTA.TA012) THEN '是' ELSE ''  END ) AS '製令開工異常警示'");
                sbSql.AppendFormat(@"  ,(CASE WHEN CONVERT(datetime,MOCTA.TA010)<CONVERT(datetime,MOCTA.TA014) THEN '是' ELSE ''  END ) AS '製令完工異常警示'");
                sbSql.AppendFormat(@"  ,(CASE WHEN MOCTA.TA017<MOCTA.TA015 THEN '是' ELSE ''  END) AS '產量不足'");
                sbSql.AppendFormat(@"  ,LRPTA.TA001 AS '批次計畫單號'");
                sbSql.AppendFormat(@"  ,(CASE WHEN ISNULL(MOCTA.TA033,'')<>''  THEN '是' ELSE ''  END )  AS '製令發放'");
                sbSql.AppendFormat(@"  ,(CASE WHEN CONVERT(datetime,TD013)<=CONVERT(datetime,MOCTA.TA009) THEN '是' ELSE ''  END )  AS '訂單是否延遲生產'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.MOCTA ON MOCTA.TA026=TD001 AND MOCTA.TA027=TD002 AND MOCTA.TA028=TD003 AND TD004=MOCTA.TA006");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.LRPTA ON LRPTA.TA023=TD001 AND LRPTA.TA024=TD002 AND LRPTA.TA025=TD003");
                sbSql.AppendFormat(@"  LEFT JOIN [TKMOC].dbo.MOCCOPCHECK ON COPTA001=TD001 AND COPTA002=TD002 AND COPTA003=TD003 ");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD013>='{0}' ", SEARCHDATE.ToString("yyyyMM")+"01");
                sbSql.AppendFormat(@"  AND TD004 LIKE '4%'");
                sbSql.AppendFormat(@"  AND (TD008-TD009+TD024-TD025)>0");
                sbSql.AppendFormat(@"  AND TD021='Y' ");
                sbSql.AppendFormat(@"  AND TD016='N'");
                sbSql.AppendFormat(@"  AND TC001 IN ('A221', 'A222','A223','A227','A228')");
                //sbSql.AppendFormat(@"  AND TD002 IN ('20190318001','20190218009')");
                sbSql.AppendFormat(@"  ORDER BY TC053,TD013,TD004");
           
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
            Excel.Range wRange;
            Excel.Range wRangepathFile;
            Excel.Range wRangepathFilePURTA;

            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                    //畫框線
                    wRange = excelWorkSheet.Cells[1, i];
                    wRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    wRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    wRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();

                        wRange = excelWorkSheet.Cells[j + 2, k + 1];

                        //畫框線
                        wRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                        //pathFilePURTA檢查需求差異量是否為負，為負就紅字
                        //string tt = table.Rows[j].ItemArray[k].ToString();

                        if (TopathFile.Equals(pathFile.ToString()) && k == 15 && string.IsNullOrEmpty(table.Rows[j].ItemArray[k].ToString()))
                        {
                            string STARTCELL = "A"+(j+2).ToString();
                            string ENDCELL = "AH" +(j+2).ToString();
                            Excel.Range newRng = excelApp.get_Range(STARTCELL, ENDCELL);                           
                            newRng.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Red);
                 
                        }

                        //pathFilePURTA檢查需求差異量是否為負，為負就紅字
                        //string tt = table.Rows[j].ItemArray[k].ToString();

                        if (TopathFile.Equals(pathFilePURTA.ToString()) && k==5 && Convert.ToDecimal(table.Rows[j].ItemArray[k].ToString())<0)
                        {
                            wRange.Select();
                            wRange.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                       
                        //wRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.DimGray);
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

        public void SETFILECOPTE()
        {
            if (Directory.Exists(DirectoryNAME))
            {
                //資料夾存在，pathFileCOPTE
                if (File.Exists(pathFileCOPTE + ".xlsx"))
                {
                    File.Delete(pathFileCOPTE + ".xlsx");
                }
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


            if (!File.Exists(pathFileCOPTE + ".xlsx"))
            {
                wBook.SaveAs(pathFileCOPTE, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
                


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


            SEARCHCOPTE();

            //if (!File.Exists(pathFileCOPTE + ".xlsx"))
            //{
            //    //SEARCH()
                
            //}
                
        }

        public void SEARCHCOPTE()
        {
            DateTime SEARCHDATE = DateTime.Now;
            SEARCHDATE = SEARCHDATE.AddMonths(-1);


            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TE006 AS '變更原因',TE001 AS '訂單',TE002 AS '訂單號',TE003 AS '訂單序號',TF005 AS '品號',TF006 AS '品名',TF007 AS '規格',TF009 AS '數量',TF020 AS '新贈品量',TF010 AS '單位',TF015 AS '新預交日'");
                sbSql.AppendFormat(@"  FROM [TKMQ].[dbo].[TRIGGERRECORD],[TK].dbo.COPTE");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.COPTF ON TE001=TF001 AND TE002=TF002 AND TE003=TF003");
                sbSql.AppendFormat(@"  WHERE TE001=IDM AND TE002=IDSUB AND TE003=IDNO");      
                sbSql.AppendFormat(@"  AND MAILYN='N'");
                sbSql.AppendFormat(@"  ORDER BY TE006,TE001,TE002,TF005");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapterCOPTE = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderCOPTE = new SqlCommandBuilder(adapterCOPTE);

                sqlConn.Open();
                dsCOPTE.Clear();
                adapterCOPTE.Fill(dsCOPTE, "TEMPdsCOPTE");
                sqlConn.Close();


                if (dsCOPTE.Tables["TEMPdsCOPTE"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsCOPTE.Tables["TEMPdsCOPTE"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(dsCOPTE, pathFileCOPTE);

                        UPDATETRIGGERRECORDMAILYN();
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

        public void UPDATETRIGGERRECORDMAILYN()
        {
            try
            {
                if(dsCOPTE.Tables["TEMPdsCOPTE"].Rows.Count >= 1)
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    foreach (DataRow dr in dsCOPTE.Tables["TEMPdsCOPTE"].Rows)
                    {
                        var TE001 = dr["訂單"].ToString();
                        var TE002 = dr["訂單號"].ToString();
                        var TE003 = dr["訂單序號"].ToString();

                        sbSql.AppendFormat(" UPDATE [TKMQ].[dbo].[TRIGGERRECORD]");
                        sbSql.AppendFormat(" SET [MAILYN]='Y'");
                        sbSql.AppendFormat(" WHERE [IDM]='{0}' AND [IDSUB]='{1}' AND [IDNO]='{2}'", TE001.ToString(), TE002.ToString(), TE003.ToString());
                        sbSql.AppendFormat(" ");

                    }

                    sbSql.AppendFormat(" ");


                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消                        
                    }
                    else
                    {
                        tran.Commit();      //執行交易  
                        
                    }
                }
                

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void SERACHMAILCOPTE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT [SENDTO],[MAIL] ");
                sbSql.AppendFormat(@"  FROM [TKMQ].[dbo].[MQSENDMAIL] ");                
                sbSql.AppendFormat(@"  WHERE [SENDTO]='COP'  ");
                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  ");

                adapterMAILCOPTE = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMAILCOPTE = new SqlCommandBuilder(adapterMAILCOPTE);
                sqlConn.Open();
                dsMAILCOPTE.Clear();
                adapterMAILCOPTE.Fill(dsMAILCOPTE, "TEMPdsMAILCOPTE");
                sqlConn.Close();


                if (dsMAILCOPTE.Tables["TEMPdsMAILCOPTE"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsMAILCOPTE.Tables["TEMPdsMAILCOPTE"].Rows.Count >= 1)
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

        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToString();
            HRAUTORUN();
        }

        public void HRAUTORUN()
        {
            SETPATH();


            string RUNTIME = DateTime.Now.ToString("HH:mm");
            string hhmm = "07:11";

            if (RUNTIME.Equals(hhmm))
            {
                StringBuilder SUBJEST = new StringBuilder();
                StringBuilder BODY = new StringBuilder();


                SETFILE();
                CLEAREXCEL();
                Thread.Sleep(10000);

                SETFILECOPTE();
                CLEAREXCEL();
                Thread.Sleep(10000);

                SETFILEPURTA();
                CLEAREXCEL();
                Thread.Sleep(10000);

                SETFILEMOCTA();
                CLEAREXCEL();


                SERACHMAILCOPTE();
                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"每日訂單變更追踨表" + DateTime.Now.ToString("yyyy/MM/dd"));
                BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日訂單變更表，請查收" + Environment.NewLine + "請製造生管修改相對的製令");
                SENDMAIL(SUBJEST, BODY, dsMAILCOPTE, pathFileCOPTE);


                Thread.Sleep(10000);

                SERACHMAILPURTA();
                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"每日製令-請購表" + DateTime.Now.ToString("yyyy/MM/dd"));
                BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每每日製令-請購表，請查收" + Environment.NewLine + " ");
                SENDMAIL(SUBJEST, BODY, dsMAILPURTA, pathFilePURTA);


                Thread.Sleep(10000);

                SERACHMAIL();
                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"每日訂單-製令追踨表" + DateTime.Now.ToString("yyyy/MM/dd"));
                BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日訂單-製令追踨表，請查收" + Environment.NewLine + "若訂單沒有相對的製令則需通知製造生管開立");
                SENDMAIL(SUBJEST, BODY, dsMAIL, pathFile);

                Thread.Sleep(10000);

                SERACHMAILMOCTA();
                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"每日製令-訂單表" + DateTime.Now.ToString("yyyy/MM/dd"));
                BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每每日製令-訂單表，請查收" + Environment.NewLine + " ");
                SENDMAIL(SUBJEST, BODY, dsMAILMOCTA, pathFileMOCTA);




                //MessageBox.Show("OK");
            }
        }

        public void SETFILEPURTA()
        {
            if (Directory.Exists(DirectoryNAME))
            {
                //資料夾存在，pathFileCOPTE
                if (File.Exists(pathFilePURTA + ".xlsx"))
                {
                    File.Delete(pathFilePURTA + ".xlsx");
                }
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


            if (!File.Exists(pathFilePURTA + ".xlsx"))
            {
                wBook.SaveAs(pathFilePURTA, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }



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

            SEARCHVPURTDINVMD();
            //SEARCHPURTA();

            //if (!File.Exists(pathFileCOPTE + ".xlsx"))
            //{
            //    //SEARCH()

            //}

        }

        public void SEARCHVPURTDINVMD()
        {
            DateTime SEARCHDATE = DateTime.Now;


            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                sbSql.AppendFormat(@"  SELECT 品號,品名,需求量,單位,現有庫存,需求差異量,總採購量,最快採購日 ");
                sbSql.AppendFormat(@"  FROM (");
                sbSql.AppendFormat(@"  SELECT TB003 AS '品號',MB002 AS '品名' ,SUM(TB004-TB005) AS '需求量',TB007 AS '單位'");
                sbSql.AppendFormat(@"  ,(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA001=TB003 AND LA009=TB009) AS '現有庫存'");
                sbSql.AppendFormat(@"  ,(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA001=TB003 AND LA009=TB009)-SUM(TB004-TB005) AS '需求差異量'");
                sbSql.AppendFormat(@"  ,(SELECT ISNULL(CONVERT(DECIMAL(16,2),SUM(NUM)),0) FROM [TK].dbo.VPURTDINVMD WHERE  TD004=TB003 AND TD007=TD007 AND TD012>='{0}') AS '總採購量'", SEARCHDATE.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ,(SELECT TOP 1 ISNULL(TD012,'')+' 預計到貨:'+CONVERT(nvarchar,CONVERT(DECIMAL(16,2),NUM))  FROM [TK].dbo.VPURTDINVMD WHERE  TD004=TB003 AND TD007=TD007 AND TD012>='{0}') AS '最快採購日'", SEARCHDATE.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ,TB009 AS '庫別'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTB,[TK].dbo.MOCTA,[TK].dbo.INVMB");
                sbSql.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
                sbSql.AppendFormat(@"  AND MB001=TB003");
                sbSql.AppendFormat(@"  AND TB018='Y'");
                sbSql.AppendFormat(@"  AND (TB003 LIKE '1%' OR TB003 LIKE '2%')");
                sbSql.AppendFormat(@"  AND TA003>='{0}'",SEARCHDATE.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND (TB004-TB005)>0");
                sbSql.AppendFormat(@"  AND TB001 NOT  IN ('A513')");
                sbSql.AppendFormat(@"  GROUP BY TB003,TB007,TB009,MB002) AS TEMP");
                sbSql.AppendFormat(@"  WHERE  需求差異量<0 ");
                sbSql.AppendFormat(@"  UNION ALL");
                sbSql.AppendFormat(@"  SELECT 品號,品名,需求量,單位,現有庫存,需求差異量,總採購量,最快採購日 ");
                sbSql.AppendFormat(@"  FROM (");
                sbSql.AppendFormat(@"  SELECT TB003 AS '品號',MB002 AS '品名' ,SUM(TB004-TB005) AS '需求量',TB007 AS '單位'");
                sbSql.AppendFormat(@"  ,(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA001=TB003 AND LA009=TB009) AS '現有庫存'");
                sbSql.AppendFormat(@"  ,(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA001=TB003 AND LA009=TB009)-SUM(TB004-TB005) AS '需求差異量'");
                sbSql.AppendFormat(@"  ,(SELECT ISNULL(CONVERT(DECIMAL(16,2),SUM(NUM)),0) FROM [TK].dbo.VPURTDINVMD WHERE  TD004=TB003 AND TD007=TD007 AND TD012>='{0}') AS '總採購量'", SEARCHDATE.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ,(SELECT TOP 1 ISNULL(TD012,'')+' 預計到貨:'+CONVERT(nvarchar,CONVERT(DECIMAL(16,2),NUM))  FROM [TK].dbo.VPURTDINVMD WHERE  TD004=TB003 AND TD007=TD007 AND TD012>='{0}') AS '最快採購日'", SEARCHDATE.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ,TB009 AS '庫別'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTB,[TK].dbo.MOCTA,[TK].dbo.INVMB");
                sbSql.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
                sbSql.AppendFormat(@"  AND MB001=TB003");
                sbSql.AppendFormat(@"  AND TB018='Y'");
                sbSql.AppendFormat(@"  AND (TB003 LIKE '1%' OR TB003 LIKE '2%')");
                sbSql.AppendFormat(@"  AND TA003>='{0}'", SEARCHDATE.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND (TB004-TB005)>0");
                sbSql.AppendFormat(@"  AND TB001 NOT  IN ('A513')");
                sbSql.AppendFormat(@"  GROUP BY TB003,TB007,TB009,MB002) AS TEMP");
                sbSql.AppendFormat(@"  WHERE  需求差異量>0 ");
                sbSql.AppendFormat(@"  ");
              

                adapterPURTA = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderPURTA = new SqlCommandBuilder(adapterPURTA);
                sqlConn.Open();

                dsPURTA.Clear();
                adapterPURTA.Fill(dsPURTA, "TEMPdsPURTA");
                sqlConn.Close();


                if (dsPURTA.Tables["TEMPdsPURTA"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsPURTA.Tables["TEMPdsPURTA"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(dsPURTA, pathFilePURTA);


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

        public void SEARCHPURTA()
        {
            DateTime SEARCHDATE = DateTime.Now;
            

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TA001 AS '製令',TA002 AS '製令單號',TA003 AS '開單日',TA006 AS '品號',TA034 AS '品名',TA015 AS '數量',TA007 AS '單位',CASE WHEN ISNULL(PURTA001,'')<>'' THEN '已請購' ELSE (CASE WHEN  ISNULL([COMMENT],'')<>'' THEN ''  ELSE '未請購'END ) END  AS '是否請購',PURTA001 AS '請購單',PURTA002 AS '請購單號' ,[COMMENT] AS '備註'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA");
                sbSql.AppendFormat(@"  LEFT JOIN [TKWAREHOUSE].[dbo].[PURTAB] ON TA001=[PURTAB].[MOCTA001] AND TA002=[PURTAB].[MOCTA002] AND TA006=[PURTAB].[MOCTA006]");
                sbSql.AppendFormat(@"  LEFT JOIN [TKWAREHOUSE].[dbo].[MOCINVCHECK] ON TA001=[MOCINVCHECK].[MOCTA001] AND TA002=[MOCINVCHECK].[MOCTA002]");
                sbSql.AppendFormat(@"  WHERE TA003>='{0}'", SEARCHDATE.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TA006 LIKE '4%'");
                sbSql.AppendFormat(@"  AND TA001 NOT IN ('A513') ");
                sbSql.AppendFormat(@"  ");

                adapterPURTA = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderPURTA = new SqlCommandBuilder(adapterPURTA);
                sqlConn.Open();

                dsPURTA.Clear();
                adapterPURTA.Fill(dsPURTA, "TEMPdsPURTA");
                sqlConn.Close();


                if (dsPURTA.Tables["TEMPdsPURTA"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsPURTA.Tables["TEMPdsPURTA"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(dsPURTA, pathFilePURTA);

                       
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

        public void SERACHMAILPURTA()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT [SENDTO],[MAIL] ");
                sbSql.AppendFormat(@"  FROM [TKMQ].[dbo].[MQSENDMAIL] ");
                sbSql.AppendFormat(@"  WHERE [SENDTO]='PUR'  ");
                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  ");

                adapterMAILPURTA = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMAILPURTA = new SqlCommandBuilder(adapterMAILPURTA);
                sqlConn.Open();
                dsMAILPURTA.Clear();
                adapterMAILPURTA.Fill(dsMAILPURTA, "TEMPdsMAILPURTA");
                sqlConn.Close();


                if (dsMAILPURTA.Tables["TEMPdsMAILPURTA"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsMAILPURTA.Tables["TEMPdsMAILPURTA"].Rows.Count >= 1)
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

        public void SETFILEMOCTA()
        {
            if (Directory.Exists(DirectoryNAME))
            {
                //資料夾存在，pathFile
                if (File.Exists(pathFile + ".xlsx"))
                {
                    File.Delete(pathFile + ".xlsx");
                }

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

            if (!File.Exists(pathFileMOCTA + ".xlsx"))
            {
                wBook.SaveAs(pathFileMOCTA, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }



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


            SEARCHMOCTA();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}
        }

        public void SEARCHMOCTA()
        {
            //DateTime SEARCHDATE = DateTime.Now;
            //SEARCHDATE = SEARCHDATE.AddMonths(-1);


            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

               

                sbSql.AppendFormat(@"  SELECT TA001 AS '製令單別',TA002 AS '製令單號',TA003 AS '開單日期',TA006 AS '產品品號',TA034 AS '產品品名',CONVERT(INT,TA015,0) AS'預計產量',TA007 AS '單位','未確認' AS '確認碼',TA026 AS '訂單單別',TA027 AS '訂單單號',TA028 AS '訂單序號'");
                sbSql.AppendFormat(@"  ,CONVERT(INT,ISNULL([NUM],0)) AS '訂單需求量',TD010 AS '訂單單位',CONVERT(INT,(TA015-ISNULL([NUM],0)),0) AS '生產需求的差異數'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].[dbo].[VCOPTDINVMD] ON TA026=TD001 AND TA027=TD002 AND TA028=TD003 ");
                sbSql.AppendFormat(@"  WHERE TA013='N'");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapterMOCTA = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMOCTA = new SqlCommandBuilder(adapterMOCTA);
                sqlConn.Open();
                dsMOCTA.Clear();
                adapterMOCTA.Fill(dsMOCTA, "dsMOCTA");
                sqlConn.Close();


                if (dsMOCTA.Tables["dsMOCTA"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsMOCTA.Tables["dsMOCTA"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(dsMOCTA, pathFileMOCTA);
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

        public void SERACHMAILMOCTA()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT [SENDTO],[MAIL] ");
                sbSql.AppendFormat(@"  FROM [TKMQ].[dbo].[MQSENDMAIL] ");
                sbSql.AppendFormat(@"  WHERE [SENDTO]='MOC'  ");
                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  ");

                adapterMAILMOCTA = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMAILMOCTA = new SqlCommandBuilder(adapterMAILMOCTA);
                sqlConn.Open();
                dsMAILMOCTA.Clear();
                adapterMAILMOCTA.Fill(dsMAILMOCTA, "dsMAILMOCTA");
                sqlConn.Close();


                if (dsMAILMOCTA.Tables["dsMAILMOCTA"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsMAILMOCTA.Tables["dsMAILMOCTA"].Rows.Count >= 1)
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
            SETPATH();
            SETFILE();
            CLEAREXCEL();
            MessageBox.Show("OK");
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETPATH();
            SETFILECOPTE();
            CLEAREXCEL();
            MessageBox.Show("OK");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SETPATH();
            SETFILEPURTA();
            CLEAREXCEL();
            MessageBox.Show("OK");
        }
        private void button3_Click(object sender, EventArgs e)
        {
            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();


            //SETFILE();
            //CLEAREXCEL();
            //SETFILECOPTE();
            //CLEAREXCEL();
            SETFILEPURTA();
            CLEAREXCEL();

            SETFILEMOCTA();
            CLEAREXCEL();


            //SERACHMAIL();
            //SUBJEST.Clear();
            //BODY.Clear();
            //SUBJEST.AppendFormat(@"每日訂單-製令追踨表" + DateTime.Now.ToString("yyyy/MM/dd"));
            //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日訂單-製令追踨表，請查收" + Environment.NewLine + "若訂單沒有相對的製令則需通知製造生管開立");
            //SENDMAIL(SUBJEST, BODY, dsMAIL, pathFile);

            //Thread.Sleep(10000);

            //SERACHMAILCOPTE();
            //SUBJEST.Clear();
            //BODY.Clear();
            //SUBJEST.AppendFormat(@"每日訂單變更追踨表" + DateTime.Now.ToString("yyyy/MM/dd"));
            //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日訂單變更表，請查收" + Environment.NewLine + "請製造生管修改相對的製令");
            //SENDMAIL(SUBJEST, BODY, dsMAILCOPTE, pathFileCOPTE);


            Thread.Sleep(10000);

            SERACHMAILPURTA();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"每日製令-請購表" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每每日製令-請購表，請查收" + Environment.NewLine + " ");
            SENDMAIL(SUBJEST, BODY, dsMAILPURTA, pathFilePURTA);


            Thread.Sleep(10000);

            SERACHMAILMOCTA();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"每日製令-訂單表" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每每日製令-訂單表，請查收" + Environment.NewLine + " ");
            SENDMAIL(SUBJEST, BODY, dsMAILMOCTA, pathFileMOCTA);

            MessageBox.Show("OK");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SETPATH();
            SETFILEMOCTA();
            CLEAREXCEL();
            MessageBox.Show("OK");
        }


        #endregion


    }
}
