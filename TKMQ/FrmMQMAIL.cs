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
using System.Diagnostics;
using System.Threading;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using NPOI.XSSF.UserModel;
using TKITDLL;
using System.Net.Http;
using System.Net;
using System.Xml;

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
        SqlDataAdapter adapterINVMOCTA = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderINVMOCTA = new SqlCommandBuilder();
        SqlDataAdapter adapterPURTB = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderPURTB = new SqlCommandBuilder();
        SqlDataAdapter adapterMOCINVCHECK = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMOCINVCHECK = new SqlCommandBuilder();
        SqlDataAdapter adapterMOCCOP = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMOCCOP = new SqlCommandBuilder();
        SqlDataAdapter adapterINVMC = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderINVMC = new SqlCommandBuilder();
        SqlDataAdapter adapterPURTD = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderPURTD = new SqlCommandBuilder();
        SqlDataAdapter adapterMOCTARE = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMOCTARE = new SqlCommandBuilder();
        SqlDataAdapter adapterLOTCHECK = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderLOTCHECK = new SqlCommandBuilder();
        SqlDataAdapter adapterMOCMANULINE = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMOCMANULINE = new SqlCommandBuilder();

        SqlDataAdapter adapterMAILCOPTE = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILCOPTE = new SqlCommandBuilder();
        SqlDataAdapter adapterMAILPURTA = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILPURTA = new SqlCommandBuilder();
        SqlDataAdapter adapterMAILMOCTA = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILMOCTA = new SqlCommandBuilder();
        SqlDataAdapter adapterMAILINVMOCTA = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILINVMOCTA = new SqlCommandBuilder();
        SqlDataAdapter adapterMAILPURTB = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILPURTB = new SqlCommandBuilder();
        SqlDataAdapter adapterMAILMOCINVCHECK = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILMOCINVCHECK = new SqlCommandBuilder();
        SqlDataAdapter adapterMAILMOCCOP = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILMOCCOP = new SqlCommandBuilder();
        SqlDataAdapter adapterMAILINVMC = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILINVMC = new SqlCommandBuilder();
        SqlDataAdapter adapterMAILPURTD = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILPURTD = new SqlCommandBuilder();
        SqlDataAdapter adapterMAILMOCTARE = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILMOCTARE = new SqlCommandBuilder();
        SqlDataAdapter adapterMAILLOTCHECK = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILLOTCHECK = new SqlCommandBuilder();
        SqlDataAdapter adapterMAILMOCMANULINE = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderMAILMOCMANULINE = new SqlCommandBuilder();

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
        DataSet dsINVMOCTA = new DataSet();
        DataSet dsMAILINVMOCTA = new DataSet();
        DataSet dsPURTB = new DataSet();
        DataSet dsMAILPURTB = new DataSet();
        DataSet dsMOCINVCHECK = new DataSet();
        DataSet dsMAILPMOCINVCHECK = new DataSet();
        DataSet dsMOCCOP = new DataSet();
        DataSet dsMAILMOCCOP = new DataSet();
        DataSet dsINVMC = new DataSet();
        DataSet dsMAILINVMC = new DataSet();
        DataSet dsPURTD = new DataSet();
        DataSet dsMAILPURTD = new DataSet();
        DataSet dsMOCTARE = new DataSet();
        DataSet dsMAILMOCTARE = new DataSet();
        DataSet dsLOTCHECK = new DataSet();
        DataSet dsMAILLOTCHECK = new DataSet();
        DataSet dsMOCMANULINE = new DataSet();
        DataSet dsMAILMOCMANULINE = new DataSet();


        string DATES = null;
        string DirectoryNAME = null;
        string pathFile = null;
        string pathFileCOPTE = null;
        string pathFilePURTA = null;
        string pathFileMOCTA = null;
        string pathFileINVMOCTA = null;
        string pathFilePURTB = null;
        string pathFileMOCINVCHECK = null;
        string pathFileMOCCOP = null;
        string pathFileINVMC = null;
        string pathFilePURTD = null;
        string pathFileMOCTARE = null;
        string pathFileLOTCHECK = null;
        string pathFileMOCMANULINE = null;

        FileInfo info;
        string[] tempFile;
        string tFileName = "";

        public FrmMQMAIL()
        {
            InitializeComponent();

            timer1.Enabled = true;
            timer1.Interval = 1000 * 60;
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
            pathFileINVMOCTA = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日半成品-製令表" + DATES.ToString();
            pathFilePURTB = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日已請購未採購表" + DATES.ToString();
            pathFileMOCINVCHECK = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日物料安全水位檢查表" + DATES.ToString();
            pathFileMOCCOP = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日製令準時完工率數量達交率表" + DATES.ToString();
            pathFileINVMC = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日物料安全水位表" + DATES.ToString();
            pathFilePURTD = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日採購單未結案表" + DATES.ToString();
            pathFileMOCTARE = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日製令重工表" + DATES.ToString();
            pathFileLOTCHECK = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日批號檢查表" + DATES.ToString();
            pathFileMOCMANULINE = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日預排製令表" + DATES.ToString();
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
        public void SENDMAIL(StringBuilder Subject, StringBuilder Body, DataSet SEND, string Attachments)
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

            Attachment attch = new Attachment(Attachments + ".xlsx");
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
                ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                //ex.ToString();
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

            if (!File.Exists(pathFile + ".xlsx"))
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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

               
                sbSql.AppendFormat(@"  
                                    SELECT  MD002 AS '線別',TC053 AS '客戶',TD013 AS '預計交貨日',TD004 AS '訂單品號',TD005 AS '訂單品名',TD006 AS '規格',TD008 AS '訂單量',TD009 AS '出貨量',TD024 AS '贈品量',TD025 AS '贈品已交量',(TD008-TD009+TD024-TD025) AS '總未出貨量',TD010 AS '品號單位',TD001 AS '訂單單別',TD002 AS '訂單單號',TD003 AS '訂單序號',TD016 AS '訂單狀態',MOCTA.TA001 AS '批次轉製令單別',MOCTA.TA002 AS '批次轉製令單號',MOCTA.TA009 AS '製令預計開工日',MOCTA.TA012 AS '製令實際開工日',MOCTA.TA010 AS '製令預計完工日' ,MOCTA.TA014 AS '製令實際完工日',MOCTA.TA006 AS '生產品號',MOCTA.TA034 AS '生產品名',MOCTA.TA007 AS '生產單位',MOCTA.TA015 AS '製令預計產量',MOCTA.TA017 AS '實際入庫數量',COMMENT AS '備註'
                                    ,(CASE WHEN MOCTA.TA011='Y' THEN '已完工' ELSE CASE WHEN MOCTA.TA011='y' THEN '指定完工' ELSE  CASE WHEN MOCTA.TA011='1' THEN '未生產' ELSE CASE WHEN MOCTA.TA011='2' THEN '已發料' ELSE CASE WHEN MOCTA.TA011='3' THEN '生產中' ELSE '' END END END END END)AS '生產進度'
                                    ,(CASE WHEN CONVERT(datetime,MOCTA.TA009)<CONVERT(datetime,MOCTA.TA012) THEN '是' ELSE ''  END ) AS '製令開工異常警示'
                                    ,(CASE WHEN CONVERT(datetime,MOCTA.TA010)<CONVERT(datetime,MOCTA.TA014) THEN '是' ELSE ''  END ) AS '製令完工異常警示'
                                    ,(CASE WHEN MOCTA.TA017<MOCTA.TA015 THEN '是' ELSE ''  END) AS '產量不足'
                                    ,LRPTA.TA001 AS '批次計畫單號'
                                    ,(CASE WHEN ISNULL(MOCTA.TA033,'')<>''  THEN '是' ELSE ''  END )  AS '製令發放'
                                    ,(CASE WHEN CONVERT(datetime,TD013)<=CONVERT(datetime,MOCTA.TA009) THEN '是' ELSE ''  END )  AS '訂單是否延遲生產'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD
                                    LEFT JOIN [TK].dbo.MOCTA ON MOCTA.TA026=TD001 AND MOCTA.TA027=TD002 AND MOCTA.TA028=TD003 AND TD004=MOCTA.TA006
                                    LEFT JOIN [TK].dbo.CMSMD ON CMSMD.MD001=MOCTA.TA021
                                    LEFT JOIN [TK].dbo.LRPTA ON LRPTA.TA023=TD001 AND LRPTA.TA024=TD002 AND LRPTA.TA025=TD003
                                    LEFT JOIN [TKMOC].dbo.MOCCOPCHECK ON COPTA001=TD001 AND COPTA002=TD002 AND COPTA003=TD003 
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TD013>='{0}'
                                    AND TD004 LIKE '4%'
                                    AND (TD008-TD009+TD024-TD025)>0
                                    AND TD021='Y' AND TD016='N'
                                    AND ((COPTD.UDF01='Y'  AND TC001 IN ('A221', 'A222','A223','A227')) OR (TC001 IN ('A228') AND ISNULL(MD002,'')<>''))
                                    ORDER BY TC053,TD013,TD004
                                    ", SEARCHDATE.ToString("yyyyMM") + "01");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = ds1.Tables["TEMPds1"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    ds1.Tables["TEMPds1"].Rows.Add(row);

                    ExportDataSetToExcel(ds1, pathFile);
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

        public void ExportDataSetToExcel(DataSet ds, string TopathFile)
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

                        if (TopathFile.Equals(pathFile.ToString()) && k == 16 && string.IsNullOrEmpty(table.Rows[j].ItemArray[k].ToString()))
                        {
                            string STARTCELL = "A" + (j + 2).ToString();
                            string ENDCELL = "AI" + (j + 2).ToString();
                            Excel.Range newRng = excelApp.get_Range(STARTCELL, ENDCELL);
                            newRng.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Red);

                        }

                        //pathFilePURTA檢查需求差異量是否為負，為負就紅字
                        //string tt = table.Rows[j].ItemArray[k].ToString();

                        if (TopathFile.Equals(pathFilePURTA.ToString()) && k == 9 && !string.IsNullOrEmpty(table.Rows[j].ItemArray[k].ToString()) && Convert.ToDecimal(table.Rows[j].ItemArray[k].ToString()) < 0)
                        {
                            wRange.Select();
                            wRange.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }

                        //wRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.DimGray);
                        // Set the range to fill. pathFileMOCINVCHECK

                        if (TopathFile.Equals(pathFileINVMOCTA) && k == 6 && Convert.ToDecimal(table.Rows[j].ItemArray[k].ToString()) >0)
                        {
                            wRange.Select();
                            wRange.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }

                        if (TopathFile.Equals(pathFileMOCINVCHECK) && k == 4 && Convert.ToDecimal(table.Rows[j].ItemArray[k].ToString()) < 0)
                        {
                            wRange.Select();
                            wRange.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        //pathFileMOCCOP
                        if(TopathFile.Equals(pathFileMOCCOP) && k == 16 && Convert.ToDecimal(table.Rows[j].ItemArray[k].ToString()) < 0)
                        {
                            wRange.Select();
                            wRange.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        if (TopathFile.Equals(pathFileMOCCOP) && k == 18 && Convert.ToDecimal(table.Rows[j].ItemArray[k].ToString()) > 0)
                        {
                            wRange.Select();
                            wRange.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        //pathFileINVMC
                        if (TopathFile.Equals(pathFileINVMC) && k == 7 && Convert.ToDecimal(table.Rows[j].ItemArray[k].ToString()) < 0)
                        {
                            wRange.Select();
                            wRange.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }


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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                sbSql.AppendFormat(@"  
                                    SELECT [SENDTO],[MAIL] 
                                    FROM [TKMQ].[dbo].[MQSENDMAIL] 
                                    WHERE [SENDTO]='COP'");

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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                ;

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"
                                    SELECT TE006 AS '變更原因',TE001 AS '訂單',TE002 AS '訂單號',TE003 AS '訂單序號',TF005 AS '品號',TF006 AS '品名',TF007 AS '規格',TF009 AS '數量',TF020 AS '新贈品量',TF010 AS '單位',TF015 AS '新預交日',TF109 AS '原訂單數量'
                                    FROM [TKMQ].[dbo].[TRIGGERRECORD],[TK].dbo.COPTE
                                    LEFT JOIN [TK].dbo.COPTF ON TE001=TF001 AND TE002=TF002 AND TE003=TF003
                                    WHERE TE001=IDM AND TE002=IDSUB AND TE003=IDNO
                                    AND MAILYN='N'
                                    ORDER BY TE006,TE001,TE002,TF005
  
                                    ");

                adapterCOPTE = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderCOPTE = new SqlCommandBuilder(adapterCOPTE);

                sqlConn.Open();
                dsCOPTE.Clear();
                adapterCOPTE.Fill(dsCOPTE, "TEMPdsCOPTE");
                sqlConn.Close();


                if (dsCOPTE.Tables["TEMPdsCOPTE"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = dsCOPTE.Tables["TEMPdsCOPTE"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    dsCOPTE.Tables["TEMPdsCOPTE"].Rows.Add(row);

                    ExportDataSetToExcel(dsCOPTE, pathFileCOPTE);
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
                if (dsCOPTE.Tables["TEMPdsCOPTE"].Rows.Count >= 1)
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  
                                     SELECT [SENDTO],[MAIL]
                                     FROM [TKMQ].[dbo].[MQSENDMAIL]
                                     WHERE [SENDTO]='COP' 
                                    ");

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
           

            string RUNTIME = DateTime.Now.ToString("HH:mm");
            string hhmm = "09:05";

            label1.Text = "每日執行時間為" + hhmm;
            label2.Text = DateTime.Now.ToString();

            // DayOfWeek 0 開始 (表示星期日) 到 6 (表示星期六)
            string RUNDATE = DateTime.Now.DayOfWeek.ToString("d");//tmp2 = 4 
            string date = "1";


            if (RUNTIME.Equals(hhmm))
            {
                HRAUTORUN();

                if(RUNDATE.Equals(date))
                {
                    HRAUTORUN2();
                }
            }

        }

        public void HRAUTORUN()
        {
            SETPATH();
          
            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();


            //通知原請購人，總務已完成採購
            FIND_UOF_GRAFFAIRS_1005();

            //通知各表單申請人
            PREPARE_UOF_TASK_TASK_APPLICATION();

            //通知各別的被交辨人
            PREPARE_TB_EIP_PRIV_MESS();

            //通知交辨人
            PREPARE_TB_EIP_PRIV_MESS_DIRECTOR();

            //校稿追踨
            PREPAREPROOFREAD();

            //IT檢查網站是否正常
            PREPAREITCHECK();

            //給採購人員，ERP未核單的單別、單號
            PREPARESENDEMAILERPPURCHECK();

            SETFILEMOCMANULINE();
            CLEAREXCEL();
            Thread.Sleep(5000);

            SETFILELOTCHECK();
            CLEAREXCEL();
            Thread.Sleep(5000);

            SETFILEMOCTARE();
            CLEAREXCEL();
            Thread.Sleep(5000);

          
            SETFILEPURTD();
            CLEAREXCEL();
            Thread.Sleep(5000);

            SETFILEINVMC();
            CLEAREXCEL();
            Thread.Sleep(5000);

            SETFILEPURTB();
            CLEAREXCEL();
            Thread.Sleep(5000);

            SETFILEINVMOCTA();
            CLEAREXCEL();
            Thread.Sleep(5000);

            SETFILEMOCTA();
            CLEAREXCEL();
            Thread.Sleep(5000);

            SETFILECOPTE();
            CLEAREXCEL();
            Thread.Sleep(5000);

            SETFILEPURTA();
            //SETFILEPURTA2();
            CLEAREXCEL();
            Thread.Sleep(5000);

            SETFILE();
            CLEAREXCEL();
            Thread.Sleep(5000);

            //MOCMANULINE
            SERACHMAILMOCMANULINE();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"系統通知-每日預排製令表" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日預排製令表，請查收" + Environment.NewLine + " ");
            SENDMAIL(SUBJEST, BODY, dsMAILMOCMANULINE, pathFileMOCMANULINE);
            Thread.Sleep(5000);


            //LOTCHECK
            SERACHMAILLOTCHECK();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"系統通知-每日批號檢查表" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日批號檢查表，請查收 (批號錯誤時，要檢查「批號資料建立作業」內的有效日期、複檢日期是否也錯誤)" + Environment.NewLine + " ");
            SENDMAIL(SUBJEST, BODY, dsMAILLOTCHECK, pathFileLOTCHECK);
            Thread.Sleep(5000);

            //MOCTARE
            SERACHMAILMOCTARE();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"系統通知-每日重工單未結案表" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日重工單未結案表，請查收" + Environment.NewLine + " ");
            SENDMAIL(SUBJEST, BODY, dsMAILMOCTARE, pathFileMOCTARE);
            Thread.Sleep(5000);

            //PURTD
            SERACHMAILPURTD();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"系統通知-每日每日採購單未結案表" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日採購單未結案表，請查收" + Environment.NewLine + " ");
            SENDMAIL(SUBJEST, BODY, dsMAILPURTD, pathFilePURTD);
            Thread.Sleep(5000);


            //INVMC
            //SERACHMAILINVMC();
            //SUBJEST.Clear();
            //BODY.Clear();
            //SUBJEST.AppendFormat(@"每日物料安全水位表" + DateTime.Now.ToString("yyyy/MM/dd"));
            //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日物料安全水位表，請查收" + Environment.NewLine + " ");
            //SENDMAIL(SUBJEST, BODY, dsMAILINVMC, pathFileINVMC);



            SERACHMAILPURTB();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"系統通知-每日已請購未採購表" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日已請購未採購表，請查收" + Environment.NewLine + " ");
            SENDMAIL(SUBJEST, BODY, dsMAILPURTB, pathFilePURTB);

            Thread.Sleep(5000);

            SERACHMAILINVMOCTA();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"系統通知-每日追踨半成品-製令的比對表，是否有半成品呆滯" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日半成品-製令表，請查收" + Environment.NewLine + " ");
            SENDMAIL(SUBJEST, BODY, dsMAILINVMOCTA, pathFileINVMOCTA);

            Thread.Sleep(5000);


            SERACHMAILMOCTA();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"系統通知-每日追踨製令未確認表" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日製令未確認表，請查收" + Environment.NewLine + " ");
            SENDMAIL(SUBJEST, BODY, dsMAILMOCTA, pathFileMOCTA);

            Thread.Sleep(5000);

            SERACHMAILCOPTE();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"系統通知-每日追踨訂單變更追踨表" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日訂單變更表，請查收" + Environment.NewLine + "請製造生管修改相對的製令");
            SENDMAIL(SUBJEST, BODY, dsMAILCOPTE, pathFileCOPTE);


            Thread.Sleep(5000);

            SERACHMAILPURTA();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"系統通知-每日追踨製令-請購表，是否有製令已開但未請購" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每每日製令-請購表，請查收" + Environment.NewLine + " ");
            SENDMAIL(SUBJEST, BODY, dsMAILPURTA, pathFilePURTA);


            Thread.Sleep(5000);

            SERACHMAIL();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"系統通知-每日追踨訂單-製令追踨表，是否有訂單未開製令" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日訂單-製令追踨表，請查收" + Environment.NewLine + "若訂單沒有相對的製令則需通知製造生管開立");
            SENDMAIL(SUBJEST, BODY, dsMAIL, pathFile);

            Thread.Sleep(5000);

            //MessageBox.Show("OK");

        }

        public void HRAUTORUN2()
        {
            SETPATH();

            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();

            SETFILEMOCCOP();
            CLEAREXCEL();
            Thread.Sleep(5000);

            SERACHMAILMOCCOP();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"系統通知-每日製令準時完工率數量達交率表" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日製令準時完工率數量達交率表，請查收" + Environment.NewLine + " ");
            SENDMAIL(SUBJEST, BODY, dsMAILMOCCOP, pathFileMOCCOP);

           

            //MessageBox.Show("OK");

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
            DateTime SEARCHDATE2 = DateTime.Now;
            DateTime SEARCHDATE3 = DateTime.Now.AddDays(7);
            DateTime SEARCHDATE4 = DateTime.Now.AddDays(14);
            DateTime SEARCHDATE5 = DateTime.Now.AddDays(21);
            DateTime SEARCHDATE6 = DateTime.Now.AddDays(30);
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();
                
                
                sbSql.AppendFormat(@"  
                                    SELECT TB003 AS '品號',MB002 AS '品名' 
                                    ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA001=TB003 AND LA009=TB009) AS '現有庫存'
                                    ,(SELECT ISNULL(SUM(TB004-TB005),0) FROM [TK].dbo.MOCTB B,[TK].dbo.MOCTA A WHERE A.TA001=B.TB001 AND A.TA002=B.TB002  AND B.TB018='Y' AND (B.TB003 LIKE '1%' OR B.TB003 LIKE '2%')  AND A.TA003>='{0}'  AND A.TA003<='{2}' AND (B.TB004-B.TB005)>0  AND B.TB001 NOT  IN ('A513') AND MOCTB.TB003=B.TB003) AS '7天內的需求量'
                                    ,(SELECT ISNULL(SUM(TB004-TB005),0) FROM [TK].dbo.MOCTB B,[TK].dbo.MOCTA A WHERE A.TA001=B.TB001 AND A.TA002=B.TB002  AND B.TB018='Y' AND (B.TB003 LIKE '1%' OR B.TB003 LIKE '2%')  AND A.TA003>='{0}'  AND A.TA003<='{3}' AND (B.TB004-B.TB005)>0  AND B.TB001 NOT  IN ('A513') AND MOCTB.TB003=B.TB003) AS '14天內的需求量'
                                    ,(SELECT ISNULL(SUM(TB004-TB005),0) FROM [TK].dbo.MOCTB B,[TK].dbo.MOCTA A WHERE A.TA001=B.TB001 AND A.TA002=B.TB002  AND B.TB018='Y' AND (B.TB003 LIKE '1%' OR B.TB003 LIKE '2%')  AND A.TA003>='{0}'  AND A.TA003<='{4}' AND (B.TB004-B.TB005)>0  AND B.TB001 NOT  IN ('A513') AND MOCTB.TB003=B.TB003) AS '21天內的需求量'
                                    ,(SELECT ISNULL(SUM(TB004-TB005),0) FROM [TK].dbo.MOCTB B,[TK].dbo.MOCTA A WHERE A.TA001=B.TB001 AND A.TA002=B.TB002  AND B.TB018='Y' AND (B.TB003 LIKE '1%' OR B.TB003 LIKE '2%')  AND A.TA003>='{0}'  AND A.TA003<='{5}' AND (B.TB004-B.TB005)>0  AND B.TB001 NOT  IN ('A513') AND MOCTB.TB003=B.TB003) AS '30天內的需求量'
                                    ,SUM(TB004-TB005) AS '需求量',TB007 AS '單位'
                                    ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA001=TB003 AND LA009=TB009)-SUM(TB004-TB005) AS '需求差異量'
                                    ,(SELECT ISNULL(CONVERT(DECIMAL(16,2),SUM(NUM)),0) FROM [TK].dbo.VPURTDINVMD WHERE  TD004=TB003 AND TD007=TD007 AND TD012>='{1}') AS '總採購量'
                                    ,(SELECT TOP 1 ISNULL(TD012,'')+' 預計到貨:'+CONVERT(nvarchar,CONVERT(DECIMAL(16,2),NUM))  FROM [TK].dbo.VPURTDINVMD WHERE  TD004=TB003 AND TD007=TD007 AND TD012>='{1}') AS '最快採購日'
                                    ,TB009 AS '庫別'
                                    FROM [TK].dbo.MOCTB,[TK].dbo.MOCTA,[TK].dbo.INVMB
                                    WHERE TA001=TB001 AND TA002=TB002
                                    AND MB001=TB003
                                    AND TB018='Y'
                                    AND (TB003 LIKE '1%' OR TB003 LIKE '2%')
                                    AND TA003>='{1}'
                                    AND (TB004-TB005)>0
                                    AND TB001 NOT  IN ('A513')
                                    GROUP BY TB003,TB007,TB009,MB002
                                    ORDER BY (SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA001=TB003 AND LA009=TB009),TB003   
  
                                    ", DateTime.Now.ToString("yyyyMMdd"), SEARCHDATE2.ToString("yyyyMMdd"), SEARCHDATE3.ToString("yyyyMMdd"), SEARCHDATE4.ToString("yyyyMMdd"), SEARCHDATE5.ToString("yyyyMMdd"), SEARCHDATE6.ToString("yyyyMMdd"));

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


        public void SETFILEPURTA2()
        {
            DateTime SEARCHDATE2 = DateTime.Now;

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();
                
                sbSql.AppendFormat(@"  
                                    SELECT 品號,品名,需求量,單位,現有庫存,需求差異量,總採購量,最快採購日
                                    FROM (
                                    SELECT TB003 AS '品號',MB002 AS '品名' ,SUM(TB004-TB005) AS '需求量',TB007 AS '單位'
                                    ,(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA001=TB003 AND LA009=TB009) AS '現有庫存'
                                    ,(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA001=TB003 AND LA009=TB009)-SUM(TB004-TB005) AS '需求差異量'
                                    ,(SELECT ISNULL(CONVERT(DECIMAL(16,2),SUM(NUM)),0) FROM [TK].dbo.VPURTDINVMD WHERE  TD004=TB003 AND TD007=TD007 AND TD012>='{0}') AS '總採購量'
                                    ,(SELECT TOP 1 ISNULL(TD012,'')+' 預計到貨:'+CONVERT(nvarchar,CONVERT(DECIMAL(16,2),NUM))  FROM [TK].dbo.VPURTDINVMD WHERE  TD004=TB003 AND TD007=TD007 AND TD012>='{0}') AS '最快採購日'
                                    ,TB009 AS '庫別'
                                    FROM [TK].dbo.MOCTB,[TK].dbo.MOCTA,[TK].dbo.INVMB
                                    WHERE TA001=TB001 AND TA002=TB002
                                    AND MB001=TB003
                                    AND TB018='Y'
                                    AND (TB003 LIKE '1%' OR TB003 LIKE '2%')
                                    AND TA003>='{0}'
                                    AND (TB004-TB005)>0
                                    AND TB001 NOT  IN ('A513')
                                    GROUP BY TB003,TB007,TB009,MB002) AS TEMP
                                    WHERE  需求差異量<0
                                    UNION ALL
                                    SELECT 品號,品名,需求量,單位,現有庫存,需求差異量,總採購量,最快採購日
                                    FROM (
                                    SELECT TB003 AS '品號',MB002 AS '品名' ,SUM(TB004-TB005) AS '需求量',TB007 AS '單位'
                                    ,(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA001=TB003 AND LA009=TB009) AS '現有庫存'
                                    ,(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA001=TB003 AND LA009=TB009)-SUM(TB004-TB005) AS '需求差異量'
                                    ,(SELECT ISNULL(CONVERT(DECIMAL(16,2),SUM(NUM)),0) FROM [TK].dbo.VPURTDINVMD WHERE  TD004=TB003 AND TD007=TD007 AND TD012>='{0}') AS '總採購量'
                                    ,(SELECT TOP 1 ISNULL(TD012,'')+' 預計到貨:'+CONVERT(nvarchar,CONVERT(DECIMAL(16,2),NUM))  FROM [TK].dbo.VPURTDINVMD WHERE  TD004=TB003 AND TD007=TD007 AND TD012>='{0}') AS '最快採購日'
                                    ,TB009 AS '庫別'
                                    FROM [TK].dbo.MOCTB,[TK].dbo.MOCTA,[TK].dbo.INVMB
                                    WHERE TA001=TB001 AND TA002=TB002
                                    AND MB001=TB003
                                    AND TB018='Y'
                                    AND (TB003 LIKE '1%' OR TB003 LIKE '2%')
                                    AND TA003>='{0}'
                                    AND (TB004-TB005)>0
                                    AND TB001 NOT  IN ('A513')
                                    GROUP BY TB003,TB007,TB009,MB002) AS TEMP
                                    WHERE  需求差異量>0
                                    ", SEARCHDATE2.ToString("yyyyMMdd"));


                adapterPURTA = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderPURTA = new SqlCommandBuilder(adapterPURTA);
                sqlConn.Open();

                dsPURTA.Clear();
                adapterPURTA.Fill(dsPURTA, "TEMPdsPURTA");
                sqlConn.Close();


                if (dsPURTA.Tables["TEMPdsPURTA"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = dsPURTA.Tables["TEMPdsPURTA"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    dsPURTA.Tables["TEMPdsPURTA"].Rows.Add(row);

                    ExportToExcel(dsPURTA.Tables["TEMPdsPURTA"], "Sheet1", pathFilePURTA);
                }
                else
                {
                    if (dsPURTA.Tables["TEMPdsPURTA"].Rows.Count >= 1)
                    {
                        ExportToExcel(dsPURTA.Tables["TEMPdsPURTA"], "Sheet1", pathFilePURTA);
                    }
                }

            }
            catch (Exception ex)
            {                
                INSERTLOG(pathFilePURTA,ex.ToString());
            }
            finally
            {

            }

        }

        public void ExportToExcel(DataTable data, string sheetName, string PATH)
        {
            try
            {
                if (Directory.Exists(DirectoryNAME))
                {
                    //資料夾存在，pathFileCOPTE
                    if (File.Exists(PATH + ".xlsx"))
                    {
                        File.Delete(PATH + ".xlsx");
                    }
                }
                else
                {
                    //新增資料夾
                    Directory.CreateDirectory(DirectoryNAME);
                }

                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet(sheetName);
                XSSFRow rowHeader = (XSSFRow)sheet.CreateRow(0);
                ICell icell;

                //填寫表頭
                for (int i = 0; i < data.Columns.Count; i++)
                {
                    //string strValue = data.Columns[i].ColumnName.ToString();
                    XSSFCell cell = (XSSFCell)rowHeader.CreateCell(i);
                    cell.SetCellValue(data.Columns[i].ColumnName.ToString());

                    //建立新的CellStyle
                    ICellStyle CellsStyle = workbook.CreateCellStyle();
                    //建立字型
                    IFont StyleFont = workbook.CreateFont();
                    //設定文字字型
                    StyleFont.FontName = "微軟正黑體";
                    //設定文字大小
                    StyleFont.FontHeightInPoints = 12; //設定文字大小為10pt
                                                       //字的顏色
                                                       //StyleFont.Color = IndexedColors.Red.Index;  
                    CellsStyle.SetFont(StyleFont);
                    // 水平置中
                    CellsStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    // 設定框線 
                    //CellsStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    //CellsStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    //CellsStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    //CellsStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

                    cell.CellStyle = CellsStyle;



                    //rowHead.CreateCell(i, CellType.String).SetCellValue(data.Columns[i].ColumnName.ToString());

                }

                //填寫內容
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    XSSFRow rowItem = (XSSFRow)sheet.CreateRow(i + 1);

                    for (int j = 0; j < data.Columns.Count; j++)
                    {
                        XSSFCell cell = (XSSFCell)rowItem.CreateCell(j);
                        cell.SetCellValue(data.Rows[i][j].ToString());

                        //建立新的CellStyle
                        ICellStyle CellsStyle = workbook.CreateCellStyle();
                        //建立字型
                        IFont StyleFont = workbook.CreateFont();
                        //設定文字字型
                        StyleFont.FontName = "微軟正黑體";
                        //設定文字大小
                        StyleFont.FontHeightInPoints = 12; //設定文字大小為10pt
                                                           //字的顏色
                        if (PATH.Equals(pathFilePURTA) && j == 5 && Convert.ToDecimal(data.Rows[i][j].ToString()) < 0)
                        {

                            StyleFont.Color = IndexedColors.Red.Index;

                        }

                        CellsStyle.SetFont(StyleFont);

                        // 設定框線 
                        //CellsStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                        //CellsStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                        //CellsStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                        //CellsStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

                        cell.CellStyle = CellsStyle;

                        //row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());

                    }
                }

                for (int i = 0; i < data.Columns.Count; i++)
                {
                    sheet.AutoSizeColumn(i);
                }

                using (FileStream stream = File.OpenWrite(PATH + ".xlsx"))
                {
                    workbook.Write(stream);
                    stream.Close();
                }

                GC.Collect();
            }

            catch (Exception ex)
            {
                INSERTLOG(pathFilePURTA, ex.ToString());
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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT TA001 AS '製令',TA002 AS '製令單號',TA003 AS '開單日',TA006 AS '品號',TA034 AS '品名',TA015 AS '數量',TA007 AS '單位',CASE WHEN ISNULL(PURTA001,'')<>'' THEN '已請購' ELSE (CASE WHEN  ISNULL([COMMENT],'')<>'' THEN ''  ELSE '未請購'END ) END  AS '是否請購',PURTA001 AS '請購單',PURTA002 AS '請購單號' ,[COMMENT] AS '備註'
                                    FROM [TK].dbo.MOCTA
                                    LEFT JOIN [TKWAREHOUSE].[dbo].[PURTAB] ON TA001=[PURTAB].[MOCTA001] AND TA002=[PURTAB].[MOCTA002] AND TA006=[PURTAB].[MOCTA006]
                                    LEFT JOIN [TKWAREHOUSE].[dbo].[MOCINVCHECK] ON TA001=[MOCINVCHECK].[MOCTA001] AND TA002=[MOCINVCHECK].[MOCTA002]
                                    WHERE TA003>='{0}'
                                    AND TA006 LIKE '4%'
                                    AND TA001 NOT IN ('A513') 
                                    ", SEARCHDATE.ToString("yyyyMMdd"));

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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  
                                    SELECT [SENDTO],[MAIL]
                                    FROM [TKMQ].[dbo].[MQSENDMAIL]
                                    WHERE [SENDTO]='PUR' 
                                    ");

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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT TA001 AS '製令單別',TA002 AS '製令單號',TA003 AS '開單日期',TA006 AS '產品品號',TA034 AS '產品品名',CONVERT(INT,TA015,0) AS'預計產量',TA007 AS '單位','未確認' AS '確認碼',TA026 AS '訂單單別',TA027 AS '訂單單號',TA028 AS '訂單序號'
                                    ,CONVERT(INT,ISNULL([NUM],0)) AS '訂單需求量',TD010 AS '訂單單位',CONVERT(INT,(TA015-ISNULL([NUM],0)),0) AS '生產需求的差異數'
                                    FROM [TK].dbo.MOCTA
                                    LEFT JOIN [TK].[dbo].[VCOPTDINVMD] ON TA026=TD001 AND TA027=TD002 AND TA028=TD003 
                                    WHERE TA013='N'
                                    ORDER BY TA001,TA002
                 
                                    ");

                adapterMOCTA = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMOCTA = new SqlCommandBuilder(adapterMOCTA);
                sqlConn.Open();
                dsMOCTA.Clear();
                adapterMOCTA.Fill(dsMOCTA, "dsMOCTA");
                sqlConn.Close();


                if (dsMOCTA.Tables["dsMOCTA"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = dsMOCTA.Tables["dsMOCTA"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    dsMOCTA.Tables["dsMOCTA"].Rows.Add(row);

                    ExportDataSetToExcel(dsMOCTA, pathFileMOCTA);
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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  
                                    SELECT [SENDTO],[MAIL] 
                                    FROM [TKMQ].[dbo].[MQSENDMAIL] 
                                    WHERE [SENDTO]='MOC'  
                                    ");

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

        public void INSERTLOG(string SOURCE, string EX)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" INSERT INTO [TKMQ].[dbo].[LOG] ([SOURCE],[EX]) VALUES ('{0}','{1}')", SOURCE, EX);
                sbSql.AppendFormat(" ");
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

            catch
            {

            }
            finally
            { }
        }

        public void SETFILEINVMOCTA()
        {
            if (Directory.Exists(DirectoryNAME))
            {
                //資料夾存在，pathFile
                if (File.Exists(pathFileINVMOCTA + ".xlsx"))
                {
                    File.Delete(pathFileINVMOCTA + ".xlsx");
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

            if (!File.Exists(pathFileINVMOCTA + ".xlsx"))
            {
                wBook.SaveAs(pathFileINVMOCTA, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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


            SEARCHINVMOCTA();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}
        }

        public void SEARCHINVMOCTA()
        {
            //DateTime SEARCHDATE = DateTime.Now;
            //SEARCHDATE = SEARCHDATE.AddMonths(-1);


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                sbSql.AppendFormat(@"  
                                    SELECT  LA001 AS '品號' ,MB002 AS '品名',MB003 AS '規格',LA016 AS '批號'  
                                    ,CAST(SUM(LA005*LA011) AS DECIMAL(18,4)) AS '庫存量' 
                                    ,(SELECT ISNULL(SUM(TB004-TB005),0) FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB WHERE TA001=TB001 AND TA002=TB002 AND TA011 NOT IN ('Y','y') AND TB003=LA001 AND TA003<=CONVERT(nvarchar,DATEADD (MONTH,1,CAST(LA016 AS datetime)),112) AND TA003>=LA016) AS '製令量(批號1個月內)'
                                    ,(CAST(SUM(LA005*LA011) AS DECIMAL(18,4))-(SELECT ISNULL(SUM(TB004-TB005),0) FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB WHERE TA001=TB001 AND TA002=TB002 AND TA011 NOT IN ('Y','y') AND TB003=LA001 AND TA003<=CONVERT(nvarchar,DATEADD (MONTH,1,CAST(LA016 AS datetime)),112) AND TA003>=LA016)) AS '庫存差異量'
                                    ,CONVERT(nvarchar,DATEADD (MONTH,1,CAST(LA016 AS datetime)),112) AS '批號製令期限日'
                                    FROM [TK].dbo.INVLA WITH (NOLOCK) 
                                    LEFT JOIN  [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001  WHERE  (LA009='20005     ')  
                                    GROUP BY  LA001,MB002,MB003,LA016 
                                    HAVING SUM(LA005*LA011)<>0 
                                    ORDER BY  LA001,MB002,MB003,LA016
                  
                                    ");

                adapterINVMOCTA = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderINVMOCTA = new SqlCommandBuilder(adapterINVMOCTA);
                sqlConn.Open();
                dsINVMOCTA.Clear();
                adapterINVMOCTA.Fill(dsINVMOCTA, "dsINVMOCTA");
                sqlConn.Close();


                if (dsINVMOCTA.Tables["dsINVMOCTA"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = dsINVMOCTA.Tables["dsINVMOCTA"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    dsINVMOCTA.Tables["dsINVMOCTA"].Rows.Add(row);

                    ExportDataSetToExcel(dsINVMOCTA, pathFileMOCTA);
                }
                else
                {
                    if (dsINVMOCTA.Tables["dsINVMOCTA"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(dsINVMOCTA, pathFileINVMOCTA);
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

        public void SERACHMAILINVMOCTA()
        {
            try
            {//20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  
                                    SELECT [SENDTO],[MAIL] 
                                    FROM [TKMQ].[dbo].[MQSENDMAIL] 
                                    WHERE [SENDTO]='INVCHECK'  
                                    ");

                adapterMAILINVMOCTA = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMAILINVMOCTA = new SqlCommandBuilder(adapterMAILINVMOCTA);
                sqlConn.Open();
                dsMAILINVMOCTA.Clear();
                adapterMAILINVMOCTA.Fill(dsMAILINVMOCTA, "dsMAILINVMOCTA");
                sqlConn.Close();


                if (dsMAILINVMOCTA.Tables["dsMAILINVMOCTA"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsMAILINVMOCTA.Tables["dsMAILINVMOCTA"].Rows.Count >= 1)
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

        public void SETFILEPURTB()
        {
            if (Directory.Exists(DirectoryNAME))
            {
                //資料夾存在，pathFile
                if (File.Exists(pathFilePURTB + ".xlsx"))
                {
                    File.Delete(pathFilePURTB + ".xlsx");
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

            if (!File.Exists(pathFilePURTB + ".xlsx"))
            {
                wBook.SaveAs(pathFilePURTB, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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


            SEARCHPURTB();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}
        }

        public void SEARCHPURTB()
        {
            DateTime SEARCHDATE = DateTime.Now;
            SEARCHDATE = SEARCHDATE.AddDays(-2);


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


               

                sbSql.AppendFormat(@"  
                                    SELECT MA002 AS '廠商',TB011 AS '需求日',TB001 AS '請購單別',TB002 AS '請購單號',TB003 AS '請購序號',TB004 AS '品號',TB005 AS '品名',TB006 AS '規格',TB008 AS '庫別',TB009 AS '請購數量',TB007  AS '單位' ,TB039 AS '是否採購'
                                    FROM [TK].dbo.PURTA,[TK].dbo.PURTB
                                    LEFT JOIN [TK].dbo.PURMA ON MA001=TB010
                                    WHERE TA001=TB001 AND TA002=TB002 
                                    AND  TA007 IN ('Y','N')
                                    AND  TB039='N'
                                    AND  TB025 NOT IN ('V')
                                    AND  TA003<='{0}'
                                    ORDER BY MA002,TB011
                                    ", SEARCHDATE.ToString("yyyyMMdd"));

                adapterPURTB = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderPURTB = new SqlCommandBuilder(adapterPURTB);
                sqlConn.Open();
                dsPURTB.Clear();
                adapterPURTB.Fill(dsPURTB, "dsPURTB");
                sqlConn.Close();


                if (dsPURTB.Tables["dsPURTB"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = dsPURTB.Tables["dsPURTB"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    dsPURTB.Tables["dsPURTB"].Rows.Add(row);

                    ExportDataSetToExcel(dsPURTB, pathFilePURTB);
                }
                else
                {
                    if (dsPURTB.Tables["dsPURTB"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(dsPURTB, pathFilePURTB);
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

        public void SERACHMAILPURTB()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  
                                    SELECT [SENDTO],[MAIL]
                                    FROM [TKMQ].[dbo].[MQSENDMAIL]
                                    WHERE [SENDTO]='PUR' 
                                    ");

                adapterMAILPURTB = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMAILPURTB = new SqlCommandBuilder(adapterMAILPURTB);
                sqlConn.Open();
                dsMAILPURTB.Clear();
                adapterMAILPURTB.Fill(dsMAILPURTB, "dsMAILPURTB");
                sqlConn.Close();


                if (dsMAILPURTB.Tables["dsMAILPURTB"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsMAILPURTB.Tables["dsMAILPURTB"].Rows.Count >= 1)
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

        public void SETFILEMOCINVCHECK()
        {
            if (Directory.Exists(DirectoryNAME))
            {
                //資料夾存在，pathFile
                if (File.Exists(pathFileMOCINVCHECK + ".xlsx"))
                {
                    File.Delete(pathFileMOCINVCHECK + ".xlsx");
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

            if (!File.Exists(pathFileMOCINVCHECK + ".xlsx"))
            {
                wBook.SaveAs(pathFileMOCINVCHECK, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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


            SEARCHMOCINVCHECK();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}
        }

        public void SEARCHMOCINVCHECK()
        {
            //DateTime SEARCHDATE = DateTime.Now;
            //SEARCHDATE = SEARCHDATE.AddDays(-2);


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                //[TK].dbo.INVMC

                sbSql.AppendFormat(@" 
                                    SELECT [MB001] AS '品號',[MB002] AS '品名',[NUM] AS '數量'
                                    ,(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA001=[MB001] AND LA009='20004')   AS '庫存量' 
                                    ,((SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA001=[MB001] AND LA009='20004')-[NUM]) AS '差異量'
                                    FROM [TKMQ].[dbo].[MOCINVCHECK]
                                    ");

                adapterMOCINVCHECK = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMOCINVCHECK = new SqlCommandBuilder(adapterMOCINVCHECK);
                sqlConn.Open();
                dsMOCINVCHECK.Clear();
                adapterMOCINVCHECK.Fill(dsMOCINVCHECK, "dsMOCINVCHECK");
                sqlConn.Close();


                if (dsMOCINVCHECK.Tables["dsMOCINVCHECK"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = dsMOCINVCHECK.Tables["dsMOCINVCHECK"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    dsMOCINVCHECK.Tables["dsMOCINVCHECK"].Rows.Add(row);

                    ExportDataSetToExcel(dsMOCINVCHECK, pathFileMOCINVCHECK);
                }
                else
                {
                    if (dsMOCINVCHECK.Tables["dsMOCINVCHECK"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(dsMOCINVCHECK, pathFileMOCINVCHECK);
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

        public void SETFILEMOCCOP()
        {
            if (Directory.Exists(DirectoryNAME))
            {
                //資料夾存在，pathFile
                if (File.Exists(pathFileMOCCOP + ".xlsx"))
                {
                    File.Delete(pathFileMOCCOP + ".xlsx");
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

            if (!File.Exists(pathFileMOCCOP + ".xlsx"))
            {
                wBook.SaveAs(pathFileMOCCOP, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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


            SEARCHMOCCOP();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}
        }

        public void SEARCHMOCCOP()
        {
            DateTime SEARCHDATES = DateTime.Now;
            SEARCHDATES = SEARCHDATES.AddDays(-8);
            DateTime SEARCHDATEE = DateTime.Now;
            SEARCHDATEE = SEARCHDATEE.AddDays(-1);



            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

               
              

                sbSql.AppendFormat(@"  
                                    SELECT TC053 AS '客戶',TA001 AS '製令單',TA002 AS '製令編號',TA006 AS '品號',TA034 AS '品名',TA007 AS '生產單位',TA009 AS '預計開工日',TA010 AS '預計完工日',TA014 AS '實際完工日',TA015 AS '預計產量',TA017 AS '已生產量',TA026 AS '訂單別',TA027 AS '訂單號',TA028 AS '訂單序',[OLDNUM] AS '訂單量'
                                    ,(SELECT ISNULL(SUM(TA017),0) FROM [TK].dbo.MOCTA A WHERE A.TA011 IN ('Y','y') AND A.TA026=[COPTD].TD001 AND A.TA027=[COPTD].TD002 AND A.TA028=[COPTD].TD003 AND A.TA006=[COPTD].TD004  ) AS '訂單總生產量'
                                    ,ISNULL(((SELECT ISNULL(SUM(TA017),0) FROM [TK].dbo.MOCTA A WHERE A.TA011 IN ('Y','y') AND A.TA026=[COPTD].TD001 AND A.TA027=[COPTD].TD002 AND A.TA028=[COPTD].TD003 AND A.TA006=[COPTD].TD004  ) -[OLDNUM]),0) AS '生產數量是否滿足訂單'
                                    ,[COPTD].TD013 AS '訂單預交日'
                                    ,ISNULL((CASE WHEN ISNULL(TA014,'')<>'' THEN DATEDIFF (DAY,[COPTD].TD013,TA014) ELSE 999 END),0) AS '是否延遲訂單預交'
                                    ,ISNULL(CASE WHEN ISNULL(TA014,'')<>'' THEN DATEDIFF (DAY,TA010,TA014) ELSE 999 END,0)  AS '是否延遲製令完工'
                                    ,ISNULL((TA017-TA015),0) AS '製令生產數量生否>預計生產'
                                    FROM [TK].dbo.MOCTA
                                    LEFT JOIN [TK].[dbo].[VCOPTDINVMD] ON [VCOPTDINVMD].TD001=TA026 AND [VCOPTDINVMD].TD002=TA027 AND [VCOPTDINVMD].TD003=TA028
                                    LEFT JOIN [TK].[dbo].[COPTD] ON [COPTD].TD001=TA026 AND [COPTD].TD002=TA027 AND [COPTD].TD003=TA028
                                    LEFT JOIN [TK].[dbo].[COPTC] ON [COPTC].TC001=TA026 AND [COPTC].TC002=TA027
                                    WHERE TA001 IN ('A510','A511')
                                    AND TA006 LIKE '4%'
                                    AND TA009>='{0}' AND TA009<='{1}'
                                    ORDER BY TC053,TA006
        
                                    ", SEARCHDATES.ToString("yyyyMMdd"), SEARCHDATEE.ToString("yyyyMMdd"));

                adapterMOCCOP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMOCCOP = new SqlCommandBuilder(adapterMOCCOP);
                
                sqlConn.Open();
                dsMOCCOP.Clear();
                adapterMOCCOP.Fill(dsMOCCOP, "dsMOCCOP");
                sqlConn.Close();


                if (dsMOCCOP.Tables["dsMOCCOP"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = dsMOCCOP.Tables["dsMOCCOP"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    dsMOCCOP.Tables["dsMOCCOP"].Rows.Add(row);

                    ExportDataSetToExcel(dsMOCCOP, pathFileMOCCOP);
                }
                else
                {
                    if (dsMOCCOP.Tables["dsMOCCOP"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(dsMOCCOP, pathFileMOCCOP);
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

        public void SERACHMAILMOCCOP()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                sbSql.AppendFormat(@"  
                                    SELECT [SENDTO],[MAIL] 
                                    FROM [TKMQ].[dbo].[MQSENDMAIL] 
                                    WHERE [SENDTO]='MOC' AND [MAIL]='tk290@tkfood.com.tw' 
                                    ");

                adapterMAILMOCCOP = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMAILMOCCOP = new SqlCommandBuilder(adapterMAILMOCCOP);
                sqlConn.Open();
                dsMAILMOCCOP.Clear();
                adapterMAILMOCCOP.Fill(dsMAILMOCCOP, "dsMAILMOCCOP");
                sqlConn.Close();


                if (dsMAILMOCCOP.Tables["dsMAILMOCCOP"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsMAILMOCCOP.Tables["dsMAILMOCCOP"].Rows.Count >= 1)
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

        public void SETFILEINVMC()
        {
            if (Directory.Exists(DirectoryNAME))
            {
                //資料夾存在，pathFile
                if (File.Exists(pathFileINVMC + ".xlsx"))
                {
                    File.Delete(pathFileINVMC + ".xlsx");
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

            if (!File.Exists(pathFileINVMC + ".xlsx"))
            {
                wBook.SaveAs(pathFileINVMC, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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


            SEARCHINVMC();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}
        }

        public void SEARCHINVMC()
        {
            DateTime SEARCHDATE2 = DateTime.Now;

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

               
                sbSql.AppendFormat(@"  
                                    SELECT MC001 AS '品號',MB002 AS '品名',MC002 AS '庫別',MB004 AS '單位',MC004 AS '安全批量',MC005 AS '補貨點'
                                    ,ISNULL((SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE MC001=LA001 AND LA009=MC002) ,0) AS '目前庫存'
                                    ,ISNULL(((SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE MC001=LA001 AND LA009=MC002) -MC004),0) AS '庫存差異量'
                                    ,(SELECT ISNULL(CONVERT(DECIMAL(16,2),SUM(NUM)),0) FROM [TK].dbo.VPURTDINVMD WHERE  TD004=MC001 AND TD007=TD007 AND TD012>='{0}') AS '總採購量'
                                    ,(SELECT TOP 1 ISNULL(TD012,'')+' 預計到貨:'+CONVERT(nvarchar,CONVERT(DECIMAL(16,2),NUM))  FROM [TK].dbo.VPURTDINVMD WHERE  TD004=MC001 AND TD007=TD007 AND TD012>='{0}') AS '最快採購日'
                                    FROM [TK].dbo.INVMC,[TK].dbo.INVMB
                                    WHERE MC001=MB001
                                    AND MC002=@MC002 AND MC003='201904制定'
                                    ORDER BY ((SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE MC001=LA001 AND LA009=MC002) -MC004),MC001
                                    ", SEARCHDATE2.ToString("yyyyMMdd"));

                adapterINVMC = new SqlDataAdapter(@"" + sbSql, sqlConn);
                adapterINVMC.SelectCommand.Parameters.AddWithValue("@MC002", "20004");

                sqlCmdBuilderINVMC = new SqlCommandBuilder(adapterINVMC);
                

                sqlConn.Open();
                dsINVMC.Clear();
                adapterINVMC.Fill(dsINVMC, "dsINVMC");
                sqlConn.Close();


                if (dsINVMC.Tables["dsINVMC"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = dsINVMC.Tables["dsINVMC"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    dsINVMC.Tables["dsINVMC"].Rows.Add(row);

                    ExportDataSetToExcel(dsINVMC, pathFileINVMC);
                }
                else
                {
                    if (dsINVMC.Tables["dsINVMC"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(dsINVMC, pathFileINVMC);
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
        public void SERACHMAILINVMC()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  
                                    SELECT [SENDTO],[MAIL] 
                                    FROM [TKMQ].[dbo].[MQSENDMAIL] 
                                    WHERE [SENDTO]='MOC'  
                                    ");

                adapterMAILINVMC = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMAILINVMC = new SqlCommandBuilder(adapterMAILINVMC);
                sqlConn.Open();
                dsMAILINVMC.Clear();
                adapterMAILINVMC.Fill(dsMAILINVMC, "dsMAILINVMC");
                sqlConn.Close();


                if (dsMAILINVMC.Tables["dsMAILINVMC"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsMAILINVMC.Tables["dsMAILINVMC"].Rows.Count >= 1)
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

        public void SETFILEPURTD()
        {
            if (Directory.Exists(DirectoryNAME))
            {
                //資料夾存在，pathFile
                if (File.Exists(pathFilePURTD + ".xlsx"))
                {
                    File.Delete(pathFilePURTD + ".xlsx");
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

            if (!File.Exists(pathFilePURTD + ".xlsx"))
            {
                wBook.SaveAs(pathFilePURTD, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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


            SEARCHPURTD();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}
        }

        public void SEARCHPURTD()
        {
            //DateTime SEARCHDATE2 = DateTime.Now;

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT MA002 AS '廠商',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',TD008 AS '採購量',TD015 AS '已進貨',TD009 AS '單位',TD012 AS '預交日',TD001 AS '採購單別',TD002 AS '採購單號',TD003 AS '序號'
                                    FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.PURMA
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TD016='N'
                                    ORDER BY MA001,TD004,TD012
                  
                                    ");

                adapterPURTD = new SqlDataAdapter(@"" + sbSql, sqlConn);
                //adapterPURTD.SelectCommand.Parameters.AddWithValue("@MC002", "20004");

                sqlCmdBuilderPURTD = new SqlCommandBuilder(adapterPURTD);


                sqlConn.Open();
                dsPURTD.Clear();
                adapterPURTD.Fill(dsPURTD, "dsPURTD");
                sqlConn.Close();


                if (dsPURTD.Tables["dsPURTD"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = dsPURTD.Tables["dsPURTD"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    dsPURTD.Tables["dsPURTD"].Rows.Add(row);

                    ExportDataSetToExcel(dsPURTD, pathFilePURTD);
                }
                else
                {
                    if (dsPURTD.Tables["dsPURTD"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(dsPURTD, pathFilePURTD);
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

        public void SERACHMAILMOCTARE()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  
                                    SELECT [SENDTO],[MAIL] 
                                    FROM [TKMQ].[dbo].[MQSENDMAIL] 
                                    WHERE [SENDTO]='MOCTARE'  
                                    ");

                adapterMAILMOCTARE = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMAILMOCTARE = new SqlCommandBuilder(adapterMAILMOCTARE);
                sqlConn.Open();
                dsMAILMOCTARE.Clear();
                adapterMAILMOCTARE.Fill(dsMAILMOCTARE, "dsMAILMOCTARE");
                sqlConn.Close();


                if (dsMAILMOCTARE.Tables["dsMAILMOCTARE"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsMAILMOCTARE.Tables["dsMAILMOCTARE"].Rows.Count >= 1)
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

        public void SERACHMAILMOCMANULINE()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  
                                    SELECT [SENDTO],[MAIL] 
                                    FROM [TKMQ].[dbo].[MQSENDMAIL] 
                                    WHERE [SENDTO]='COP'  
                                    ");

                adapterMAILMOCMANULINE = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMAILMOCMANULINE = new SqlCommandBuilder(adapterMAILMOCMANULINE);
                sqlConn.Open();
                dsMAILMOCMANULINE.Clear();
                adapterMAILMOCMANULINE.Fill(dsMAILMOCMANULINE, "dsMAILMOCMANULINE");
                sqlConn.Close();


                if (dsMAILMOCMANULINE.Tables["dsMAILMOCMANULINE"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsMAILMOCMANULINE.Tables["dsMAILMOCMANULINE"].Rows.Count >= 1)
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

        public void SERACHMAILLOTCHECK()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  
                                    SELECT [SENDTO],[MAIL] 
                                    FROM [TKMQ].[dbo].[MQSENDMAIL] 
                                    WHERE [SENDTO]='LOTCHECK'  
                                    ");

                adapterMAILLOTCHECK = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMAILLOTCHECK = new SqlCommandBuilder(adapterMAILLOTCHECK);
                sqlConn.Open();
                dsMAILLOTCHECK.Clear();
                adapterMAILLOTCHECK.Fill(dsMAILLOTCHECK, "dsMAILLOTCHECK");
                sqlConn.Close();


                if (dsMAILLOTCHECK.Tables["dsMAILLOTCHECK"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsMAILLOTCHECK.Tables["dsMAILLOTCHECK"].Rows.Count >= 1)
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

        public void SERACHMAILPURTD()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  
                                    SELECT [SENDTO],[MAIL] 
                                    FROM [TKMQ].[dbo].[MQSENDMAIL] 
                                    WHERE [SENDTO]='PUR'  
                                    ");

                adapterMAILPURTD = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderMAILPURTD = new SqlCommandBuilder(adapterMAILPURTD);
                sqlConn.Open();
                dsMAILPURTD.Clear();
                adapterMAILPURTD.Fill(dsMAILPURTD, "dsMAILPURTD");
                sqlConn.Close();


                if (dsMAILPURTD.Tables["dsMAILPURTD"].Rows.Count == 0)
                {

                }
                else
                {
                    if (dsMAILPURTD.Tables["dsMAILPURTD"].Rows.Count >= 1)
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

        public void SETFILEMOCTARE()
        {
            if (Directory.Exists(DirectoryNAME))
            {
                //資料夾存在，pathFile
                if (File.Exists(pathFileMOCTARE + ".xlsx"))
                {
                    File.Delete(pathFileMOCTARE + ".xlsx");
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

            if (!File.Exists(pathFileMOCTARE + ".xlsx"))
            {
                wBook.SaveAs(pathFileMOCTARE, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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


            SEARCHMOCTARE();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}
        }

        public void SEARCHMOCTARE()
        {
            //DateTime SEARCHDATE2 = DateTime.Now;

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT TA001 AS '製令單別',TA002 AS '製令單號',TA009 AS '開工日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位'
                                    FROM [TK].dbo.MOCTA
                                    WHERE TA013='Y' AND TA011 NOT IN ('Y','y')
                                    AND TA001 IN ('A521')
                                    ");
         

                adapterMOCTARE = new SqlDataAdapter(@"" + sbSql, sqlConn);
                //adapterPURTD.SelectCommand.Parameters.AddWithValue("@MC002", "20004");

                sqlCmdBuilderMOCTARE = new SqlCommandBuilder(adapterMOCTARE);


                sqlConn.Open();
                dsMOCTARE.Clear();
                adapterMOCTARE.Fill(dsMOCTARE, "dsMOCTARE");
                sqlConn.Close();


                if (dsMOCTARE.Tables["dsMOCTARE"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = dsMOCTARE.Tables["dsMOCTARE"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    dsMOCTARE.Tables["dsMOCTARE"].Rows.Add(row);

                    ExportDataSetToExcel(dsMOCTARE, pathFileMOCTARE);
                }
                else
                {
                    if (dsMOCTARE.Tables["dsMOCTARE"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(dsMOCTARE, pathFileMOCTARE);
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

        public void SETFILELOTCHECK()
        {
            if (Directory.Exists(DirectoryNAME))
            {
                //資料夾存在，pathFile
                if (File.Exists(pathFileLOTCHECK + ".xlsx"))
                {
                    File.Delete(pathFileLOTCHECK + ".xlsx");
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

            if (!File.Exists(pathFileLOTCHECK + ".xlsx"))
            {
                wBook.SaveAs(pathFileLOTCHECK, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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


            SEARCHLOTCHECK();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}
        }

        public void SEARCHLOTCHECK()
        {
            //DateTime SEARCHDATE2 = DateTime.Now;

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


               
                sbSql.AppendFormat(@"  
                                    SELECT TH004 AS '品號',TH005 AS '品名',TH010 AS '批號',TH036 AS '有效日',TH117 AS '製造日',TH001 AS '單別',TH002 AS '單號',TH003 AS '序號',COMMET AS '備註' 
                                    FROM 
                                    ( 
                                    SELECT TG003,TH004,TH005,TH010,TH036,TH117,TH001,TH002,TH003,'批號<>有效日' AS COMMET 
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH 
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TH030='Y' 
                                    AND TH004 LIKE '1%' 
                                    AND TH010<>TH036 
                                    UNION ALL 
                                    SELECT TG003,TH004,TH005,TH010,TH036,TH117,TH001,TH002,TH003,'批號<>製造日' AS COMMET 
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH 
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TH030='Y' 
                                    AND TH004 LIKE '2%' 
                                    AND TH010<>TH117 
                                    UNION ALL 
                                    SELECT TG003,TH004,TH005,TH010,TH036,TH117,TH001,TH002,TH003,'批號<>製造日' AS COMMET 
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH 
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TH030='Y' 
                                    AND TH004 LIKE '3%' 
                                    AND TH010<>TH117 
                                    UNION ALL 
                                    SELECT TG003,TH004,TH005,TH010,TH036,TH117,TH001,TH002,TH003,'批號<>有效日' AS COMMET 
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH 
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TH030='Y' 
                                    AND TH004 LIKE '4%' 
                                    AND TH010<>TH036 
                                    UNION ALL 
                                    SELECT TG003,TH004,TH005,TH010,TH036,TH117,TH001,TH002,TH003,'批號<>有效日' AS COMMET 
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH 
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TH030='Y' 
                                    AND TH004 LIKE '5%' 
                                    AND TH010<>TH036 
                                    UNION ALL 
                                    SELECT TF003,TG004,TG005,TG017,TG018,TG040,TG001,TG002,TG003,'批號<>製造日' AS COMMET 
                                    FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG 
                                    WHERE TF001=TG001 AND TF002=TG002 
                                    AND TF003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TG022='Y' 
                                    AND TG004 LIKE '3%'  
                                    AND TG004 NOT LIKE '307%' 
                                    AND TG004 NOT LIKE '308%' 
                                    AND TG004 NOT LIKE '309%' 
                                    AND TG017<>TG040 
                                    UNION ALL 
                                    SELECT TF003,TG004,TG005,TG017,TG018,TF003,TG001,TG002,TG003,'批號<>有效日' AS COMMET 
                                    FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG 
                                    WHERE TF001=TG001 AND TF002=TG002 
                                    AND TF003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TG022='Y' 
                                    AND TG004 LIKE '4%' 
                                    AND TG004 NOT LIKE '408%' 
                                    AND TG004 NOT LIKE '409%' 
                                    AND TG017<>TG018 
                                    UNION ALL 
                                    SELECT TH003,TI004,TI005,TI010,TI011,TI061,TI001,TI002,TI003,'批號<>製造日' AS COMMET 
                                    FROM [TK].dbo.MOCTH,[TK].dbo.MOCTI 
                                    WHERE TH001=TI001 AND TH002=TI002 
                                    AND TI061>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TI004 LIKE '3%'   
                                    AND TI037='Y' 
                                    AND TI010<>TI061 
                                    AND TI001+TI002+TI003 NOT IN ('A591201906240010001','A591201911220010001','A591201911250030001')  
                                    UNION ALL 
                                    SELECT TH003,TI004,TI005,TI010,TI011,TI061,TI001,TI002,TI003,'批號<>有效日' AS COMMET 
                                    FROM [TK].dbo.MOCTH,[TK].dbo.MOCTI 
                                    WHERE TH001=TI001 AND TH002=TI002 
                                    AND TI061>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TI004 LIKE '4%' 
                                    AND TI037='Y' 
                                    AND TI010<>TI011  
                                    UNION ALL 
                                    SELECT TG003,TH004,TH005,TH010,TH036,TH117,TH001,TH002,TH003,'批號日錯誤' AS COMMET 
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH 
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TH030='Y' 
                                    AND TH004 LIKE '1%' 
                                    AND ISDATE(TH010)<>1
                                    AND TH009 NOT LIKE '21%'
                                    UNION ALL 
                                    SELECT TF003,TG004,TG005,TG017,TG018,TG040,TG001,TG002,TG003,'批號日錯誤' AS COMMET 
                                    FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG 
                                    WHERE TF001=TG001 AND TF002=TG002 
                                    AND TF003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TG022='Y' 
                                    AND TG004 LIKE '3%'  
                                    AND TG004 NOT LIKE '307%' 
                                    AND TG004 NOT LIKE '308%' 
                                    AND TG004 NOT LIKE '309%' 
                                    AND ISDATE(TG017)<>1
                                    ) 
                                    AS TEMP 
                                    ORDER BY TH004  
                                    ");


                adapterLOTCHECK = new SqlDataAdapter(@"" + sbSql, sqlConn);
                //adapterPURTD.SelectCommand.Parameters.AddWithValue("@MC002", "20004");

                sqlCmdBuilderLOTCHECK = new SqlCommandBuilder(adapterLOTCHECK);


                sqlConn.Open();
                dsLOTCHECK.Clear();
                adapterLOTCHECK.Fill(dsLOTCHECK, "dsLOTCHECK");
                sqlConn.Close();


                if (dsLOTCHECK.Tables["dsLOTCHECK"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = dsLOTCHECK.Tables["dsLOTCHECK"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    dsLOTCHECK.Tables["dsLOTCHECK"].Rows.Add(row);

                    ExportDataSetToExcel(dsLOTCHECK, pathFileLOTCHECK);
                }
                else
                {
                    if (dsLOTCHECK.Tables["dsLOTCHECK"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(dsLOTCHECK, pathFileLOTCHECK);
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


        public void SETFILEMOCMANULINE()
        {
            if (Directory.Exists(DirectoryNAME))
            {
                //資料夾存在，pathFile
                if (File.Exists(pathFileMOCMANULINE + ".xlsx"))
                {
                    File.Delete(pathFileMOCMANULINE + ".xlsx");
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

            if (!File.Exists(pathFileLOTCHECK + ".xlsx"))
            {
                wBook.SaveAs(pathFileMOCMANULINE, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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


            SEARCHMOCMANULINE();

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}
        }

        public void SEARCHMOCMANULINE()
        {
            //DateTime SEARCHDATE2 = DateTime.Now;

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                sbSql.AppendFormat(@"  
                                  SELECT MANU AS '線別' ,TC053 AS '客戶',MV002 AS '業務員',MANUDATE AS '生產日',[MB002] AS '品名',BAR AS '桶數',NUM AS '數量',PACKAGE AS '包裝數',TD001 AS '訂單',TD002 AS '訂單單號',TD003 AS '訂單序號',MOCTA001 AS '製令',MOCTA002 AS '製令號'
                                    FROM (
                                    SELECT  [MOCMANULINE].[MANU] ,CONVERT(nvarchar,[MOCMANULINE].[MANUDATE],112) MANUDATE,[MOCMANULINE].[MB002]
                                    ,ISNULL([MOCMANULINE].[BAR],0) BAR,ISNULL([MOCMANULINE].[NUM],0) NUM,ISNULL([MOCMANULINE].[PACKAGE],0) PACKAGE
                                    ,[MOCMANULINE].[COPTD001] AS TD001
                                    ,[MOCMANULINE].[COPTD002] AS TD002
                                    ,[MOCMANULINE].[COPTD003] AS TD003
                                    ,[COPTC].TC053,[CMSMV].MV002
                                    ,ISNULL([MOCMANULINERESULT].[MOCTA001],'') AS 'MOCTA001' 
                                    ,ISNULL([MOCMANULINERESULT].[MOCTA002],'') AS 'MOCTA002' 
                                    ,(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCMANULINERESULT].[MOCTA001] AND TG015=[MOCMANULINERESULT].[MOCTA002])+(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCTA].TA001 AND TG015=[MOCTA].TA002)  AS '入庫量'  
                                    ,(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCMANULINERESULT].[MOCTA001] AND TG015=[MOCMANULINERESULT].[MOCTA002]) AS '入庫量A'  
                                    ,(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCTA].TA001 AND TG015=[MOCTA].TA002)  AS '入庫量B'              
                                    ,[MOCMANULINEMERGE].[NO],[MOCTA].TA033,ISNULL([MOCMANULINERESULT].[MOCTA001],'') AS MOCTA001A,ISNULL([MOCMANULINERESULT].[MOCTA002],'')  AS MOCTA002A,ISNULL([MOCTA].TA001,'')  AS MOCTA001B,ISNULL([MOCTA].TA002,'')  AS MOCTA002B  
                                    FROM [TKMOC].[dbo].[MOCMANULINE]
                                    LEFT JOIN [TK].dbo.[COPTD] ON [MOCMANULINE].[COPTD001]=[COPTD].TD001 AND [MOCMANULINE].[COPTD002]=[COPTD].TD002 AND[MOCMANULINE].[COPTD003]=[COPTD].TD003 
                                    LEFT JOIN [TK].dbo.[COPTC] ON [COPTD].TD001=[COPTC].TC001 AND [COPTD].TD002=[COPTC].TC002
                                    LEFT JOIN [TK].dbo.[CMSMV] ON [CMSMV].MV001=[COPTC].TC006
                                    LEFT JOIN [TKMOC].[dbo].[MOCMANULINERESULT] ON [MOCMANULINERESULT].[SID]=[MOCMANULINE].[ID]
                                    LEFT JOIN [TKMOC].[dbo].[MOCMANULINEMERGE] ON [MOCMANULINEMERGE].[SID]=[MOCMANULINE].[ID]  
                                    LEFT JOIN [TK].dbo.[MOCTA] ON [MOCTA].TA033=[MOCMANULINEMERGE].[NO]  
                                    WHERE CONVERT(nvarchar,[MOCMANULINE].[MANUDATE],112)>='{0}' 
                                    AND [MOCMANULINE].[MB001] NOT IN (SELECT MB001 FROM  [TKMOC].[dbo].[MOCMANULINELIMITBARCOUNT])
                                    UNION ALL  
                                    SELECT  [MOCMANULINETEMP].[MANU] ,CONVERT(nvarchar,dateadd(ms,-3,dateadd(yy, datediff(yy,0,getdate())+2, 0)) ,112) MANUDATE,[MOCMANULINETEMP].[MB002]
                                    ,ISNULL([MOCMANULINETEMP].[BAR],0) BAR,ISNULL([MOCMANULINETEMP].[NUM],0) NUM,ISNULL([MOCMANULINETEMP].[PACKAGE],0) PACKAGE
                                    ,[MOCMANULINETEMP].[COPTD001] AS TD001
                                    ,[MOCMANULINETEMP].[COPTD002] AS TD002
                                    ,[MOCMANULINETEMP].[COPTD003] AS TD003
                                    ,[COPTC].TC053,[CMSMV].MV002
                                    ,ISNULL([MOCMANULINERESULT].[MOCTA001],'') AS 'MOCTA001' 
                                    ,ISNULL([MOCMANULINERESULT].[MOCTA002],'') AS 'MOCTA002' 
                                    ,(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCMANULINERESULT].[MOCTA001] AND TG015=[MOCMANULINERESULT].[MOCTA002])+(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCTA].TA001 AND TG015=[MOCTA].TA002)  AS '入庫量'  
                                    ,(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCMANULINERESULT].[MOCTA001] AND TG015=[MOCMANULINERESULT].[MOCTA002]) AS '入庫量A'  
                                    ,(SELECT ISNULL(SUM(TG011),0) FROM  [TK].dbo.[MOCTG] WHERE TG014=[MOCTA].TA001 AND TG015=[MOCTA].TA002)  AS '入庫量B'  
                                    ,[MOCMANULINEMERGE].[NO],[MOCTA].TA033,ISNULL([MOCMANULINERESULT].[MOCTA001],'') AS MOCTA001A,ISNULL([MOCMANULINERESULT].[MOCTA002],'')  AS MOCTA002A,ISNULL([MOCTA].TA001,'')  AS MOCTA001B,ISNULL([MOCTA].TA002,'')  AS MOCTA002B  
                                    FROM [TKMOC].[dbo].[MOCMANULINETEMP]  
                                    LEFT JOIN [TK].dbo.[COPTD] ON [MOCMANULINETEMP].[COPTD001]=[COPTD].TD001 AND [MOCMANULINETEMP].[COPTD002]=[COPTD].TD002 AND[MOCMANULINETEMP].[COPTD003]=[COPTD].TD003   
                                    LEFT JOIN [TK].dbo.[COPTC] ON [COPTD].TD001=[COPTC].TC001 AND [COPTD].TD002=[COPTC].TC002  
                                    LEFT JOIN [TK].dbo.[CMSMV] ON [CMSMV].MV001=[COPTC].TC006  
                                    LEFT JOIN [TKMOC].[dbo].[MOCMANULINE] ON [MOCMANULINE].ID=[MOCMANULINETEMP].TID  
                                    LEFT JOIN [TKMOC].[dbo].[MOCMANULINERESULT] ON [MOCMANULINERESULT].[SID]=[MOCMANULINE].[ID]  
                                    LEFT JOIN [TKMOC].[dbo].[MOCMANULINEMERGE] ON [MOCMANULINEMERGE].[SID]=[MOCMANULINE].[ID]  
                                    LEFT JOIN [TK].dbo.[MOCTA] ON [MOCTA].TA033=[MOCMANULINEMERGE].[NO]  
                                    WHERE CONVERT(nvarchar,[MOCMANULINETEMP].[MANUDATE],112)>='{0}' 
                                    AND [MOCMANULINETEMP].TID IS NULL  
                                    AND [MOCMANULINE].[MB001] NOT IN (SELECT MB001 FROM  [TKMOC].[dbo].[MOCMANULINELIMITBARCOUNT])
                                    ) AS TEMP
                                    ORDER BY  TEMP.[MANU],CONVERT(nvarchar, TEMP.[MANUDATE],112)    
                                    ", DateTime.Now.ToString("yyyyMMdd"));


                adapterMOCMANULINE = new SqlDataAdapter(@"" + sbSql, sqlConn);
                //adapterPURTD.SelectCommand.Parameters.AddWithValue("@MC002", "20004");

                sqlCmdBuilderMOCMANULINE = new SqlCommandBuilder(adapterMOCMANULINE);


                sqlConn.Open();
                dsMOCMANULINE.Clear();
                adapterMOCMANULINE.Fill(dsMOCMANULINE, "dsMOCMANULINE");
                sqlConn.Close();


                if (dsMOCMANULINE.Tables["dsMOCMANULINE"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = dsMOCMANULINE.Tables["dsMOCMANULINE"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //新增資料至DataTable的dt內
                    dsMOCMANULINE.Tables["dsMOCMANULINE"].Rows.Add(row);

                    ExportDataSetToExcel(dsMOCMANULINE, pathFileMOCMANULINE);
                }
                else
                {
                    if (dsMOCMANULINE.Tables["dsMOCMANULINE"].Rows.Count >= 1)
                    {
                        ExportDataSetToExcel(dsMOCMANULINE, pathFileMOCMANULINE);
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


        public void ADDLOG(DateTime DATES,string SOURCE,string EX)
        {
            Guid NEWGUID = new Guid();
            NEWGUID = Guid.NewGuid();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKMQ].[dbo].[LOG]
                                    ([ID],[DATES],[SOURCE],[EX])
                                    VALUES 
                                    ('{0}','{1}','{2}','{3}')
                                   ", NEWGUID, DATES.ToString("yyyy/MM/dd HH:mm:ss"), SOURCE, EX);



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
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        /// <summary>
        /// 準備寄給採購人員跟生管
        /// ERP 採購相關單別、單號未核準的明細 及 昨天該到貨的採購單，但沒有進貨明細數量或進貨數量少於採購數量
        /// </summary>
        public void PREPARESENDEMAILERPPURCHECK()
        {
            DataSet DSPURCHECK = ERPPURCHECK();
            DataSet DSPURTDCHECK = ERPPURTDCHECK();
            DataSet DSTKPUR_PURTATBCHAGE_DCHECK = TKPUR_PURTATBCHAGE_DCHECK();
            DataSet DS_PURTB_NOTIN_PURTD = FRIN_DS_PURTB_NOTIN_PURTD();


            try
            {
                StringBuilder SUBJEST = new StringBuilder();
                StringBuilder BODY = new StringBuilder();

                ////加上附圖
                //string path = System.Environment.CurrentDirectory+@"/Images/emaillogo.jpg";
                //LinkedResource res = new LinkedResource(path);
                //res.ContentId = Guid.NewGuid().ToString();

                SUBJEST.Clear();
                BODY.Clear();

             
                SUBJEST.AppendFormat(@"系統通知-老楊食品-ERP 採購相關單別、單號未核準的明細 及 本月該到貨的採購單，但沒有進貨明細數量或進貨數量少於採購數量 及 請購變更單不在採購變更單 ，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);
               
                //ERP 採購相關單別、單號未核準的明細
                //
                BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                    + "<br>" + "ERP 採購相關單別、單號未核準的明細如下"
                   
                    );


                if (DSPURCHECK.Tables[0].Rows.Count > 0)
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                    BODY.AppendFormat(@"<table> ");
                    BODY.AppendFormat(@"<tr >");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">部門</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">類別</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">單別</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">單號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">變更版次</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">是否送簽</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">備註</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">UOF的單號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">目前簽核人</th>");
                    BODY.AppendFormat(@"</tr> ");

                    foreach (DataRow DR in DSPURCHECK.Tables[0].Rows)
                    {

                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["部門"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["單別"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TC001"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TC002"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["SERNO"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["UDF01"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["UDF02"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["DOC_NBR"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["NAME"].ToString() + "</td>");
                        BODY.AppendFormat(@"</tr> ");

                        //BODY.AppendFormat("<span></span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br> " + "品名     " + DR["TD005"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購數量 " + DR["TD008"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購單位 " + DR["TD009"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>");
                    }
                    BODY.AppendFormat(@"</table> ");
                }
                else
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "本日無資料");
                }



                //請購變更單不在採購變更單
                //
                BODY.AppendFormat(" "
                    + "<br>" + "請購變更單不在採購變更單的明細如下"

                    );


                if (DSTKPUR_PURTATBCHAGE_DCHECK.Tables[0].Rows.Count > 0)
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                    BODY.AppendFormat(@"<table> ");
                    BODY.AppendFormat(@"<tr >");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">變更版號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">請購單</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">請購單號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">單頭備註</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">請購序號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品名</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">請購教量</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">需求日</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">單身備註</th>");

                    BODY.AppendFormat(@"</tr> ");

                    foreach (DataRow DR in DSTKPUR_PURTATBCHAGE_DCHECK.Tables[0].Rows)
                    {

                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["變更版號"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["請購單"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["請購單號"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["單頭備註"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["請購序號"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["品號"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["品名"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["請購教量"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["需求日"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["單身備註"].ToString() + "</td>");

                        BODY.AppendFormat(@"</tr> ");

                        //BODY.AppendFormat("<span></span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br> " + "品名     " + DR["TD005"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購數量 " + DR["TD008"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購單位 " + DR["TD009"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>");
                    }
                    BODY.AppendFormat(@"</table> ");
                }
                else
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "本日無資料");
                }


                //昨天該到貨的採購單，但沒有進貨明細數量或進貨數量少於採購數量                
                //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                //
                BODY.AppendFormat(" "
                    + "<br>" + "本月 該到貨的採購單，但沒有進貨明細數量或進貨數量少於採購數量"
                   
                    );


                if (DSPURTDCHECK.Tables[0].Rows.Count > 0)
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                    BODY.AppendFormat(@"<table> ");
                    BODY.AppendFormat(@"<tr >");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">預計到貨日</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">廠商</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">廠商名稱</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購單別</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購單號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購序號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品名</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購數量</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">單位</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">已進貨總數量</th>");
                    BODY.AppendFormat(@"</tr> ");

                    foreach (DataRow DR in DSPURTDCHECK.Tables[0].Rows)
                    {

                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TD012"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TC004"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["MA002"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TC001"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TC002"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TD003"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TD004"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TD005"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TD008"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TD009"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["SUMTH007"].ToString() + "</td>");
                        BODY.AppendFormat(@"</tr> ");

                        //BODY.AppendFormat("<span></span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br> " + "品名     " + DR["TD005"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購數量 " + DR["TD008"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購單位 " + DR["TD009"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>");
                    }
                    BODY.AppendFormat(@"</table> ");
                }
                else
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "本日無資料");
                }

                //
                //已存在的請購單，需求日>=今日，但未有採購單
                BODY.AppendFormat(" "
                    + "<br>" + "已存在的請購單，需求日>=今日，但未有採購單"

                    );


                if (DS_PURTB_NOTIN_PURTD.Tables[0].Rows.Count > 0)
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                    BODY.AppendFormat(@"<table> ");
                    BODY.AppendFormat(@"<tr >");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">廠商</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">需求日</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">請購單別</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">請購單號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">請購序號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品名</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">規格</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">庫別</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">請購數量</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">單位</th>");
                    BODY.AppendFormat(@"</tr> ");

                    foreach (DataRow DR in DS_PURTB_NOTIN_PURTD.Tables[0].Rows)
                    {

                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["廠商"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["需求日"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["請購單別"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["請購單號"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["請購序號"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["品號"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["品名"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["規格"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["庫別"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["請購數量"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["單位"].ToString() + "</td>");
                        BODY.AppendFormat(@"</tr> ");

                        //BODY.AppendFormat("<span></span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br> " + "品名     " + DR["TD005"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購數量 " + DR["TD008"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購單位 " + DR["TD009"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>");
                    }
                    BODY.AppendFormat(@"</table> ");
                }
                else
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "本日無資料");
                }

                BODY.AppendFormat(" "
                             + "<br>" + "謝謝"

                             + "</span><br>");



                SENDEMAILERPPURCHECK(SUBJEST, BODY);

            }
            catch
            {

            }
            finally
            {

            }
        }

        /// <summary>
        /// 準備寄給採購人員，ERP未核單的單別、單號
        /// </summary>
        public DataSet ERPPURCHECK()
        {
            DataSet DSPURCHECK = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT 部門,單別,TC001,TC002,SERNO,UDF01,UDF02,DOC_NBR
                                    ,(SELECT  TB_EB_USER.NAME+' ' FROM [192.168.1.223].[UOF].dbo.TB_EB_USER,[192.168.1.223].[UOF].[dbo].View_TB_WKF_TASK_NODE,[192.168.1.223].[UOF].[dbo].View_TB_WKF_TASK WHERE TB_EB_USER.USER_GUID=View_TB_WKF_TASK_NODE.ORIGINAL_SIGNER AND View_TB_WKF_TASK_NODE.TASK_ID=View_TB_WKF_TASK.TASK_ID AND NODE_STATUS='1' AND ISNULL(SIGN_STATUS,'')='' AND View_TB_WKF_TASK.DOC_NBR=TEMP.DOC_NBR FOR XML PATH('')) AS 'NAME'
                                    FROM 
                                    (
                                    SELECT  DISTINCT '採購單' AS '單別','' AS '部門',TC001,TC002,'' AS SERNO,UDF01,UDF02,View_TB_WKF_EXTERNAL_TASK.DOC_NBR,TB_EB_USER.NAME
                                    FROM [TK].dbo.PURTC
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].View_TB_WKF_EXTERNAL_TASK ON View_TB_WKF_EXTERNAL_TASK.EXTERNAL_FORM_NBR LIKE TC001+TC002+'%' COLLATE Chinese_Taiwan_Stroke_BIN
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].View_TB_WKF_TASK ON View_TB_WKF_EXTERNAL_TASK.DOC_NBR=View_TB_WKF_TASK.DOC_NBR
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].View_TB_WKF_TASK_NODE  ON View_TB_WKF_TASK_NODE.TASK_ID=View_TB_WKF_TASK.TASK_ID AND NODE_STATUS='1' AND ISNULL(SIGN_STATUS,'')=''
                                    LEFT JOIN [192.168.1.223].[UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=View_TB_WKF_TASK_NODE.ORIGINAL_SIGNER
                                    WHERE TC014='N' 
                                    UNION ALL
                                    SELECT  DISTINCT '採購單變更' AS '單別','' AS '部門',TE001,TE002,TE003,UDF01,UDF02,View_TB_WKF_EXTERNAL_TASK.DOC_NBR,TB_EB_USER.NAME
                                    FROM [TK].dbo.PURTE
                                    LEFT JOIN  [192.168.1.223].[UOF].[dbo].View_TB_WKF_EXTERNAL_TASK ON View_TB_WKF_EXTERNAL_TASK.EXTERNAL_FORM_NBR LIKE TE001+TE002+TE003+'%' COLLATE Chinese_Taiwan_Stroke_BIN
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].View_TB_WKF_TASK ON View_TB_WKF_EXTERNAL_TASK.DOC_NBR=View_TB_WKF_TASK.DOC_NBR
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].View_TB_WKF_TASK_NODE  ON View_TB_WKF_TASK_NODE.TASK_ID=View_TB_WKF_TASK.TASK_ID AND NODE_STATUS='1' AND ISNULL(SIGN_STATUS,'')=''
                                    LEFT JOIN [192.168.1.223].[UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=View_TB_WKF_TASK_NODE.ORIGINAL_SIGNER
                                    WHERE TE017='N' 
                                    UNION ALL
                                    SELECT  DISTINCT '採購核價單' AS '單別','' AS '部門',TL001,TL002,'',UDF01,UDF02,View_TB_WKF_EXTERNAL_TASK.DOC_NBR,TB_EB_USER.NAME
                                    FROM [TK].dbo.PURTL
                                    LEFT JOIN  [192.168.1.223].[UOF].[dbo].View_TB_WKF_EXTERNAL_TASK ON View_TB_WKF_EXTERNAL_TASK.EXTERNAL_FORM_NBR LIKE TL001+TL002+'%' COLLATE Chinese_Taiwan_Stroke_BIN
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].View_TB_WKF_TASK ON View_TB_WKF_EXTERNAL_TASK.DOC_NBR=View_TB_WKF_TASK.DOC_NBR
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].View_TB_WKF_TASK_NODE  ON View_TB_WKF_TASK_NODE.TASK_ID=View_TB_WKF_TASK.TASK_ID AND NODE_STATUS='1' AND ISNULL(SIGN_STATUS,'')=''
                                    LEFT JOIN [192.168.1.223].[UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=View_TB_WKF_TASK_NODE.ORIGINAL_SIGNER
                                    WHERE TL002>='20220101'
                                    AND TL006='N'
                                    UNION ALL
                                    SELECT DISTINCT '請購單' AS '單別',ME002 AS '部門',TA001,TA002,'',PURTA.UDF01,PURTA.UDF02,View_TB_WKF_EXTERNAL_TASK.DOC_NBR,TB_EB_USER.NAME
                                    FROM [TK].dbo.PURTA
                                    LEFT JOIN  [192.168.1.223].[UOF].[dbo].View_TB_WKF_EXTERNAL_TASK ON View_TB_WKF_EXTERNAL_TASK.EXTERNAL_FORM_NBR LIKE TA001+TA002+'%' COLLATE Chinese_Taiwan_Stroke_BIN
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].View_TB_WKF_TASK ON View_TB_WKF_EXTERNAL_TASK.DOC_NBR=View_TB_WKF_TASK.DOC_NBR
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].View_TB_WKF_TASK_NODE  ON View_TB_WKF_TASK_NODE.TASK_ID=View_TB_WKF_TASK.TASK_ID AND NODE_STATUS='1' AND ISNULL(SIGN_STATUS,'')=''
                                    LEFT JOIN [192.168.1.223].[UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=View_TB_WKF_TASK_NODE.ORIGINAL_SIGNER
                                    LEFT JOIN [TK].dbo.CMSME ON ME001=TA004
                                    WHERE TA007='N' 
                                    UNION ALL
                                    SELECT DISTINCT '請購變更單' AS '單別',ME002 AS '部門', [PURTATBCHAGE].[TA001],[PURTATBCHAGE].[TA002],[VERSIONS],'UOF','',View_TB_WKF_EXTERNAL_TASK.DOC_NBR,TB_EB_USER.NAME
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].View_TB_WKF_EXTERNAL_TASK ON View_TB_WKF_EXTERNAL_TASK.EXTERNAL_FORM_NBR LIKE TA001+TA002+CONVERT(NVARCHAR,[VERSIONS])+'%' COLLATE Chinese_Taiwan_Stroke_BIN
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].View_TB_WKF_TASK ON View_TB_WKF_EXTERNAL_TASK.DOC_NBR=View_TB_WKF_TASK.DOC_NBR
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].View_TB_WKF_TASK_NODE  ON View_TB_WKF_TASK_NODE.TASK_ID=View_TB_WKF_TASK.TASK_ID AND NODE_STATUS='1' AND ISNULL(SIGN_STATUS,'')=''
                                    LEFT JOIN [192.168.1.223].[UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=View_TB_WKF_TASK_NODE.ORIGINAL_SIGNER
                                    LEFT JOIN [TK].dbo.PURTA ON PURTA.TA001=[PURTATBCHAGE].TA001 AND PURTA.TA002=[PURTATBCHAGE].TA002
                                    LEFT JOIN [TK].dbo.CMSME ON ME001=PURTA.TA004  
                                    WHERE ISNULL(View_TB_WKF_EXTERNAL_TASK.DOC_NBR,'')<>''
                                    ) AS TEMP
                                    GROUP BY 部門,單別,TC001,TC002,SERNO,UDF01,UDF02,DOC_NBR
                                    ORDER BY 部門,單別,TC001,TC002
    
                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DSPURCHECK.Clear();
                adapter.Fill(DSPURCHECK, "DSPURCHECK");
                sqlConn.Close();
           


                if (DSPURCHECK.Tables["DSPURCHECK"].Rows.Count > 0)
                {
                    return DSPURCHECK;
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }


        }
        /// <summary>
        /// 昨天該到貨的採購單，但沒有進貨明細數量
        /// </summary>
        /// <returns></returns>
        public DataSet ERPPURTDCHECK()
        {
            DataSet DSERPPURTDCHECK = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                DateTime SDAYS = DateTime.Now.AddDays(-31);
                DateTime EDAYS = DateTime.Now.AddDays(-1);

                sbSql.AppendFormat(@"  
                                    SELECT TD012,TC004,MA002,TC001,TC002,TD003,TD004,TD005,TD008,TD009,ISNULL(SUMTH007,0) SUMTH007
                                    FROM [TK].dbo.PURMA,[TK].dbo.PURTC,[TK].dbo.PURTD
                                    LEFT JOIN (SELECT SUM(TH007) SUMTH007,TH011,TH012,TH013 FROM [TK].dbo.PURTH GROUP BY TH011,TH012,TH013) AS TEMP  ON TH011=TD001 AND TH012=TD002 AND TH013=TD003
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TD018='Y'
                                    AND MA001=TC004
                                    AND (TD008>ISNULL(SUMTH007,0))
                                    AND TD012>='{0}' AND TD012<='{1}'
    
                                    ORDER BY TD012,TC004,TC001,TC002,TD003

                                   ", SDAYS.ToString("yyyyMMdd"), EDAYS.ToString("yyyyMMdd"));

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DSERPPURTDCHECK.Clear();
                adapter.Fill(DSERPPURTDCHECK, "DSERPPURTDCHECK");
                sqlConn.Close();



                if (DSERPPURTDCHECK.Tables["DSERPPURTDCHECK"].Rows.Count > 0)
                {
                    return DSERPPURTDCHECK;
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }


        }

        public DataSet TKPUR_PURTATBCHAGE_DCHECK()
        {
            DataSet DSTKPUR_PURTATBCHAGE = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                DateTime SDAYS = DateTime.Now.AddDays(-14);
           

                sbSql.AppendFormat(@"  
                                 
                                SELECT 
                                [VERSIONS] AS '變更版號'
                                ,[TA001] AS '請購單'
                                ,[TA002] AS '請購單號'
                                ,[TA006] AS '單頭備註'
                                ,[TB003] AS '請購序號'
                                ,[TB004] AS '品號'
                                ,[TB005] AS '品名'
                                ,[TB009] AS '請購教量'
                                ,[TB011] AS '需求日'
                                ,[TB012] AS '單身備註'
                                ,[CHANGEDATES] AS '變更日期'
                                ,CONVERT(NVARCHAR,[VERSIONS])+[TA001]+[TA002]+[TB003]
                                FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                WHERE [TB011]>='{0}'
                                AND CONVERT(NVARCHAR,[VERSIONS])+[TA001]+[TA002]+[TB003] NOT IN (SELECT UDF01 FROM [TK].dbo.PURTF WHERE ISNULL(UDF01,'')<>'')

                                AND TA001+TA002 NOT IN (SELECT TA001+TA002 FROM [TKPUR].[dbo].[PURTATBSTOP])

                                ORDER BY [TB011]

                                   ", SDAYS.ToString("yyyyMMdd"));

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DSTKPUR_PURTATBCHAGE.Clear();
                adapter.Fill(DSTKPUR_PURTATBCHAGE, "DSTKPUR_PURTATBCHAGE");
                sqlConn.Close();



                if (DSTKPUR_PURTATBCHAGE.Tables["DSTKPUR_PURTATBCHAGE"].Rows.Count > 0)
                {
                    return DSTKPUR_PURTATBCHAGE;
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }


        }

        public DataSet FRIN_DS_PURTB_NOTIN_PURTD()
        {
            DataSet DS_PURTB_NOTIN_PURTD = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                DateTime SDAYS = DateTime.Now.AddDays(0);


                sbSql.AppendFormat(@"  
                                    SELECT MA002 AS '廠商',TB011 AS '需求日',TB001 AS '請購單別',TB002 AS '請購單號',TB003 AS '請購序號',TB004 AS '品號',TB005 AS '品名',TB006 AS '規格',TB008 AS '庫別',TB009 AS '請購數量',TB007  AS '單位' ,TB039 AS '是否採購'
                                    FROM [TK].dbo.PURTA,[TK].dbo.PURTB
                                    LEFT JOIN [TK].dbo.PURMA ON MA001=TB010
                                    WHERE TA001=TB001 AND TA002=TB002 
                                    AND TB001+TB002+TB003 NOT  IN (SELECT TD026+TD027+TD028 FROM [TK].dbo.PURTD WHERE ISNULL(TD026+TD027,'')<>'')
                                    AND  TA007 IN ('Y','N')
                                    AND  TB039='N'
                                    AND  TB025 NOT IN ('V')
                                    AND  TB011>='{0}'
                                    ORDER BY MA002,TB011
                               

                                   ", SDAYS.ToString("yyyyMMdd"));

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_PURTB_NOTIN_PURTD.Clear();
                adapter.Fill(DS_PURTB_NOTIN_PURTD, "DS_PURTB_NOTIN_PURTD");
                sqlConn.Close();



                if (DS_PURTB_NOTIN_PURTD.Tables["DS_PURTB_NOTIN_PURTD"].Rows.Count > 0)
                {
                    return DS_PURTB_NOTIN_PURTD;
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }
        }



        /// <summary>
        /// 找出要寄送給誰
        /// </summary>
        public DataSet FINDPURCHECKMAILTO(string SENDTO)
        {
            //SENDTO = "PURCHECK";

            DataSet FINDSENDMAILTO = new DataSet();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                sbSql.AppendFormat(@"  
                                    SELECT  [ID]
                                    ,[SENDTO]
                                    ,[MAIL]
                                    ,[NAME]
                                    FROM [TKMQ].[dbo].[MQSENDMAIL]
                                    WHERE [SENDTO]='{0}'
                                    ", SENDTO);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                FINDSENDMAILTO.Clear();

                adapter.Fill(FINDSENDMAILTO, "FINDSENDMAILTO");
                sqlConn.Close();


                if (FINDSENDMAILTO.Tables["FINDSENDMAILTO"].Rows.Count >= 1)
                {
                    return FINDSENDMAILTO;
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }
        }
        /// <summary>
        /// 實際寄出MEAIL，採購人員，ERP未核單的單別、單號
        /// </summary>
        public void SENDEMAILERPPURCHECK(StringBuilder Subject, StringBuilder Body)
        {
            DataSet DSFINDPURCHECKMAILTO = FINDPURCHECKMAILTO("PURCHECK");

            try
            {
                if (DSFINDPURCHECKMAILTO.Tables[0].Rows.Count>0)
                {
                    foreach(DataRow DR in DSFINDPURCHECKMAILTO.Tables[0].Rows)
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
                        MyMail.IsBodyHtml = true; //是否使用html格式

                        //加上附圖
                        //string path = System.Environment.CurrentDirectory + @"/Images/emaillogo.jpg";
                        //MyMail.AlternateViews.Add(GetEmbeddedImage(path, Body));

                        System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
                        MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);

                        


                        try
                        {
                            MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail
                                                    //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                            MySMTP.Send(MyMail);

                            MyMail.Dispose(); //釋放資源


                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("有錯誤");

                            //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                            //ex.ToString();
                        }
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
        //測試來客記錄的網站是否正常
        public void WEBTEST()
        {
            DataTable DTWEBLINKS = SEARCHLINKS();

            foreach(DataRow DR in DTWEBLINKS.Rows)
            {
                if (CheckUrlVisit(DR["WEBLINKS"].ToString())!=true)
                {
                    
                    //MessageBox.Show(DR["COMMENTS"].ToString() + " " + DR["WEBLINKS"].ToString() + ":" + CheckUrlVisit(DR["WEBLINKS"].ToString()));
                }

                //MessageBox.Show(DR["COMMENTS"].ToString()+ " "+DR["WEBLINKS"].ToString() + ":" + CheckUrlVisit(DR["WEBLINKS"].ToString()));
            }

            //string[] links = { "http://192.168.1.101:8900/index.html", "http://192.168.1.101:8900/index1.html" };
            //foreach (string link in links)
            //{
            //    MessageBox.Show(link.ToString()+":"+ CheckUrlVisit(link));
            //}
        }
        public bool CheckUrlVisit(string url)
        {
            try
            {
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
                if (resp.StatusCode == HttpStatusCode.OK)
                {
                    resp.Close();
                    return true;
                }
            }
            catch (WebException webex)
            {
                return false;
            }
            return false;
        }

        public DataTable SEARCHLINKS()
        {
            DataSet DS = new DataSet();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();



                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [WEBLINKS]
                                    ,[COMMENTS]
                                    FROM [TKIT].[dbo].[WEBLINKS]
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS.Clear();

                adapter.Fill(DS, "DS");
                sqlConn.Close();


                if (DS.Tables["DS"].Rows.Count >= 1)
                {
                    return DS.Tables[0];
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }
        }
        /// <summary>
        /// 資訊的每日檢查
        /// </summary>
        public void PREPAREITCHECK()
        {
            DataTable DTWEBLINKS = SEARCHLINKS();

            string ISCHECK = "Y";         

            try
            {
                StringBuilder SUBJEST = new StringBuilder();
                StringBuilder BODY = new StringBuilder();

                ////加上附圖
                //string path = System.Environment.CurrentDirectory+@"/Images/emaillogo.jpg";
                //LinkedResource res = new LinkedResource(path);
                //res.ContentId = Guid.NewGuid().ToString();

                SUBJEST.Clear();
                BODY.Clear();


                SUBJEST.AppendFormat(@"系統通知-資訊每日檢查 ，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                //ERP 採購相關單別、單號未核準的明細
                //
                BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                    + "<br>" + "資訊每日檢查"
                    + " <br>"
                    );

                foreach (DataRow DR in DTWEBLINKS.Rows)
                {
                    if (CheckUrlVisit(DR["WEBLINKS"].ToString()) != true)
                    {                      
                        BODY.AppendFormat(" <br>"
                                 + "{0} 此網站不通，請檢查網站狀況"
                                   + " <br>"
                                  , DR["COMMENTS"].ToString() + " " + DR["WEBLINKS"].ToString());

                        ISCHECK = "N";
                    }
                    else
                    {
                        BODY.AppendFormat(" <br>"
                                 + "{0} 此網站正常"
                                 +" <br>"
                                  , DR["COMMENTS"].ToString() + " " + DR["WEBLINKS"].ToString());
                    }


                }
               


               BODY.AppendFormat(" "
                             + "<br>" + "謝謝"

                             + "</span><br>");


                if(ISCHECK.Equals("N"))
                {
                    SUBJEST.AppendFormat(@" 有異常");
                }
                else
                {
                    SUBJEST.AppendFormat(@" ");
                }

                SENDEMAILITCHECK(SUBJEST, BODY);

            }
            catch
            {

            }
            finally
            {

            }
        }
        public void SENDEMAILITCHECK(StringBuilder Subject, StringBuilder Body)
        {
            DataSet DSFINDITCHECKMAILTO = FINDPURCHECKMAILTO("IT");

            try
            {
                if (DSFINDITCHECKMAILTO.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow DR in DSFINDITCHECKMAILTO.Tables[0].Rows)
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
                        MyMail.IsBodyHtml = true; //是否使用html格式

                        //加上附圖
                        //string path = System.Environment.CurrentDirectory + @"/Images/emaillogo.jpg";
                        //MyMail.AlternateViews.Add(GetEmbeddedImage(path, Body));

                        System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
                        MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);




                        try
                        {
                            MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail
                                                                  //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                            MySMTP.Send(MyMail);

                            MyMail.Dispose(); //釋放資源


                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("有錯誤");

                            //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                            //ex.ToString();
                        }
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

        public void PREPAREPROOFREAD()
        {
            DataSet DSPROOFREAD = UOFPROOFREAD();
           


            try
            {
                StringBuilder SUBJEST = new StringBuilder();
                StringBuilder BODY = new StringBuilder();

                ////加上附圖
                //string path = System.Environment.CurrentDirectory+@"/Images/emaillogo.jpg";
                //LinkedResource res = new LinkedResource(path);
                //res.ContentId = Guid.NewGuid().ToString();

                SUBJEST.Clear();
                BODY.Clear();


                SUBJEST.AppendFormat(@"系統通知-老楊食品-每日-校稿未完成的項目及交辨人回覆狀況 ，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                //ERP 採購相關單別、單號未核準的明細
                //
                BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                    + "<br>" + "校稿未完成的項目及交辨人回覆狀況 明細如下"

                    );


                if (DSPROOFREAD.Tables[0].Rows.Count > 0)
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                    BODY.AppendFormat(@"<table> ");
                    BODY.AppendFormat(@"<tr >");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=10% "">交辨開始時間</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=40% "">交辨項目</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=40% "">交辨回覆</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=10% "">被交辨人</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=10% "">交辨狀態</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=10% "">回覆時間</th>");

                    BODY.AppendFormat(@"</tr> ");

                    foreach (DataRow DR in DSPROOFREAD.Tables[0].Rows)
                    {

                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["交辨開始時間"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體'  width=40% "">" + DR["交辨項目"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體'  width=30% "">" + DR["交辨回覆"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["被交辨人"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["交辨狀態"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["回覆時間"].ToString() + "</td>");


                        BODY.AppendFormat(@"</tr> ");

                        //BODY.AppendFormat("<span></span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br> " + "品名     " + DR["TD005"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購數量 " + DR["TD008"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購單位 " + DR["TD009"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>");
                    }
                    BODY.AppendFormat(@"</table> ");
                }
                else
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "無資料");
                }

               

                BODY.AppendFormat(" "
                             + "<br>" + "謝謝"

                             + "</span><br>");



                SENDEMAILUOFPROOFEAD(SUBJEST, BODY);

            }
            catch
            {

            }
            finally
            {

            }
        }

        /// <summary>
        /// 準備寄給採購人員，ERP未核單的單別、單號
        /// </summary>
        public DataSet UOFPROOFREAD()
        {
            DataSet DSPROOFREAD = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                   SELECT CONVERT(nvarchar,TB_EIP_SCH_WORK.CREATE_TIME,111) AS '交辨開始時間'
                                    ,TB_EIP_SCH_WORK.SUBJECT AS '交辨項目'
                                    ,(ISNULL(TB_EIP_SCH_WORK.PROCEEDING_DESC,'')+ISNULL(TB_EIP_SCH_WORK.COMPLETE_DESC,''))  AS '交辨回覆'
                                    ,TB_EB_USER.NAME AS '被交辨人'
                                    ,(CASE  WHEN TB_EIP_SCH_WORK.WORK_STATE='Completed' THEN '審稿完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Audit' THEN '交辨完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Proceeding' THEN '處理中' WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN '未開始' END) AS '交辨狀態'
                                    ,(CASE WHEN ISNULL(TB_EIP_SCH_WORK.COMPLETE_TIME,'')<>'' THEN CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,24),1,8) ELSE CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,24),1,8) END)  AS '回覆時間'
                                    ,TB_EB_USER.ACCOUNT AS '被交辨人工號'
                                    FROM [UOF].dbo.TB_EIP_SCH_DEVOLVE
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_DEVOLVE_EXAMINE_LOG ON TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_WORK ON TB_EIP_SCH_WORK.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=TB_EIP_SCH_WORK.EXECUTE_USER
                                    WHERE 1=1
                                    AND TB_EIP_SCH_WORK.SUBJECT  LIKE '%校稿%'
                                    AND ISNULL(TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.STATUS,'') NOT IN ('Approve')
                                    AND TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID NOT IN (SELECT [DEVOLVE_GUID]  FROM [UOF].[dbo].[Z_TB_EIP_SCH_DEVOLVE_IGNORES])
                                    ORDER BY TB_EIP_SCH_DEVOLVE.CREATE_TIME
    
                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DSPROOFREAD.Clear();
                adapter.Fill(DSPROOFREAD, "DSPROOFREAD");
                sqlConn.Close();



                if (DSPROOFREAD.Tables["DSPROOFREAD"].Rows.Count > 0)
                {
                    return DSPROOFREAD;
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }


        }

        /// <summary>
        /// 實際寄出給校稿交辨人
        /// </summary>
        public void SENDEMAILUOFPROOFEAD(StringBuilder Subject, StringBuilder Body)
        {
            DataSet UOFPROOFEAD = FINDPURCHECKMAILTO("UOFPROOFEAD");     

            try
            {
                if (UOFPROOFEAD.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow DR in UOFPROOFEAD.Tables[0].Rows)
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
                        MyMail.IsBodyHtml = true; //是否使用html格式

                        //加上附圖
                        //string path = System.Environment.CurrentDirectory + @"/Images/emaillogo.jpg";
                        //MyMail.AlternateViews.Add(GetEmbeddedImage(path, Body));

                        System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
                        MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);




                        try
                        {
                            MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail
                                                                  //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                            MySMTP.Send(MyMail);

                            MyMail.Dispose(); //釋放資源


                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("有錯誤");

                            //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                            //ex.ToString();
                        }
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

        public void PREPARE_TB_EIP_PRIV_MESS_DIRECTOR()
        {
            DataTable DTFIND_USER_GUID = FIND_USER_GUID_DIRECTOR();
            string MESS = null; 
            

            if(DTFIND_USER_GUID.Rows.Count>0)
            {
                foreach (DataRow DR in DTFIND_USER_GUID.Rows)
                {
                    MESS = FIND_SUBJECT_DIRECTOR(DR["DIRECTOR"].ToString());
                    ADD_TB_EIP_PRIV_MESS_DIRECTOR(DR["DIRECTOR"].ToString(), MESS);
                }
            }

            //訊息給交辨人
            //ADD_TB_EIP_PRIV_MESS_DIRECTOR("b6f50a95-17ec-47f2-b842-4ad12512b431", MESS);
        }

        public DataTable FIND_USER_GUID_DIRECTOR()
        {
            DataSet DSPROOFREAD = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT DIRECTOR
                                    FROM 
                                    (
                                    SELECT CONVERT(nvarchar,TB_EIP_SCH_WORK.CREATE_TIME,111) AS '交辨開始時間'
                                    ,TB_EIP_SCH_DEVOLVE.SUBJECT AS '校稿區內容'
                                    ,TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID AS 'DEVOLVE_GUID'
                                    ,TB_EIP_SCH_WORK.SUBJECT AS '交辨項目'
                                    ,TB_EIP_SCH_WORK.EXECUTE_USER AS '交辨'
                                    ,TB_EIP_SCH_WORK.WORK_STATE AS 'WORK_STATE'
                                    ,(ISNULL(TB_EIP_SCH_WORK.PROCEEDING_DESC,'')+ISNULL(TB_EIP_SCH_WORK.COMPLETE_DESC,''))  AS '交辨回覆'
                                    ,TB_EB_USER.NAME AS '被交辨人'
                                    ,(CASE  WHEN TB_EIP_SCH_WORK.WORK_STATE='Completed' THEN '審稿完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Audit' THEN '交辨完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Proceeding' THEN '處理中' WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN '未開始' END) AS '交辨狀態'
                                    ,(CASE WHEN ISNULL(TB_EIP_SCH_WORK.COMPLETE_TIME,'')<>'' THEN CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,24),1,8) ELSE CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,24),1,8) END)  AS '回覆時間'
                                    ,TB_EB_USER.ACCOUNT
                                    ,TB_EB_USER.USER_GUID
                                    ,TB_EIP_SCH_DEVOLVE.DIRECTOR
                                    ,USER2.NAME AS '交辨人'
                                    FROM [UOF].dbo.TB_EIP_SCH_DEVOLVE
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_DEVOLVE_EXAMINE_LOG ON TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_WORK ON TB_EIP_SCH_WORK.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=TB_EIP_SCH_WORK.EXECUTE_USER
                                    LEFT JOIN [UOF].dbo.TB_EB_USER USER2 ON USER2.USER_GUID=TB_EIP_SCH_DEVOLVE.DIRECTOR

                                    WHERE 1=1
                                    AND TB_EIP_SCH_WORK.SUBJECT  LIKE '%校稿%'
                                    AND ISNULL(TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.STATUS,'') NOT IN ('Approve')
                                    AND TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID NOT IN (SELECT [DEVOLVE_GUID]  FROM [UOF].[dbo].[Z_TB_EIP_SCH_DEVOLVE_IGNORES])
                                    AND TB_EIP_SCH_WORK.WORK_STATE  NOT IN ('Audit')

                                    ) AS TEMP
                                    GROUP BY DIRECTOR

                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DSPROOFREAD.Clear();
                adapter.Fill(DSPROOFREAD, "DSPROOFREAD");
                sqlConn.Close();



                if (DSPROOFREAD.Tables["DSPROOFREAD"].Rows.Count > 0)
                {
                    return DSPROOFREAD.Tables["DSPROOFREAD"];
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }

        }

        public string FIND_SUBJECT_DIRECTOR(string USER_GUID)
        {
            StringBuilder MESS = new StringBuilder();
            DataSet DSPROOFREAD = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"
                                    
                                    SELECT CONVERT(nvarchar,TB_EIP_SCH_WORK.CREATE_TIME,111) AS '交辨開始時間'
                                    ,TB_EIP_SCH_DEVOLVE.SUBJECT AS '校稿區內容'
                                    ,TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID AS 'DEVOLVE_GUID'
                                    ,TB_EIP_SCH_WORK.SUBJECT AS '交辨項目'
                                    ,TB_EIP_SCH_WORK.EXECUTE_USER AS '交辨'
                                    ,TB_EIP_SCH_WORK.WORK_STATE AS 'WORK_STATE'
                                    ,(ISNULL(TB_EIP_SCH_WORK.PROCEEDING_DESC,'')+ISNULL(TB_EIP_SCH_WORK.COMPLETE_DESC,''))  AS '交辨回覆'
                                    ,TB_EB_USER.NAME AS '被交辨人'
                                    ,(CASE  WHEN TB_EIP_SCH_WORK.WORK_STATE='Completed' THEN '審稿完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Audit' THEN '交辨完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Proceeding' THEN '處理中' WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN '未開始' END) AS '交辨狀態'
                                    ,(CASE WHEN ISNULL(TB_EIP_SCH_WORK.COMPLETE_TIME,'')<>'' THEN CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,24),1,8) ELSE CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,24),1,8) END)  AS '回覆時間'
                                    ,TB_EB_USER.ACCOUNT
                                    ,TB_EB_USER.USER_GUID
                                    ,CONVERT(nvarchar,TB_EIP_SCH_DEVOLVE.END_TIME,111) AS '交辨預計結案日'
                                    ,DATEDIFF(day, TB_EIP_SCH_DEVOLVE.END_TIME, GETDATE()) AS '逾期天數' 
                                    ,TB_EIP_SCH_DEVOLVE.DIRECTOR
                                    ,USER2.NAME AS '交辨人'

                                    FROM [UOF].dbo.TB_EIP_SCH_DEVOLVE
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_DEVOLVE_EXAMINE_LOG ON TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_WORK ON TB_EIP_SCH_WORK.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=TB_EIP_SCH_WORK.EXECUTE_USER
                                    LEFT JOIN [UOF].dbo.TB_EB_USER USER2 ON USER2.USER_GUID=TB_EIP_SCH_DEVOLVE.DIRECTOR
                                    
                                    WHERE 1=1
                                    AND TB_EIP_SCH_WORK.SUBJECT  LIKE '%校稿%'
                                    AND ISNULL(TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.STATUS,'') NOT IN ('Approve')
                                    AND TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID NOT IN (SELECT [DEVOLVE_GUID]  FROM [UOF].[dbo].[Z_TB_EIP_SCH_DEVOLVE_IGNORES])
                                    AND TB_EIP_SCH_WORK.WORK_STATE  NOT IN ('Audit')

                                    AND TB_EIP_SCH_DEVOLVE.DIRECTOR='{0}'

                                   ", USER_GUID);

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DSPROOFREAD.Clear();
                adapter.Fill(DSPROOFREAD, "DSPROOFREAD");
                sqlConn.Close();



                if (DSPROOFREAD.Tables["DSPROOFREAD"].Rows.Count > 0)
                {
                    MESS.AppendFormat(@"<table> 
                                        <tr>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">交辨人</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">交辨項目</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">交辨開始時間</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">交辨預計結案日</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">交辨狀態</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">被交辨人</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">交辨回覆</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">回覆時間</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">逾期天數</td>

                                        </tr>
                                        ");

                    foreach (DataRow DR in DSPROOFREAD.Tables["DSPROOFREAD"].Rows)
                    {
                        MESS.AppendFormat(@"<tr>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["交辨人"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["交辨項目"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["交辨開始時間"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["交辨預計結案日"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["交辨狀態"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["被交辨人"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["交辨回覆"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["回覆時間"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["逾期天數"].ToString() + "</td>");

                        MESS.AppendFormat(@"</tr>");
                    }

                    MESS.AppendFormat(@"</table> ");

                    return MESS.ToString();
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }

            return MESS.ToString();
        }


        public void ADD_TB_EIP_PRIV_MESS_DIRECTOR(string USER_GUID,string MESS)
        {
            Guid NEW = Guid.NewGuid();
            string MESSAGE_GUID= NEW.ToString();
            string TOPIC= "系統通知-每日校稿的被交辨人未回覆項目" + DateTime.Now.ToString("yyyyMMdd");
            string MESSAGE_CONTENT= MESS;
            string MESSAGE_TO= USER_GUID;
            string MESSAGE_FROM= "916e213c-7b2e-46e3-8821-b7066378042b";
            string REPLY_MESSAGE_GUID=null;
            string CREATE_TIME= DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fffffffK");
            string READED_TIME = null;
            string REPLY_TIME = null;
            string FROM_DELETED="0";
            string TO_DELETED = "0";
            string FILE_GROUP_ID = null;
            string MASTER_GUID=NEW.ToString();
            string EVENT_ID = null;

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            connectionString = sqlConn.ConnectionString.ToString();

            StringBuilder queryString = new StringBuilder();




            queryString.AppendFormat(@"                     
                                        INSERT INTO [UOF].[dbo].[TB_EIP_PRIV_MESS]
                                        (
                                        [MESSAGE_GUID]
                                        ,[TOPIC]
                                        ,[MESSAGE_CONTENT]
                                        ,[MESSAGE_TO]
                                        ,[MESSAGE_FROM]
                                        ,[REPLY_MESSAGE_GUID]
                                        ,[CREATE_TIME]
                                        ,[READED_TIME]
                                        ,[REPLY_TIME]
                                        ,[FROM_DELETED]
                                        ,[TO_DELETED]
                                        ,[FILE_GROUP_ID]
                                        ,[MASTER_GUID]
                                        ,[EVENT_ID]
                                        )
                                        VALUES
                                        (
                                        '{0}'
                                        ,'{1}' 
                                        ,'{2}' 
                                        , '{3}'
                                        , '{4}'
                                        , NULL
                                        , '{5}'
                                        , NULL
                                        , NULL
                                        , '0'
                                        , '0'
                                        , ''
                                        , '{6}'
                                        , ''
                                        )
                                        ", MESSAGE_GUID
                                        , TOPIC
                                        , MESSAGE_CONTENT
                                        , MESSAGE_TO
                                        , MESSAGE_FROM
                                        , CREATE_TIME
                                        , MASTER_GUID

                                        );
                                       

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);                   

                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                }
            }
            catch
            {

            }
            finally
            {

            }


        }

        public void PREPARE_TB_EIP_PRIV_MESS()
        {
            DataTable DTFIND_USER_GUID = FIND_USER_GUID();
            string MESS = null;


            if (DTFIND_USER_GUID.Rows.Count > 0)
            {
                foreach (DataRow DR in DTFIND_USER_GUID.Rows)
                {
                    MESS = FIND_SUBJECT(DR["USER_GUID"].ToString());
                    ADD_TB_EIP_PRIV_MESS(DR["USER_GUID"].ToString(), MESS);
                }
            }

            //訊息給被交辨人
            //ADD_TB_EIP_PRIV_MESS("b6f50a95-17ec-47f2-b842-4ad12512b431", MESS);
        }

        public DataTable FIND_USER_GUID()
        {
            DataSet DSPROOFREAD = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT USER_GUID
                                    FROM 
                                    (
                                    SELECT CONVERT(nvarchar,TB_EIP_SCH_WORK.CREATE_TIME,111) AS '交辨開始時間'
                                    ,TB_EIP_SCH_DEVOLVE.SUBJECT AS '校稿區內容'
                                    ,TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID AS 'DEVOLVE_GUID'
                                    ,TB_EIP_SCH_WORK.SUBJECT AS '交辨項目'
                                    ,TB_EIP_SCH_WORK.EXECUTE_USER AS '交辨'
                                    ,TB_EIP_SCH_WORK.WORK_STATE AS 'WORK_STATE'
                                    ,(ISNULL(TB_EIP_SCH_WORK.PROCEEDING_DESC,'')+ISNULL(TB_EIP_SCH_WORK.COMPLETE_DESC,''))  AS '交辨回覆'
                                    ,TB_EB_USER.NAME AS '被交辨人'
                                    ,(CASE  WHEN TB_EIP_SCH_WORK.WORK_STATE='Completed' THEN '審稿完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Audit' THEN '交辨完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Proceeding' THEN '處理中' WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN '未開始' END) AS '交辨狀態'
                                    ,(CASE WHEN ISNULL(TB_EIP_SCH_WORK.COMPLETE_TIME,'')<>'' THEN CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,24),1,8) ELSE CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,24),1,8) END)  AS '回覆時間'
                                    ,TB_EB_USER.ACCOUNT
                                    ,TB_EB_USER.USER_GUID
                                    ,TB_EIP_SCH_DEVOLVE.DIRECTOR
                                    ,USER2.NAME AS '交辨人'
                                    FROM [UOF].dbo.TB_EIP_SCH_DEVOLVE
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_DEVOLVE_EXAMINE_LOG ON TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_WORK ON TB_EIP_SCH_WORK.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=TB_EIP_SCH_WORK.EXECUTE_USER
                                    LEFT JOIN [UOF].dbo.TB_EB_USER USER2 ON USER2.USER_GUID=TB_EIP_SCH_DEVOLVE.DIRECTOR

                                    WHERE 1=1
                                    AND TB_EIP_SCH_WORK.SUBJECT  LIKE '%校稿%'
                                    AND ISNULL(TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.STATUS,'') NOT IN ('Approve')
                                    AND TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID NOT IN (SELECT [DEVOLVE_GUID]  FROM [UOF].[dbo].[Z_TB_EIP_SCH_DEVOLVE_IGNORES])
                                    AND TB_EIP_SCH_WORK.WORK_STATE  NOT IN ('Audit')

                                    ) AS TEMP
                                    GROUP BY USER_GUID

                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DSPROOFREAD.Clear();
                adapter.Fill(DSPROOFREAD, "DSPROOFREAD");
                sqlConn.Close();



                if (DSPROOFREAD.Tables["DSPROOFREAD"].Rows.Count > 0)
                {
                    return DSPROOFREAD.Tables["DSPROOFREAD"];
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }

        }

        public string FIND_SUBJECT(string USER_GUID)
        {
            StringBuilder MESS = new StringBuilder();
            DataSet DSPROOFREAD = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"
                                    
                                    SELECT CONVERT(nvarchar,TB_EIP_SCH_WORK.CREATE_TIME,111) AS '交辨開始時間'
                                    ,TB_EIP_SCH_DEVOLVE.SUBJECT AS '校稿區內容'
                                    ,TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID AS 'DEVOLVE_GUID'
                                    ,TB_EIP_SCH_WORK.SUBJECT AS '交辨項目'
                                    ,TB_EIP_SCH_WORK.EXECUTE_USER AS '交辨'
                                    ,TB_EIP_SCH_WORK.WORK_STATE AS 'WORK_STATE'
                                    ,(ISNULL(TB_EIP_SCH_WORK.PROCEEDING_DESC,'')+ISNULL(TB_EIP_SCH_WORK.COMPLETE_DESC,''))  AS '交辨回覆'
                                    ,TB_EB_USER.NAME AS '被交辨人'
                                    ,(CASE  WHEN TB_EIP_SCH_WORK.WORK_STATE='Completed' THEN '審稿完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Audit' THEN '交辨完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Proceeding' THEN '處理中' WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN '未開始' END) AS '交辨狀態'
                                    ,(CASE WHEN ISNULL(TB_EIP_SCH_WORK.COMPLETE_TIME,'')<>'' THEN CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,24),1,8) ELSE CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,24),1,8) END)  AS '回覆時間'
                                    ,TB_EB_USER.ACCOUNT
                                    ,TB_EB_USER.USER_GUID
                                    ,CONVERT(nvarchar,TB_EIP_SCH_DEVOLVE.END_TIME,111) AS '交辨預計結案日'
                                    ,DATEDIFF(day, TB_EIP_SCH_DEVOLVE.END_TIME, GETDATE()) AS '逾期天數' 
                                    ,TB_EIP_SCH_DEVOLVE.DIRECTOR
                                    ,USER2.NAME AS '交辨人'

                                    FROM [UOF].dbo.TB_EIP_SCH_DEVOLVE
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_DEVOLVE_EXAMINE_LOG ON TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_WORK ON TB_EIP_SCH_WORK.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=TB_EIP_SCH_WORK.EXECUTE_USER
                                    LEFT JOIN [UOF].dbo.TB_EB_USER USER2 ON USER2.USER_GUID=TB_EIP_SCH_DEVOLVE.DIRECTOR
                                    
                                    WHERE 1=1
                                    AND TB_EIP_SCH_WORK.SUBJECT  LIKE '%校稿%'
                                    AND ISNULL(TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.STATUS,'') NOT IN ('Approve')
                                    AND TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID NOT IN (SELECT [DEVOLVE_GUID]  FROM [UOF].[dbo].[Z_TB_EIP_SCH_DEVOLVE_IGNORES])
                                    AND TB_EIP_SCH_WORK.WORK_STATE  NOT IN ('Audit')

                                    AND TB_EB_USER.USER_GUID='{0}'

                                   ", USER_GUID);

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DSPROOFREAD.Clear();
                adapter.Fill(DSPROOFREAD, "DSPROOFREAD");
                sqlConn.Close();



                if (DSPROOFREAD.Tables["DSPROOFREAD"].Rows.Count > 0)
                {
                    MESS.AppendFormat(@"<table> 
                                        <tr>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">交辨人</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">交辨項目</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">交辨開始時間</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">交辨預計結案日</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">交辨狀態</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">被交辨人</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">交辨回覆</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">回覆時間</td>
                                        <td style=""border: 1px solid #999;font-size:12.0pt width=10% "">逾期天數</td>

                                        </tr>
                                        ");

                    foreach (DataRow DR in DSPROOFREAD.Tables["DSPROOFREAD"].Rows)
                    {
                        MESS.AppendFormat(@"<tr>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["交辨人"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["交辨項目"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["交辨開始時間"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["交辨預計結案日"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["交辨狀態"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["被交辨人"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["交辨回覆"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["回覆時間"].ToString() + "</td>");
                        MESS.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt "">" + DR["逾期天數"].ToString() + "</td>");

                        MESS.AppendFormat(@"</tr>");
                    }

                    MESS.AppendFormat(@"</table> ");

                    return MESS.ToString();
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }

            return MESS.ToString();
        }


        public void ADD_TB_EIP_PRIV_MESS(string USER_GUID, string MESS)
        {
            Guid NEW = Guid.NewGuid();
            string MESSAGE_GUID = NEW.ToString();
            string TOPIC = "系統通知-每日校稿的未回覆項目，請於3天內至交辨區回覆校稿" + DateTime.Now.ToString("yyyyMMdd");
            string MESSAGE_CONTENT = MESS;
            string MESSAGE_TO = USER_GUID;
            string MESSAGE_FROM = "916e213c-7b2e-46e3-8821-b7066378042b";
            string REPLY_MESSAGE_GUID = null;
            string CREATE_TIME = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fffffffK");
            string READED_TIME = null;
            string REPLY_TIME = null;
            string FROM_DELETED = "0";
            string TO_DELETED = "0";
            string FILE_GROUP_ID = null;
            string MASTER_GUID = NEW.ToString();
            string EVENT_ID = null;

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            connectionString = sqlConn.ConnectionString.ToString();

            StringBuilder queryString = new StringBuilder();




            queryString.AppendFormat(@"                     
                                        INSERT INTO [UOF].[dbo].[TB_EIP_PRIV_MESS]
                                        (
                                        [MESSAGE_GUID]
                                        ,[TOPIC]
                                        ,[MESSAGE_CONTENT]
                                        ,[MESSAGE_TO]
                                        ,[MESSAGE_FROM]
                                        ,[REPLY_MESSAGE_GUID]
                                        ,[CREATE_TIME]
                                        ,[READED_TIME]
                                        ,[REPLY_TIME]
                                        ,[FROM_DELETED]
                                        ,[TO_DELETED]
                                        ,[FILE_GROUP_ID]
                                        ,[MASTER_GUID]
                                        ,[EVENT_ID]
                                        )
                                        VALUES
                                        (
                                        '{0}'
                                        ,'{1}' 
                                        ,'{2}' 
                                        , '{3}'
                                        , '{4}'
                                        , NULL
                                        , '{5}'
                                        , NULL
                                        , NULL
                                        , '0'
                                        , '0'
                                        , ''
                                        , '{6}'
                                        , ''
                                        )
                                        ", MESSAGE_GUID
                                        , TOPIC
                                        , MESSAGE_CONTENT
                                        , MESSAGE_TO
                                        , MESSAGE_FROM
                                        , CREATE_TIME
                                        , MASTER_GUID

                                        );


            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);

                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                }
            }
            catch
            {

            }
            finally
            {

            }


        }

        public void PREPARE_UOF_TASK_TASK_APPLICATION()
        {
            DataTable DT_FIND_UOF_TASK_APPLICATION = FIND_UOF_TASK_APPLICATION();
            DataTable DT_FIND_UOF_TASK_APPLICATION_FORM = new DataTable();

            if (DT_FIND_UOF_TASK_APPLICATION != null && DT_FIND_UOF_TASK_APPLICATION.Rows.Count>=1)
            {
                foreach(DataRow DR in DT_FIND_UOF_TASK_APPLICATION.Rows)
                {
                    DT_FIND_UOF_TASK_APPLICATION_FORM = FIND_UOF_TASK_APPLICATION_FORM(DR["APPLICANT_NAME"].ToString());
                 
                    if(DT_FIND_UOF_TASK_APPLICATION_FORM!=null && DT_FIND_UOF_TASK_APPLICATION_FORM.Rows.Count>=1)
                    {
                        SEND_UOF_TASK_APPLICATION_FORM(DR["APPLICANT_NAME"].ToString(), DR["APPLICANT_EMAIL"].ToString(), DT_FIND_UOF_TASK_APPLICATION_FORM);
                    } 
                }
               
            }
        }


        public DataTable FIND_UOF_TASK_APPLICATION()
        {
            StringBuilder MESS = new StringBuilder();
            DataSet DS_FIND_UOF_TASK_APPLICATION = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                //AND APPLICANT_NAME='張健洲'
                sbSql.AppendFormat(@"
                                    SELECT APPLICANT_NAME, APPLICANT_EMAIL
                                    FROM 
                                    (
                                    SELECT
                                    usr2.NAME AS 'CURRENTNAME'
                                    ,[TB_EB_JOB_TITLE].TITLE_NAME AS 'CURRENTTITLENAME'
                                    ,[TB_EB_JOB_TITLE].RANK AS 'CURRENTRANK'
                                    ,(CASE WHEN  usr.IS_SUSPENDED = 1 THEN  usr.NAME + '(x)' WHEN  ISNULL(usr.ACCOUNT,'''') = '' THEN  'unknown user' ELSE usr.NAME END) AS APPLICANT_NAME
                                    ,usr.[EMAIL] AS 'APPLICANT_EMAIL'
                                    ,form.FORM_NAME
                                    ,DOC_NBR
                                    ,CONVERT(NVARCHAR,NODES.START_TIME,111) AS 'START_TIME'
                                    ,DATEDIFF(HOUR,START_TIME,GETDATE()) AS 'HRS'
                                    ,CONVERT(NVARCHAR,BEGIN_TIME,111) AS BEGIN_TIME
                                    ,task.TASK_ID
                                    ,END_TIME
                                    ,TASK_RESULT
                                    ,TASK_STATUS
                                    ,task.USER_GUID
                                    ,formVer.FORM_VERSION_ID
                                    ,formVer.FORM_ID
                                    ,CURRENT_SITE_ID
                                    ,MESSAGE_CONTENT
                                    ,LOCK_STATUS
                                    ,ISNULL(formVer.DISPLAY_TITLE,'') AS VERSION_TITLE
                                    ,ISNULL(task.JSON_DISPLAY,'') AS JSON_DISPLAY
                                    ,[NODES].SIGN_STATUS
                                    FROM dbo.TB_WKF_TASK task
                                    INNER JOIN dbo.TB_WKF_FORM_VERSION formVer ON task.FORM_VERSION_ID = formVer.FORM_VERSION_ID
                                    INNER JOIN dbo.TB_WKF_FORM form  ON  formVer.FORM_ID = form.FORM_ID 
                                    LEFT JOIN dbo.TB_EB_USER [usr]  ON task.USER_GUID = usr.USER_GUID
                                    LEFT JOIN dbo.TB_WKF_TASK_NODE [NODES] ON NODES.SITE_ID=task.CURRENT_SITE_ID 
                                    LEFT JOIN dbo.TB_EB_USER [usr2]  ON NODES.ORIGINAL_SIGNER = [usr2].USER_GUID
                                    LEFT JOIN dbo.[TB_EB_EMPL_DEP] ON [TB_EB_EMPL_DEP].USER_GUID=[usr2].USER_GUID
                                    LEFT JOIN dbo.[TB_EB_JOB_TITLE] ON [TB_EB_EMPL_DEP].TITLE_ID=[TB_EB_JOB_TITLE].TITLE_ID


                                    WHERE
                                    1=1  
                                    AND  TASK_STATUS NOT IN ('2')
                                    AND ISNULL([NODES].SIGN_STATUS,999)<>0
                                    AND DATEDIFF(HOUR,START_TIME,GETDATE())>=24
                                    )  AS TEMP 
                                    WHERE ISNULL(APPLICANT_EMAIL,'')<>''
                                    
                                    GROUP BY APPLICANT_NAME,APPLICANT_EMAIL
                                    ORDER BY APPLICANT_NAME,APPLICANT_EMAIL

                                   

                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_FIND_UOF_TASK_APPLICATION.Clear();
                adapter.Fill(DS_FIND_UOF_TASK_APPLICATION, "DS_FIND_UOF_TASK_APPLICATION");
                sqlConn.Close();



                if (DS_FIND_UOF_TASK_APPLICATION.Tables["DS_FIND_UOF_TASK_APPLICATION"].Rows.Count > 0)
                {

                    return DS_FIND_UOF_TASK_APPLICATION.Tables["DS_FIND_UOF_TASK_APPLICATION"];
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }
            
        }

        public DataTable FIND_UOF_TASK_APPLICATION_FORM(string APPLICANT_NAME)
        {

            StringBuilder MESS = new StringBuilder();
            DataSet DS_FIND_UOF_TASK_APPLICATION_FORM = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"
                                    
                                   SELECT
                                    APPLICANT_NAME,FORM_NAME,DOC_NBR,START_TIME,CURRENTNAME
                                    FROM
                                    (
                                    SELECT 
                                    usr2.NAME AS 'CURRENTNAME'
                                    ,[TB_EB_JOB_TITLE].TITLE_NAME AS 'CURRENTTITLENAME'
                                    ,[TB_EB_JOB_TITLE].RANK AS 'CURRENTRANK'
                                    ,(CASE WHEN  usr.IS_SUSPENDED = 1 THEN  usr.NAME + '(x)' WHEN  ISNULL(usr.ACCOUNT,'''') = '' THEN  'unknown user' ELSE usr.NAME END) AS APPLICANT_NAME
                                    ,usr.[EMAIL] AS 'APPLICANT_EMAIL'
                                    ,form.FORM_NAME
                                    ,DOC_NBR
                                    ,CONVERT(NVARCHAR,NODES.START_TIME,111) AS 'START_TIME'
                                    ,DATEDIFF(HOUR,START_TIME,GETDATE()) AS 'HRS'
                                    ,CONVERT(NVARCHAR,BEGIN_TIME,111) AS BEGIN_TIME
                                    ,task.TASK_ID
                                    ,END_TIME
                                    ,TASK_RESULT
                                    ,TASK_STATUS
                                    ,task.USER_GUID
                                    ,formVer.FORM_VERSION_ID
                                    ,formVer.FORM_ID
                                    ,CURRENT_SITE_ID
                                    ,MESSAGE_CONTENT
                                    ,LOCK_STATUS
                                    ,ISNULL(formVer.DISPLAY_TITLE,'') AS VERSION_TITLE
                                    ,ISNULL(task.JSON_DISPLAY,'') AS JSON_DISPLAY
                                    ,[NODES].SIGN_STATUS
                                    FROM dbo.TB_WKF_TASK task
                                    INNER JOIN dbo.TB_WKF_FORM_VERSION formVer ON task.FORM_VERSION_ID = formVer.FORM_VERSION_ID
                                    INNER JOIN dbo.TB_WKF_FORM form  ON  formVer.FORM_ID = form.FORM_ID 
                                    LEFT JOIN dbo.TB_EB_USER [usr]  ON task.USER_GUID = usr.USER_GUID
                                    LEFT JOIN dbo.TB_WKF_TASK_NODE [NODES] ON NODES.SITE_ID=task.CURRENT_SITE_ID 
                                    LEFT JOIN dbo.TB_EB_USER [usr2]  ON NODES.ORIGINAL_SIGNER = [usr2].USER_GUID
                                    LEFT JOIN dbo.[TB_EB_EMPL_DEP] ON [TB_EB_EMPL_DEP].USER_GUID=[usr2].USER_GUID
                                    LEFT JOIN dbo.[TB_EB_JOB_TITLE] ON [TB_EB_EMPL_DEP].TITLE_ID=[TB_EB_JOB_TITLE].TITLE_ID


                                    WHERE
                                    1=1  
                                    AND  TASK_STATUS NOT IN ('2')
                                    AND ISNULL([NODES].SIGN_STATUS,999)<>0
                                    AND DATEDIFF(HOUR,START_TIME,GETDATE())>=24

                                    ) AS TEMP
                                    WHERE 1=1
                                    AND  APPLICANT_NAME='{0}'
                                    GROUP BY APPLICANT_NAME,FORM_NAME,DOC_NBR,START_TIME,CURRENTNAME
                                    ORDER BY FORM_NAME,DOC_NBR

                                   

                                   ", APPLICANT_NAME);

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_FIND_UOF_TASK_APPLICATION_FORM.Clear();
                adapter.Fill(DS_FIND_UOF_TASK_APPLICATION_FORM, "DS_FIND_UOF_TASK_APPLICATION_FORM");
                sqlConn.Close();



                if (DS_FIND_UOF_TASK_APPLICATION_FORM.Tables["DS_FIND_UOF_TASK_APPLICATION_FORM"].Rows.Count > 0)
                {

                    return DS_FIND_UOF_TASK_APPLICATION_FORM.Tables["DS_FIND_UOF_TASK_APPLICATION_FORM"];
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }
        }

        public void SEND_UOF_TASK_APPLICATION_FORM(string APPLICANT_NAME,string APPLICANT_EMAIL, DataTable DT)
        {
            try
            {
                StringBuilder SUBJEST = new StringBuilder();
                StringBuilder BODY = new StringBuilder();

                ////加上附圖
                //string path = System.Environment.CurrentDirectory+@"/Images/emaillogo.jpg";
                //LinkedResource res = new LinkedResource(path);
                //res.ContentId = Guid.NewGuid().ToString();

                SUBJEST.Clear();
                BODY.Clear();


                SUBJEST.AppendFormat(@"系統通知-請查收-每日-UOF表單中，表單尚未核準的明細及目前表單簽核人員(超過24小時)，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                //ERP 採購相關單別、單號未核準的明細
                //
                BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                    + "<br>" + "系統通知-請查收-每日-UOF表單中，表單尚未核準的明細及目前表單簽核人員(超過24小時)，謝謝"
                    + " <br>"
                    );

          



                if (DT.Rows.Count > 0)
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                    BODY.AppendFormat(@"<table> ");
                    BODY.AppendFormat(@"<tr >");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">申請人員</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">申請表單</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">表單單號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">申請時間</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">目前簽核人員</th>");


                    BODY.AppendFormat(@"</tr> ");

                    foreach (DataRow DR in DT.Rows)
                    {

                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["APPLICANT_NAME"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["FORM_NAME"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["DOC_NBR"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["START_TIME"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["CURRENTNAME"].ToString() + "</td>");
                      
                        BODY.AppendFormat(@"</tr> ");

                        //BODY.AppendFormat("<span></span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br> " + "品名     " + DR["TD005"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購數量 " + DR["TD008"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購單位 " + DR["TD009"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>");
                    }
                    BODY.AppendFormat(@"</table> ");
                }


                try
                {
                    string MySMTPCONFIG = ConfigurationManager.AppSettings["MySMTP"];
                    string NAME = ConfigurationManager.AppSettings["NAME"];
                    string PW = ConfigurationManager.AppSettings["PW"];

                    System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
                    MyMail.From = new System.Net.Mail.MailAddress("tk290@tkfood.com.tw");

                    //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
                    //MyMail.Subject = "每日訂單-製令追踨表"+DateTime.Now.ToString("yyyy/MM/dd");
                    MyMail.Subject = SUBJEST.ToString();
                    //MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
                    MyMail.Body = BODY.ToString();
                    MyMail.IsBodyHtml = true; //是否使用html格式

                    //加上附圖
                    //string path = System.Environment.CurrentDirectory + @"/Images/emaillogo.jpg";
                    //MyMail.AlternateViews.Add(GetEmbeddedImage(path, Body));

                    System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
                    MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);




                    try
                    {
                        MyMail.To.Add(APPLICANT_EMAIL); //設定收件者Email，多筆mail
                                                              //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                        MySMTP.Send(MyMail);

                        MyMail.Dispose(); //釋放資源


                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("有錯誤");

                        //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                        //ex.ToString();
                    }
                }
                catch
                {

                }
                finally
                {

                }


            }
            catch
            {

            }
            finally
            {

            }
        }

        /// <summary>
        /// 找出昨天核準過的採購單，通知原請購人到貨了
        /// </summary>
        public void FIND_UOF_GRAFFAIRS_1005()
        {
            DataTable DTSEARCHUOF_GRAFFAIRS_1005 = SEARCHUOF_GRAFFAIRS_1005();

            if(DTSEARCHUOF_GRAFFAIRS_1005!=null && DTSEARCHUOF_GRAFFAIRS_1005.Rows.Count>=1)
            {
                foreach(DataRow DR in DTSEARCHUOF_GRAFFAIRS_1005.Rows)
                {
                    string USER_GUID = DR["USER_GUID"].ToString();
                    string EMAILTO = DR["EMAIL"].ToString();

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(DR["CURRENT_DOC"].ToString());

                    //XmlNode node = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='ID']");
                    string ID = "";
                    string GA002 = "";
                    string GA003 = "";
                    string GA005 = "";
                    string GA015 = "";
                    string GA999 = "";


                    try
                    {
                        ID = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='ID']").Attributes["fieldValue"].Value;
                        

                    }
                    catch { }
                    try
                    {
                        GA002 = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='GA002']").Attributes["fieldValue"].Value;
                       
                    }
                    catch { }
                    try
                    {
                        GA003 = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='GA003']").Attributes["fieldValue"].Value;
                       
                    }
                    catch { }
                    try
                    {
                        GA005 = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='GA005']").Attributes["fieldValue"].Value;
                        
                    }
                    catch { }
                    try
                    {
                        GA015 = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='GA015']").Attributes["fieldValue"].Value;
                        
                    }
                    catch { }
                    try
                    {
                        
                        GA999 = xmlDoc.SelectSingleNode($"/Form/FormFieldValue/FieldItem[@fieldId='GA999']").Attributes["fieldValue"].Value;
                    }
                    catch { }


                    string MESSAGES = GA003 + " 申請的請購單:"+ GA002 + "，物品:"+ GA005+"，已由"+GA999+" 在"+ GA015+"購買完成。";

                 
                    if(!string.IsNullOrEmpty(USER_GUID))
                    {
                        SEND_MESSAGE_UOF_GRAFFAIRS_1005(USER_GUID, MESSAGES);
                    }
                    if (!string.IsNullOrEmpty(EMAILTO))
                    {
                        SEND_EMAIL_UOF_GRAFFAIRS_1005(EMAILTO, MESSAGES, MESSAGES);
                    }

                       
                }
            }


        }
        //找出昨天，已核單完成的採購單-1005.雜項採購單
        public DataTable SEARCHUOF_GRAFFAIRS_1005()
        {
            StringBuilder MESS = new StringBuilder();
            DataSet DS_FIND_UOF_TASK_APPLICATION_FORM = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            string END_TIME = DateTime.Now.AddDays(-1).ToString("yyyyMMdd");

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                //   AND DOC_NBR = 'GA1005230100006'

                sbSql.AppendFormat(@"                                    
                                   SELECT CURRENT_DOC,* 
                                   
                                    FROM [UOF].[dbo].TB_WKF_TASK ,[UOF].[dbo].[TB_EB_USER],[UOF].dbo.TB_WKF_FORM,[UOF].dbo.TB_WKF_FORM_VERSION
                                    WHERE TB_WKF_TASK.USER_GUID=[TB_EB_USER].USER_GUID
                                    AND TB_WKF_TASK.FORM_VERSION_ID=TB_WKF_FORM_VERSION.FORM_VERSION_ID
                                    AND TB_WKF_FORM.FORM_ID=TB_WKF_FORM_VERSION.FORM_ID
                                    AND TB_WKF_FORM.FORM_NAME='1005.雜項採購單'
                                    AND TASK_RESULT='0' AND TASK_STATUS='2'
                                    AND CONVERT(NVARCHAR,END_TIME,112)='{0}'
                                 

                                   ", END_TIME);

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_FIND_UOF_TASK_APPLICATION_FORM.Clear();
                adapter.Fill(DS_FIND_UOF_TASK_APPLICATION_FORM, "DS_FIND_UOF_TASK_APPLICATION_FORM");
                sqlConn.Close();



                if (DS_FIND_UOF_TASK_APPLICATION_FORM.Tables["DS_FIND_UOF_TASK_APPLICATION_FORM"].Rows.Count > 0)
                {

                    return DS_FIND_UOF_TASK_APPLICATION_FORM.Tables["DS_FIND_UOF_TASK_APPLICATION_FORM"];
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }

        }
        /// <summary>
        /// 通知原請購人到貨了，用UOF訊息
        /// </summary>
        public void SEND_MESSAGE_UOF_GRAFFAIRS_1005(string USER_GUID,string MESS)
        {
            Guid NEW = Guid.NewGuid();
            string MESSAGE_GUID = NEW.ToString();
            string TOPIC = "系統通知 "+ MESS;
            string MESSAGE_CONTENT="系統通知 " + MESS;
            string MESSAGE_TO = USER_GUID;
            string MESSAGE_FROM = "916e213c-7b2e-46e3-8821-b7066378042b";
            string REPLY_MESSAGE_GUID = null;
            string CREATE_TIME = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fffffffK");
            string READED_TIME = null;
            string REPLY_TIME = null;
            string FROM_DELETED = "0";
            string TO_DELETED = "0";
            string FILE_GROUP_ID = null;
            string MASTER_GUID = NEW.ToString();
            string EVENT_ID = null;

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            connectionString = sqlConn.ConnectionString.ToString();

            StringBuilder queryString = new StringBuilder();




            queryString.AppendFormat(@"                     
                                        INSERT INTO [UOF].[dbo].[TB_EIP_PRIV_MESS]
                                        (
                                        [MESSAGE_GUID]
                                        ,[TOPIC]
                                        ,[MESSAGE_CONTENT]
                                        ,[MESSAGE_TO]
                                        ,[MESSAGE_FROM]
                                        ,[REPLY_MESSAGE_GUID]
                                        ,[CREATE_TIME]
                                        ,[READED_TIME]
                                        ,[REPLY_TIME]
                                        ,[FROM_DELETED]
                                        ,[TO_DELETED]
                                        ,[FILE_GROUP_ID]
                                        ,[MASTER_GUID]
                                        ,[EVENT_ID]
                                        )
                                        VALUES
                                        (
                                        '{0}'
                                        ,'{1}' 
                                        ,'{2}' 
                                        , '{3}'
                                        , '{4}'
                                        , NULL
                                        , '{5}'
                                        , NULL
                                        , NULL
                                        , '0'
                                        , '0'
                                        , ''
                                        , '{6}'
                                        , ''
                                        )
                                        ", MESSAGE_GUID
                                        , TOPIC
                                        , MESSAGE_CONTENT
                                        , MESSAGE_TO
                                        , MESSAGE_FROM
                                        , CREATE_TIME
                                        , MASTER_GUID

                                        );


            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);

                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                }
            }
            catch
            {

            }
            finally
            {

            }
        }
        /// <summary>
        /// 通知原請購人到貨了, 用EMAIL
        /// </summary>
        public void SEND_EMAIL_UOF_GRAFFAIRS_1005(string EMAILTO,string Subject,string Body)
        {
            try
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
                MyMail.IsBodyHtml = true; //是否使用html格式

                //加上附圖
                //string path = System.Environment.CurrentDirectory + @"/Images/emaillogo.jpg";
                //MyMail.AlternateViews.Add(GetEmbeddedImage(path, Body));

                System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
                MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);




                try
                {
                    MyMail.To.Add(EMAILTO); //設定收件者Email，多筆mail
                                                          //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                    MySMTP.Send(MyMail);

                    MyMail.Dispose(); //釋放資源


                }
                catch (Exception ex)
                {
                    MessageBox.Show("有錯誤");

                    //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                    //ex.ToString();
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
            //CAN"T USE SQL-CTE QUERY
            SETFILEPURTA();
            //SETFILEPURTA2();

            CLEAREXCEL();
            MessageBox.Show("OK");
        }
        private void button3_Click(object sender, EventArgs e)
        {
            HRAUTORUN();

            MessageBox.Show("OK");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SETPATH();
            SETFILEMOCTA();
            CLEAREXCEL();
            MessageBox.Show("OK");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SETPATH();
            SETFILEINVMOCTA();
            CLEAREXCEL();
            MessageBox.Show("OK");
        }
        private void button7_Click(object sender, EventArgs e)
        {
            SETPATH();
            SETFILEPURTB();
            CLEAREXCEL();
            MessageBox.Show("OK");
        }
        private void button8_Click(object sender, EventArgs e)
        {
            SETPATH();
            SETFILEMOCINVCHECK();
            CLEAREXCEL();
            MessageBox.Show("OK");
        }
        private void button9_Click(object sender, EventArgs e)
        {
            SETPATH();
            SETFILEMOCCOP();
            CLEAREXCEL();
            MessageBox.Show("OK");
        }
        private void button10_Click(object sender, EventArgs e)
        {
            SETPATH();
            SETFILEINVMC();
            CLEAREXCEL();
            MessageBox.Show("OK");
        }
        private void button11_Click(object sender, EventArgs e)
        {
            SETPATH();
            SETFILEPURTD();
            CLEAREXCEL();
            MessageBox.Show("OK");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            SETPATH();
            SETFILEMOCTARE();
            CLEAREXCEL();
            MessageBox.Show("OK");
        }
        private void button13_Click(object sender, EventArgs e)
        {

            SETPATH();
            SETFILELOTCHECK();
            
            CLEAREXCEL();

            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();
            //LOTCHECK
            SERACHMAILLOTCHECK();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"每日批號檢查表" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日批號檢查表，請查收 (批號錯誤時，要檢查「批號資料建立作業」內的有效日期、複檢日期是否也錯誤)" + Environment.NewLine + " ");
            SENDMAIL(SUBJEST, BODY, dsMAILLOTCHECK, pathFileLOTCHECK);
       

            MessageBox.Show("OK");
        }

        private void button15_Click(object sender, EventArgs e)
        {
            SETPATH();
            SETFILEMOCMANULINE();

            CLEAREXCEL();
            MessageBox.Show("OK");
        }
        private void button14_Click(object sender, EventArgs e)
        {
            ADDLOG(DateTime.Now,"TEST","EX");
        }
        private void button16_Click(object sender, EventArgs e)
        {
            PREPARESENDEMAILERPPURCHECK();
        }
        private void button17_Click(object sender, EventArgs e)
        {
            //WEBTEST();

            PREPAREITCHECK();
        }
        private void button18_Click(object sender, EventArgs e)
        {
            PREPAREPROOFREAD();
        }
        private void button19_Click(object sender, EventArgs e)
        {
            //通知各別的被交辨人
            PREPARE_TB_EIP_PRIV_MESS();

            //通知交辨人
            PREPARE_TB_EIP_PRIV_MESS_DIRECTOR();
        }
        private void button20_Click(object sender, EventArgs e)
        {
            //通知各表單申請人
            PREPARE_UOF_TASK_TASK_APPLICATION();
        }

        private void button21_Click(object sender, EventArgs e)
        {
            //通知原請購人，總務已完成採購
            FIND_UOF_GRAFFAIRS_1005();
        }
        #endregion


    }
}
