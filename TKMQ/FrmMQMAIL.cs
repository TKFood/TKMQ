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
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using NPOI.XSSF.UserModel;
using TKITDLL;
using System.Net.Http;
using System.Net;
using System.Xml;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Collections.Specialized;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using System.Net.Mime;

namespace TKMQ
{
    public partial class FrmMQMAIL : Form
    {
        int TIMEOUT_LIMITS = 240;
        private System.Timers.Timer timer;

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
        string path_File_NEWSLAES = null;
        string path_File_POSINV = null;
        string path_File_COPTCD = null;
        string pathFile_SALES_MONEYS = null;
        string pathFile_QC_CHECK = null;

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

            timer2.Enabled = true;
            timer2.Interval = 1000 * 60;
            //timer1.Interval = 1000 ;
            timer2.Start();

            // 在適當的地方設置 timer3 的屬性
            timer3.Enabled = true;
            timer3.Interval = (int)TimeSpan.FromDays(1).TotalMilliseconds; // 設置為一天的毫秒數
            timer3.Tick += new EventHandler(timer3_Tick);
            timer3.Start();

            CLEAREXCEL();

            SETPATH();
        }

        #region FUNCTION
        private void timer3_Tick(object sender, EventArgs e)
        {
            // 檢查是否為每季度的1號
            // 每3、6、9和12月的1號執行一次
            if (DateTime.Now.Day == 1 && (DateTime.Now.Month % 3 == 0 || DateTime.Now.Month % 6 == 0 || DateTime.Now.Month % 9 == 0 || DateTime.Now.Month % 12 == 0))
            {
                // 執行您的操作
                ///版費退回
                try
                {
                    DataTable DT = SEARCH_PURVERSIONSNUMS();
                    if (DT != null && DT.Rows.Count >= 1)
                    {
                        SEND_PURVERSIONSNUMS(DT);
                    }
                }
                catch
                {

                }
                finally { }


            }
        }
        private void timer2_Tick(object sender, EventArgs e)
        {
            //    ///測試用的
            //    ///
            //    // 取得目前日期和時間
            //    DateTime now = DateTime.Now;

            //    int HH = 9;
            //    int MM = 55;
            //    // 判斷是否是星期一至星期五，且時間是 09:05
            //    if (now.DayOfWeek >= DayOfWeek.Monday && now.DayOfWeek <= DayOfWeek.Friday && now.Hour == HH && now.Minute == MM)
            //    {
            //        HRAUTORUN3();
            //    }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            // 取得目前日期和時間
            DateTime now = DateTime.Now;


            string targetTime1 = "08:30";
            string currentTime1 = DateTime.Now.ToString("HH:mm");

            string targetTime2 = "08:50";
            string currentTime2 = DateTime.Now.ToString("HH:mm");

            string targetTime3 = "15:00";
            string currentTime3 = DateTime.Now.ToString("HH:mm");

            string targetTime4 = "17:00";
            string currentTime4 = DateTime.Now.ToString("HH:mm");

            label1.Text = "每日執行時間為: " + targetTime2;
            label2.Text = DateTime.Now.ToString();

            //// DayOfWeek 0 開始 (表示星期日) 到 6 (表示星期六)
            //string RUNDATE = DateTime.Now.DayOfWeek.ToString("d");//tmp2 = 4 
            //string date = "1";

            //一般用08:30
            if (currentTime1 == targetTime1)
            {
                //每星期一~星期五寄送
                if (now.DayOfWeek >= DayOfWeek.Monday && now.DayOfWeek <= DayOfWeek.Friday)
                {
                    HRAUTORUN3();
                }


            }
            //一般用08:50
            if (currentTime2 == targetTime2)
            {
                //每星期一寄送
                if (now.DayOfWeek == DayOfWeek.Monday)
                {
                    HRAUTORUN2();
                }

                //每日寄送               
                HRAUTORUN();

            }


            //採購用-15:00
            if (currentTime3 == targetTime3)
            {
                //每星期一~星期五寄送
                if (now.DayOfWeek >= DayOfWeek.Monday && now.DayOfWeek <= DayOfWeek.Friday)
                {
                    HRAUTORUN4();
                }
            }
            //採購用-17:00
            if (currentTime4 == targetTime4)
            {
                //每星期一~星期五寄送
                if (now.DayOfWeek >= DayOfWeek.Monday && now.DayOfWeek <= DayOfWeek.Friday)
                {
                    HRAUTORUN4();
                }
            }

        }
        /// <summary>
        ///  //每日寄送
        /// </summary>
        public void HRAUTORUN()
        {
            StringBuilder MSG = new StringBuilder();
            SETPATH();

            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();

            try
            {
               
                //Thread.Sleep(5000);
            }
            catch
            {
                //MSG.AppendFormat(@" 溫濕度 失敗 ||");
            }
            finally
            {

            }

            try
            {
                //針對昨天核單的 總務採購單，給申請人發出公告
                NEW_GRAFFAIRS_1005_TB_EIP_BULLETIN();
                Thread.Sleep(5000);
            }
            catch
            {
                MSG.AppendFormat(@" 針對昨天核單的 總務採購單，給申請人發出公告 失敗 ||");
            }
            finally
            {

            }

            try
            {
                //溫濕度明細
                SENDEMAIL_DAILY_QC_TEMP_CHECK();
                Thread.Sleep(5000);
            }
            catch
            {
                MSG.AppendFormat(@" 溫濕度明細 失敗 ||");
            }
            finally
            {

            }


            try
            {
                //派車
                SENDEMAIL_DAILY_TKWH_CALENDAR();
                Thread.Sleep(5000);
            }
            catch
            {
                MSG.AppendFormat(@" 派車報表  失敗 ||");
            }
            finally
            {

            }

            //溫濕度-測試
            try
            {
                SENDEMAIL_DAILY_QC_CHECK();
                Thread.Sleep(5000);
            }
            catch
            {
                MSG.AppendFormat(@" 溫濕度 失敗 ||");
            }
            finally
            {

            }

            //每日訂單明細表
            try
            {
                //path_File_COPTCD
                //每日訂單明細表
                SETPATH();
                SETFILE_COPTCD(path_File_COPTCD);
                CLEAREXCEL();

                PREPARESENDEMAIL_COPTCD(path_File_COPTCD);

                Thread.Sleep(5000);
            }
            catch
            {
                MSG.AppendFormat(@" 每日訂單明細表 失敗 ||");
            }
            finally
            {

            }
            //營銷各庫庫存通知
            try
            {
                SETPATH();
                SETFILE_POSINV(path_File_POSINV);
                CLEAREXCEL();

                PREPARESENDEMAIL_POSINV(path_File_POSINV);

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 營銷各庫庫存通知");
                MSG.AppendFormat(@" 營銷各庫庫存通知 失敗 ||");
            }
            finally
            {

            }

            ///本年新品的銷售報表
            try
            {

                SETPATH();
                SETFILE_NEWSLAES(path_File_NEWSLAES);
                CLEAREXCEL();

                PREPARESENDEMAIL_NEWSLAES(path_File_NEWSLAES);

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 本年新品的銷售報表");
                MSG.AppendFormat(@" 本年新品的銷售報表 失敗 ||");
            }
            finally
            {

            }

            //測試UOF交辨未完成
            try
            {
                CHECK_TB_EIP_SCH_DEVOLVE();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 測試UOF交辨未完成");
                MSG.AppendFormat(@" 測試UOF交辨未完成 失敗 ||");
            }
            finally
            {

            }

            //測試主管UOF交辨未完成
            try
            {
                CHECK_TB_EIP_SCH_DEVOLVE_MANAGER();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 測試主管UOF交辨未完成");
                MSG.AppendFormat(@" 測試主管UOF交辨未完成 失敗 ||");
            }
            finally
            {

            }

            //通知副總，總務未簽核的表單           
            try
            {
                PREPARE_UOF_TASK_TASK_GRAFFIR();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 通知副總，總務未簽核的表單");
                MSG.AppendFormat(@"  通知副總，總務未簽核的表單 失敗 ||");
            }
            finally
            {

            }


            //通知原請購人，總務已完成採購           
            try
            {
                FIND_UOF_GRAFFAIRS_1005();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 通知原請購人，總務已完成採購  ");
                MSG.AppendFormat(@"  通知原請購人，總務已完成採購 失敗 ||");

            }
            finally
            {

            }

            //通知各表單申請人           
            try
            {
                PREPARE_UOF_TASK_TASK_APPLICATION();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 通知各表單申請人");
                MSG.AppendFormat(@"  通知各表單申請人 失敗 ||");
            }
            finally
            {

            }

            //通知各別的被交辨人
            try
            {
                PREPARE_TB_EIP_PRIV_MESS();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 通知各別的被交辨人");
                MSG.AppendFormat(@"  通知各別的被交辨人 失敗 ||");
            }
            finally
            {

            }

            //通知交辨人         
            try
            {
                PREPARE_TB_EIP_PRIV_MESS_DIRECTOR();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 通知交辨人");
                MSG.AppendFormat(@"  通知交辨人 失敗 ||");
            }
            finally
            {

            }

            //校稿追踨          
            try
            {
                PREPAREPROOFREAD();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 校稿追踨");
                MSG.AppendFormat(@"  校稿追踨 失敗 ||");
            }
            finally
            {

            }

            //IT檢查網站是否正常           
            try
            {
                PREPAREITCHECK();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 IT檢查網站是否正常");
                MSG.AppendFormat(@"  IT檢查網站是否正常 失敗 ||");
            }
            finally
            {

            }

            //給採購人員，ERP未核單的單別、單號           
            try
            {
                PREPARESENDEMAILERPPURCHECK();

                Thread.Sleep(5000);
            }
            catch
            {
                //essageBox.Show("有錯誤 給採購人員，ERP未核單的單別、單號           ");
                MSG.AppendFormat(@"  給採購人員，ERP未核單的單別、單號  失敗 ||");
            }
            finally
            {

            }

            //測試預排製令
            ///SENDEMAIL_DAILY_MOCMANULINE
            try
            {
                SENDEMAIL_DAILY_MOCMANULINE();
                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 測試預排製令");
                MSG.AppendFormat(@"  測試預排製令  失敗 ||");
            }
            finally
            {

            }
            //測試批號錯誤
            ///SETFILELOTCHECK
            try
            {
                SETFILELOTCHECK();
                CLEAREXCEL();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 測試批號錯誤");
                MSG.AppendFormat(@"  測試批號錯誤  失敗 ||");
            }
            finally
            {

            }
            //測試未完重工單
            ///SETFILEMOCTARE
            try
            {
                SETFILEMOCTARE();
                CLEAREXCEL();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 測試未完重工單");
                MSG.AppendFormat(@"  測試未完重工單  失敗 ||");
            }
            finally
            {

            }

            ///測試已採購未結案
            ///SETFILEPURTD
            try
            {
                SETFILEPURTD();
                CLEAREXCEL();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 測試已採購未結案");
                MSG.AppendFormat(@"  測試已採購未結案  失敗 ||");
            }
            finally
            {

            }
            //測試物料安全水位
            ///SETFILEINVMC
            try
            {
                SETFILEINVMC();
                CLEAREXCEL();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 測試物料安全水位 ");
                MSG.AppendFormat(@"  測試物料安全水位  失敗 ||");
            }
            finally
            {

            }

            //測試已請購未採購
            ///SETFILEPURTB
            try
            {
                SETFILEPURTB();
                CLEAREXCEL();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 測試已請購未採購");
                MSG.AppendFormat(@"  測試已請購未採購  失敗 ||");
            }
            finally
            {

            }
            //測試半成品-製令
            ///SETFILEINVMOCTA
            try
            {
                SETFILEINVMOCTA();
                CLEAREXCEL();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 測試半成品-製令");
                MSG.AppendFormat(@"  測試半成品-製令  失敗 ||");
            }
            finally
            {

            }
            //測試製令-訂單
            ///SETFILEMOCTA            
            try
            {
                SETFILEMOCTA();
                CLEAREXCEL();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 測試製令-訂單");
                MSG.AppendFormat(@"  測試製令-訂單  失敗 ||");
            }
            finally
            {

            }
            //測試訂單變更
            /// SETFILECOPTE
            try
            {
                SETFILECOPTE();
                CLEAREXCEL();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 測試訂單變更");
                MSG.AppendFormat(@"  測試訂單變更  失敗 ||");
            }
            finally
            {

            }
            //測試請購
            ///SETFILEPURTA
            try
            {
                SETFILEPURTA();
                //SETFILEPURTA2();
                CLEAREXCEL();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 測試請購");
                MSG.AppendFormat(@"  測試請購  失敗 ||");
            }
            finally
            {

            }
            //測試訂單
            ///SETFILE
            try
            {
                SETFILE();
                CLEAREXCEL();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 測試訂單");
                MSG.AppendFormat(@"  測試訂單  失敗 ||");
            }
            finally
            {

            }

            //系統通知-每日批號檢查表         
            try
            {
                SERACHMAILLOTCHECK();
                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"系統通知-每日批號檢查表" + DateTime.Now.ToString("yyyy/MM/dd"));
                BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日批號檢查表，請查收 (批號錯誤時，要檢查「批號資料建立作業」內的有效日期、複檢日期是否也錯誤)" + Environment.NewLine + " ");
                SENDMAIL(SUBJEST, BODY, dsMAILLOTCHECK, pathFileLOTCHECK);

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 每日批號檢查表");
                MSG.AppendFormat(@"  每日批號檢查表  失敗 ||");
            }
            finally
            {

            }

            //系統通知-每日重工單未結案表       
            try
            {
                SERACHMAILMOCTARE();
                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"系統通知-每日重工單未結案表" + DateTime.Now.ToString("yyyy/MM/dd"));
                BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日重工單未結案表，請查收" + Environment.NewLine + " ");
                SENDMAIL(SUBJEST, BODY, dsMAILMOCTARE, pathFileMOCTARE);

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 每日重工單未結案表");
                MSG.AppendFormat(@"  每日重工單未結案表  失敗 ||");
            }
            finally
            {

            }

            ///系統通知-每日每日採購單未結案表
            try
            {
                //PURTD
                SERACHMAILPURTD();
                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"系統通知-每日每日採購單未結案表" + DateTime.Now.ToString("yyyy/MM/dd"));
                BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日採購單未結案表，請查收" + Environment.NewLine + " ");
                SENDMAIL(SUBJEST, BODY, dsMAILPURTD, pathFilePURTD);

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 每日每日採購單未結案表");
                MSG.AppendFormat(@"  每日每日採購單未結案表  失敗 ||");
            }
            finally
            {

            }


            /////每日物料安全水位表
            //try
            //{
            //    //INVMC
            //    //SERACHMAILINVMC();
            //    //SUBJEST.Clear();
            //    //BODY.Clear();
            //    //SUBJEST.AppendFormat(@"每日物料安全水位表" + DateTime.Now.ToString("yyyy/MM/dd"));
            //    //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日物料安全水位表，請查收" + Environment.NewLine + " ");
            //    //SENDMAIL(SUBJEST, BODY, dsMAILINVMC, pathFileINVMC);
            //}
            //catch
            //{

            //}
            //finally
            //{

            //}

            ///系統通知-每日已請購未採購表
            try
            {
                SERACHMAILPURTB();
                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"系統通知-每日已請購未採購表" + DateTime.Now.ToString("yyyy/MM/dd"));
                BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日已請購未採購表，請查收" + Environment.NewLine + " ");
                SENDMAIL(SUBJEST, BODY, dsMAILPURTB, pathFilePURTB);

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 每日已請購未採購表");
                MSG.AppendFormat(@"  每日已請購未採購表  失敗 ||");
            }
            finally
            {

            }

            ///系統通知-每日追踨半成品-製令的比對表，是否有半成品呆滯
            try
            {
                SERACHMAILINVMOCTA();
                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"系統通知-每日追踨半成品-製令的比對表，是否有半成品呆滯" + DateTime.Now.ToString("yyyy/MM/dd"));
                BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日半成品-製令表，請查收" + Environment.NewLine + " ");
                SENDMAIL(SUBJEST, BODY, dsMAILINVMOCTA, pathFileINVMOCTA);

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 每日追踨半成品-製令的比對表");
                MSG.AppendFormat(@"  每日追踨半成品-製令的比對表  失敗 ||");
            }
            finally
            {

            }


            ///系統通知-每日追踨製令未確認表
            try
            {
                SERACHMAILMOCTA();
                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"系統通知-每日追踨製令未確認表" + DateTime.Now.ToString("yyyy/MM/dd"));
                BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日製令未確認表，請查收" + Environment.NewLine + " ");
                SENDMAIL(SUBJEST, BODY, dsMAILMOCTA, pathFileMOCTA);

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 每日追踨製令未確認表");
                MSG.AppendFormat(@"  每日追踨製令未確認表  失敗 ||");
            }
            finally
            {

            }

            ///系統通知-每日追踨訂單變更追踨表
            try
            {
                SERACHMAILCOPTE();
                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"系統通知-每日追踨訂單變更追踨表" + DateTime.Now.ToString("yyyy/MM/dd"));
                BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日訂單變更表，請查收" + Environment.NewLine + "請製造生管修改相對的製令");
                SENDMAIL(SUBJEST, BODY, dsMAILCOPTE, pathFileCOPTE);


                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 每日追踨訂單變更追踨表");
                MSG.AppendFormat(@"  每日追踨訂單變更追踨表  失敗 ||");
            }
            finally
            {

            }

            ///系統通知-每日追踨製令-請購表，是否有製令已開但未請購
            try
            {
                SERACHMAILPURTA();
                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"系統通知-每日追踨製令-請購表，是否有製令已開但未請購" + DateTime.Now.ToString("yyyy/MM/dd"));
                BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日製令-請購表，請查收" + Environment.NewLine + " ");
                SENDMAIL(SUBJEST, BODY, dsMAILPURTA, pathFilePURTA);


                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 請購表，是否有製令已開但未請購");
                MSG.AppendFormat(@"  請購表，是否有製令已開但未請購  失敗 ||");
            }
            finally
            {

            }

            //系統通知-每日追踨訂單-製令追踨表，是否有訂單未開製令
            try
            {
                SERACHMAIL();
                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"系統通知-每日追踨訂單-製令追踨表，是否有訂單未開製令" + DateTime.Now.ToString("yyyy/MM/dd"));
                BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日訂單-製令追踨表，請查收" + Environment.NewLine + "若訂單沒有相對的製令則需通知製造生管開立");
                SENDMAIL(SUBJEST, BODY, dsMAIL, pathFile);

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("有錯誤 製令追踨表，是否有訂單未開製令");
                MSG.AppendFormat(@" 製令追踨表，是否有訂單未開製令  失敗 ||");
            }
            finally
            {

            }

            if (!string.IsNullOrEmpty(MSG.ToString()))
            {
                MessageBox.Show(MSG.ToString());
            }



        }
        /// <summary>
        /// //每星期一寄送
        /// </summary>
        public void HRAUTORUN2()
        {
            StringBuilder MSG = new StringBuilder();

            //研發每週通知該月樣品
            try
            {
                SENDEMAIL_TB_DEVE_NEWLISTS();
            }

            catch
            {
                MSG.AppendFormat(@" 研發每週通知該月樣品 失敗 ||");
            }
            finally
            {
            }


            //業務活動通知行銷-測試
            try
            {
                SENDEMAIL_TB_SALES_PROMOTIONS(); }

            catch
            {
                MSG.AppendFormat(@" 業務活動通知行銷 失敗 ||");
            }
            finally
            {
            }

            //每日製令準時完工率數量達交率
            try
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
            }
            catch
            {
                MSG.AppendFormat(@" 每日製令準時完工率數量達交率 失敗 ||");
            }
            finally
            {

            }

            //MessageBox.Show("OK");

            if (!string.IsNullOrEmpty(MSG.ToString()))
            {
                MessageBox.Show(MSG.ToString());
            }

        }
        /// <summary>
        /// //每星期一~星期五寄送
        /// </summary>
        public void HRAUTORUN3()
        {
            StringBuilder MSG = new StringBuilder();

            try
            {

            }
            catch
            {

            }
            finally
            {

            }

            //
            try
            {
                // 國內、外業務部業績日報表
                SENDEMAIL_DAILY_SALES_MONEY();

                Thread.Sleep(5000);
            }
            catch
            {
                //MessageBox.Show("國內、外業務部業績日報表 有錯誤");
                MSG.AppendFormat(@" 國內、外業務部業績日報表  失敗 ||");
            }
            finally
            {

            }

            if (!string.IsNullOrEmpty(MSG.ToString()))
            {
                MessageBox.Show(MSG.ToString());
            }

        }
        /// <summary>
        /// 每日提醒用
        /// 採購
        /// </summary>
        public void HRAUTORUN4()
        {
            StringBuilder MSG = new StringBuilder();

            try
            {
                //採購今日未傳真
                SENDEMAIL_TBPURCHECKFAX();

                Thread.Sleep(5000);
            }
            catch
            {
                MSG.AppendFormat(@"採購今日未傳真  失敗 ||");
            }
            finally
            {
            }


            try
            {
                //預計採購未到貨
                SENDEMAIL_PURNOTIN();

                Thread.Sleep(5000);
            }
            catch
            {
                MSG.AppendFormat(@"預計採購未到貨  失敗 ||");
            }
            finally
            {
            }

            if (!string.IsNullOrEmpty(MSG.ToString()))
            {
                MessageBox.Show(MSG.ToString());
            }

        }
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
            path_File_NEWSLAES = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日新品銷售表" + DATES.ToString();
            path_File_POSINV = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日庫存表" + DATES.ToString();
            path_File_COPTCD = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日訂單明細表" + DATES.ToString();

            pathFile_SALES_MONEYS = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日業務單位業績日報表" + DATES.ToString() + ".pdf";
            pathFile_QC_CHECK = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日溫溼度警報" + DATES.ToString() + ".pdf";
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
                // 設置查詢的超時時間，以秒為單位
                adapter1.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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

            //Add a new worksheet to workbook with the Datatable name
            Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();


            foreach (DataTable table in ds.Tables)
            {
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

                        if (TopathFile.Equals(pathFileINVMOCTA) && k == 6 && Convert.ToDecimal(table.Rows[j].ItemArray[k].ToString()) > 0)
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
                        if (TopathFile.Equals(pathFileMOCCOP) && k == 16 && Convert.ToDecimal(table.Rows[j].ItemArray[k].ToString()) < 0)
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

            // 释放COM对象
            Marshal.ReleaseComObject(excelWorkSheet);
            Marshal.ReleaseComObject(excelWorkBook);
            Marshal.ReleaseComObject(excelApp);
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
                // 設置查詢的超時時間，以秒為單位
                adapterMAIL.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterCOPTE.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                    cmd.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterMAILCOPTE.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterPURTA.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterPURTA.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                INSERTLOG(pathFilePURTA, ex.ToString());
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
                // 設置查詢的超時時間，以秒為單位
                adapterPURTA.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterMAILPURTA.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterMOCTA.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterMAILMOCTA.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterINVMOCTA.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterMAILINVMOCTA.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterPURTB.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterMAILPURTB.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterMOCINVCHECK.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterMOCCOP.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterMAILMOCCOP.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterINVMC.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterMAILINVMC.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterPURTD.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterMAILMOCTARE.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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

        public DataTable SERACH_MAIL_MOCMANULINE()
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
                // 設置查詢的超時時間，以秒為單位
                adapterMAILMOCMANULINE.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapterMAILMOCMANULINE.Fill(dsMAILMOCMANULINE, "dsMAILMOCMANULINE");
                sqlConn.Close();

                if (dsMAILMOCMANULINE.Tables["dsMAILMOCMANULINE"].Rows.Count >= 1)
                {
                    return dsMAILMOCMANULINE.Tables["dsMAILMOCMANULINE"];
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
                // 設置查詢的超時時間，以秒為單位
                adapterMAILLOTCHECK.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterMAILPURTD.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterMOCTARE.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                                    SELECT KINDS,TH004 AS '品號',TH005 AS '品名',TH010 AS '批號',TH036 AS '有效日',TH117 AS '製造日',TH001 AS '單別',TH002 AS '單號',TH003 AS '序號',COMMET AS '備註' 
                                    FROM 
                                    ( 
                                    SELECT '進貨單' AS KINDS,TG003,TH004,TH005,TH010,TH036,TH117,TH001,TH002,TH003,'批號<>有效日' AS COMMET 
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH 
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TH030='Y' 
                                    AND TH004 LIKE '1%' 
                                    AND TH010<>TH036 
                                    UNION ALL 
                                    SELECT '進貨單' AS KINDS,TG003,TH004,TH005,TH010,TH036,TH117,TH001,TH002,TH003,'批號<>製造日' AS COMMET 
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH 
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TH030='Y' 
                                    AND TH004 LIKE '2%' 
                                    AND TH010<>TH117 
                                    UNION ALL 
                                    SELECT '進貨單' AS KINDS,TG003,TH004,TH005,TH010,TH036,TH117,TH001,TH002,TH003,'批號<>製造日' AS COMMET 
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH 
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TH030='Y' 
                                    AND TH004 LIKE '3%' 
                                    AND TH010<>TH117 
                                    UNION ALL 
                                    SELECT '進貨單' AS KINDS,TG003,TH004,TH005,TH010,TH036,TH117,TH001,TH002,TH003,'批號<>有效日' AS COMMET 
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH 
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TH030='Y' 
                                    AND TH004 LIKE '4%' 
                                    AND TH010<>TH036 
                                    UNION ALL 
                                    SELECT '進貨單' AS KINDS,TG003,TH004,TH005,TH010,TH036,TH117,TH001,TH002,TH003,'批號<>有效日' AS COMMET 
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH 
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TH030='Y' 
                                    AND TH004 LIKE '5%' 
                                    AND TH010<>TH036 
                                    UNION ALL 
                                    SELECT  '進貨單' AS KINDS,TG003,TH004,TH005,TH010,TH036,TH117,TH001,TH002,TH003,'批號日錯誤' AS COMMET 
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH 
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TH030='Y' 
                                    AND TH004 LIKE '1%' 
                                    AND ISDATE(TH010)<>1
                                    AND TH009 NOT LIKE '21%'
                                    UNION ALL 
                                    SELECT '入庫單' AS KINDS ,TF003,TG004,TG005,TG017,TG018,TG040,TG001,TG002,TG003,'批號<>製造日' AS COMMET 
                                    FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG 
                                    WHERE TF001=TG001 AND TF002=TG002 
                                    AND TF003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TG022='Y' 
                                    AND TG004 LIKE '3%'  
                                    AND TG004 NOT LIKE '307%' 
                                    AND TG017<>TG040 
                                    UNION ALL 
                                    SELECT '入庫單' AS KINDS ,TF003,TG004,TG005,TG017,TG018,TF003,TG001,TG002,TG003,'批號<>有效日' AS COMMET 
                                    FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG 
                                    WHERE TF001=TG001 AND TF002=TG002 
                                    AND TF003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TG022='Y' 
                                    AND TG004 LIKE '4%' 
                                    AND TG017<>TG018 

                                    UNION ALL 
                                    SELECT '入庫單' AS KINDS ,TF003,TG004,TG005,TG017,TG018,TG040,TG001,TG002,TG003,'批號日錯誤' AS COMMET 
                                    FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG 
                                    WHERE TF001=TG001 AND TF002=TG002 
                                    AND TF003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TG022='Y' 
                                    AND TG004 LIKE '3%'  
                                    AND TG004 NOT LIKE '307%' 
                                    AND ISDATE(TG017)<>1
                                    UNION ALL 
                                    SELECT '託外入庫單' AS KINDS ,TH003,TI004,TI005,TI010,TI011,TI061,TI001,TI002,TI003,'批號<>製造日' AS COMMET 
                                    FROM [TK].dbo.MOCTH,[TK].dbo.MOCTI 
                                    WHERE TH001=TI001 AND TH002=TI002 
                                    AND TI061>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TI004 LIKE '3%'   
                                    AND TI037='Y' 
                                    AND TI010<>TI061 
                                    AND TI001+TI002+TI003 NOT IN ('A591201906240010001','A591201911220010001','A591201911250030001')  
                                    UNION ALL 
                                    SELECT '託外入庫單' AS KINDS ,TH003,TI004,TI005,TI010,TI011,TI061,TI001,TI002,TI003,'批號<>有效日' AS COMMET 
                                    FROM [TK].dbo.MOCTH,[TK].dbo.MOCTI 
                                    WHERE TH001=TI001 AND TH002=TI002 
                                    AND TI061>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TI004 LIKE '4%' 
                                    AND TI037='Y' 
                                    AND TI010<>TI011
                                    UNION ALL 
                                    SELECT '託外入庫單' AS KINDS ,TH003,TI004,TI005,TI010,TI011,TI061,TI001,TI002,TI003,'批號不是日期' AS COMMET 
                                    FROM [TK].dbo.MOCTH,[TK].dbo.MOCTI 
                                    WHERE TH001=TI001 AND TH002=TI002 
                                    AND TI061>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TI004 LIKE '4%' 
                                    AND TI037='Y' 
                                    AND ISDATE(TI011)<>1
                                    UNION 

                                    SELECT '進貨單' AS KINDS,TG003,TH004,TH005,TH010,TH036,TH117,TH001,TH002,TH003,'製造日是未來日' AS COMMET 
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH 
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TH030='Y' 
                                    AND TH117>=CONVERT(NVARCHAR,DATEADD(DAY,-0,GETDATE()),112  ) 
                                    UNION ALL
                                    SELECT '進貨單' AS KINDS,TG003,TH004,TH005,TH010,TH036,TH117,TH001,TH002,TH003,'製造日不是日期' AS COMMET 
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH 
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TH030='Y' 
                                    AND ISDATE(TH117)<>1
                                    UNION ALL
                                    SELECT '入庫單' AS KINDS ,TF003,TG004,TG005,TG017,TG018,TG040,TG001,TG002,TG003,'製造日是未來日' AS COMMET 
                                    FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG 
                                    WHERE TF001=TG001 AND TF002=TG002 
                                    AND TF003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TG022='Y' 
                                    AND TG004 LIKE '3%'  
                                    AND TG004 NOT LIKE '307%' 
                                    AND TG040>=CONVERT(NVARCHAR,DATEADD(DAY,-0,GETDATE()),112  ) 
                                    UNION ALL
                                    SELECT '入庫單' AS KINDS ,TF003,TG004,TG005,TG017,TG018,TG040,TG001,TG002,TG003,'製造日不是日期' AS COMMET 
                                    FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG 
                                    WHERE TF001=TG001 AND TF002=TG002 
                                    AND TF003>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TG022='Y' 
                                    AND TG004 LIKE '3%'  
                                    AND TG004 NOT LIKE '307%' 
                                    AND ISDATE(TG040)<>1
                                    UNION ALL
                                    SELECT '託外入庫單' AS KINDS ,TH003,TI004,TI005,TI010,TI011,TI061,TI001,TI002,TI003,'製造日是未來日' AS COMMET 
                                    FROM [TK].dbo.MOCTH,[TK].dbo.MOCTI 
                                    WHERE TH001=TI001 AND TH002=TI002 
                                    AND TI061>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TI004 LIKE '4%' 
                                    AND TI037='Y' 
                                    AND TI061>=CONVERT(NVARCHAR,DATEADD(DAY,-0,GETDATE()),112  ) 
                                    UNION ALL
                                    SELECT '託外入庫單' AS KINDS ,TH003,TI004,TI005,TI010,TI011,TI061,TI001,TI002,TI003,'製造日不是日期' AS COMMET 
                                    FROM [TK].dbo.MOCTH,[TK].dbo.MOCTI 
                                    WHERE TH001=TI001 AND TH002=TI002 
                                    AND TI061>= CONVERT(NVARCHAR,DATEADD(DAY,-7,GETDATE()),112  ) 
                                    AND TI004 LIKE '4%' 
                                    AND TI037='Y' 
                                    AND ISDATE(TI061)<>1
                                    ) 
                                    AS TEMP 
                                    WHERE  TH004 IN (
                                    SELECT MB001
                                    FROM [TK].dbo.INVMB
                                    WHERE MB022 NOT IN ('N')
                                    )
                                    ORDER BY TH004  
                                    ");


                adapterLOTCHECK = new SqlDataAdapter(@"" + sbSql, sqlConn);
                //adapterPURTD.SelectCommand.Parameters.AddWithValue("@MC002", "20004");

                sqlCmdBuilderLOTCHECK = new SqlCommandBuilder(adapterLOTCHECK);


                sqlConn.Open();
                dsLOTCHECK.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapterLOTCHECK.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapterMOCMANULINE.SelectCommand.CommandTimeout = 600;
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


        public void ADDLOG(DateTime DATES, string SOURCE, string EX)
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


                if (DSPURCHECK != null && DSPURCHECK.Tables[0].Rows.Count > 0)
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


                if (DSTKPUR_PURTATBCHAGE_DCHECK != null && DSTKPUR_PURTATBCHAGE_DCHECK.Tables[0].Rows.Count > 0)
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


                if (DSPURTDCHECK != null && DSPURTDCHECK.Tables[0].Rows.Count > 0)
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


                if (DS_PURTB_NOTIN_PURTD != null && DS_PURTB_NOTIN_PURTD.Tables[0].Rows.Count > 0)
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
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                if (DSFINDPURCHECKMAILTO.Tables[0].Rows.Count > 0)
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

                    foreach (DataRow DR in DSFINDPURCHECKMAILTO.Tables[0].Rows)
                    {

                        try
                        {
                            MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail
                            //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show("有錯誤");

                            //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                            //ex.ToString();
                        }
                    }

                    MySMTP.Send(MyMail);

                    MyMail.Dispose(); //釋放資源
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
        /// 實際寄出MEAIL，採購人員，ERP未核單的單別、單號
        /// </summary>
        public void SENDE_TO_PURTYPES(StringBuilder Subject, StringBuilder Body)
        {
            DataSet DSFMAILTO = FINDPURCHECKMAILTO("PURTYPES");

            try
            {
                if (DSFMAILTO.Tables[0].Rows.Count > 0)
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

                    //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email

                    foreach (DataRow DR in DSFMAILTO.Tables[0].Rows)
                    {

                        try
                        {
                            MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail

                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show("有錯誤");

                            //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                            //ex.ToString();
                        }
                    }

                    MySMTP.Send(MyMail);

                    MyMail.Dispose(); //釋放資源
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

            foreach (DataRow DR in DTWEBLINKS.Rows)
            {
                if (CheckUrlVisit(DR["WEBLINKS"].ToString()) != true)
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
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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

            StringBuilder LINE_NOTIFY = new StringBuilder();

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

                        LINE_NOTIFY.AppendFormat(@"%0D%0A {0} %0D%0A 此網站不通，請檢查網站狀況 %0D%0A {1}  %0D%0A {2}", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), DR["COMMENTS"].ToString(), DR["WEBLINKS"].ToString());
                    }
                    else
                    {
                        BODY.AppendFormat(" <br>"
                                 + "{0} 此網站正常"
                                 + " <br>"
                                  , DR["COMMENTS"].ToString() + " " + DR["WEBLINKS"].ToString());

                        LINE_NOTIFY.AppendFormat(@"%0D%0A {0} %0D%0A 此網站正常%0D%0A {1}  %0D%0A {2}", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), DR["COMMENTS"].ToString(), DR["WEBLINKS"].ToString());

                    }


                }



                BODY.AppendFormat(" "
                              + "<br>" + "謝謝"

                              + "</span><br>");


                if (ISCHECK.Equals("N"))
                {
                    SUBJEST.AppendFormat(@" 有異常");

                }
                else
                {
                    SUBJEST.AppendFormat(@" ");
                }

                SENDEMAILITCHECK(SUBJEST, BODY);

                SEND_LINE(LINE_NOTIFY.ToString());

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


                    foreach (DataRow DR in DSFINDITCHECKMAILTO.Tables[0].Rows)
                    {
                        try
                        {
                            MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail
                                                                  //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                        }
                        catch (Exception ex)
                        {
                            // MessageBox.Show("有錯誤");

                            //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                            //ex.ToString();
                        }
                    }

                    MySMTP.Send(MyMail);

                    MyMail.Dispose(); //釋放資源
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
            DataSet DSUOFUOFFORM1002 = UOFUOFFORM1002();


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


                SUBJEST.AppendFormat(@"系統通知-老楊食品-每日-交辨未完成的項目及設計表單簽核未完成的項目 ，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));


                //交辨未完成的項目及交辨人回覆狀況
                BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                    + "<br>" + "交辨未完成的項目及交辨人回覆狀況 明細如下"

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


                //設計表單簽核未完成的項目
                BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                    + "<br>" + "設計表單簽核未完成的項目 明細如下"

                    );


                if (DSUOFUOFFORM1002.Tables[0].Rows.Count > 0)
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                    BODY.AppendFormat(@"<table> ");
                    BODY.AppendFormat(@"<tr >");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=10% "">目前簽核人員</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=40% "">申請人員</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=40% "">申請表單</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=10% "">表單編號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=10% "">表單申請日期</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' width=10% "">表單逾期天數</th>");


                    BODY.AppendFormat(@"</tr> ");

                    foreach (DataRow DR in DSUOFUOFFORM1002.Tables[0].Rows)
                    {

                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["NAME"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體'  width=40% "">" + DR["APPLICANT_NAME"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體'  width=30% "">" + DR["FORM_NAME"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["DOC_NBR"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["START_TIME"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["DAYS"].ToString() + "</td>");


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
                                    ,TB_EIP_SCH_DEVOLVE.SUBJECT AS '校稿區內容'
                                    ,TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID AS 'DEVOLVE_GUID'
                                    ,TB_EIP_SCH_WORK.SUBJECT AS '交辨項目'
                                    ,TB_EIP_SCH_WORK.EXECUTE_USER AS '交辨'
                                    ,TB_EIP_SCH_WORK.WORK_STATE AS 'WORK_STATE'
                                    ,(ISNULL(TB_EIP_SCH_WORK.PROCEEDING_DESC,'')+ISNULL(TB_EIP_SCH_WORK.COMPLETE_DESC,''))  AS '交辨回覆'
                                    ,TB_EB_USER.NAME AS '被交辨人'
                                    ,(CASE  WHEN TB_EIP_SCH_WORK.WORK_STATE='Completed' THEN '審稿完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Audit' THEN '交辨完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Proceeding' THEN TB_EB_USER.NAME+' 已回覆，但交辨人  '+USER2.NAME+' 未確認' WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN TB_EB_USER.NAME+' 未開始回覆' END) AS '交辨狀態'
                                    ,(CASE WHEN ISNULL(TB_EIP_SCH_WORK.COMPLETE_TIME,'')<>'' THEN CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,24),1,8) ELSE CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,24),1,8) END)  AS '回覆時間'

                                    ,TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.*
                                    ,TB_EB_USER.ACCOUNT
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
                                    AND TB_EIP_SCH_WORK.WORK_STATE NOT IN ('Completed','Audit')
                                    ORDER BY TB_EIP_SCH_DEVOLVE.CREATE_TIME DESC
                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DSPROOFREAD.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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


        public DataSet UOFUOFFORM1002()
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
                                   
                                        SELECT
                                        usr2.NAME
                                        ,(CASE WHEN  usr.IS_SUSPENDED = 1 THEN  usr.NAME + '(x)' WHEN  ISNULL(usr.ACCOUNT,'''') = '' THEN  'unknown user' ELSE usr.NAME END) AS APPLICANT_NAME
                                        ,form.FORM_NAME
                                        ,DOC_NBR
                                        ,CONVERT(NVARCHAR,NODES.START_TIME,111) AS 'START_TIME'
                                        ,DATEDIFF(DAY,START_TIME,GETDATE()) AS 'DAYS'
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
                                        ,task.CURRENT_DOC.value('(/Form/FormFieldValue/FieldItem[@fieldId=""00010""]/@fieldValue)[1]', 'nvarchar(50)') AS '產品設計'
                                        ,task.CURRENT_DOC.value('(/Form/FormFieldValue/FieldItem[@fieldId=""RDFrm1002PD""]/@fieldValue)[1]', 'nvarchar(50)') AS '設計需求'

                                        FROM [UOF].dbo.TB_WKF_TASK task
                                        INNER JOIN [UOF].dbo.TB_WKF_FORM_VERSION formVer ON task.FORM_VERSION_ID = formVer.FORM_VERSION_ID
                                        INNER JOIN [UOF].dbo.TB_WKF_FORM form  ON  formVer.FORM_ID = form.FORM_ID 
                                        LEFT JOIN [UOF].dbo.TB_EB_USER [usr]  ON task.USER_GUID = usr.USER_GUID
                                        LEFT JOIN [UOF].dbo.TB_WKF_TASK_NODE [NODES] ON NODES.SITE_ID=task.CURRENT_SITE_ID 
                                        LEFT JOIN [UOF].dbo.TB_EB_USER [usr2]  ON NODES.ORIGINAL_SIGNER = [usr2].USER_GUID
                                        WHERE
                                        1=1  
                                        AND  TASK_STATUS NOT IN ('2')
                                        AND ISNULL([NODES].SIGN_STATUS,999)<>0
                                        AND form.FORM_NAME IN ('1002.產品設計申請','1002.設計需求內容清單')

                                        ORDER BY form.FORM_NAME,usr2.NAME,DAYS DESC,DOC_NBR
                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DSPROOFREAD.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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


                    foreach (DataRow DR in UOFPROOFEAD.Tables[0].Rows)
                    {

                        try
                        {
                            MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail

                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show("有錯誤");

                            //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                            //ex.ToString();
                        }
                    }

                    //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                    MySMTP.Send(MyMail);

                    MyMail.Dispose(); //釋放資源


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


            if (DTFIND_USER_GUID.Rows.Count > 0)
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
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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


        public void ADD_TB_EIP_PRIV_MESS_DIRECTOR(string USER_GUID, string MESS)
        {
            Guid NEW = Guid.NewGuid();
            string MESSAGE_GUID = NEW.ToString();
            string TOPIC = "系統通知-每日校稿的被交辨人未回覆項目" + DateTime.Now.ToString("yyyyMMdd");
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
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
            string TOPIC = "系統通知-每日校稿的未回覆項目，請至交辨區回覆校稿" + DateTime.Now.ToString("yyyyMMdd");
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

            if (DT_FIND_UOF_TASK_APPLICATION != null && DT_FIND_UOF_TASK_APPLICATION.Rows.Count >= 1)
            {
                foreach (DataRow DR in DT_FIND_UOF_TASK_APPLICATION.Rows)
                {
                    DT_FIND_UOF_TASK_APPLICATION_FORM = FIND_UOF_TASK_APPLICATION_FORM(DR["APPLICANT_NAME"].ToString());

                    if (DT_FIND_UOF_TASK_APPLICATION_FORM != null && DT_FIND_UOF_TASK_APPLICATION_FORM.Rows.Count >= 1)
                    {
                        SEND_UOF_TASK_APPLICATION_FORM(DR["APPLICANT_NAME"].ToString(), DR["APPLICANT_EMAIL"].ToString(), DT_FIND_UOF_TASK_APPLICATION_FORM);
                    }
                }

            }
        }


        public DataTable FIND_UOF_TASK_APPLICATION()
        {
            StringBuilder MESS = new StringBuilder();
            DataSet DS = new DataSet();

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
                                    FROM [UOF].dbo.TB_WKF_TASK task
                                    INNER JOIN [UOF].dbo.TB_WKF_FORM_VERSION formVer ON task.FORM_VERSION_ID = formVer.FORM_VERSION_ID
                                    INNER JOIN [UOF].dbo.TB_WKF_FORM form  ON  formVer.FORM_ID = form.FORM_ID 
                                    LEFT JOIN [UOF].dbo.TB_EB_USER [usr]  ON task.USER_GUID = usr.USER_GUID
                                    LEFT JOIN [UOF].dbo.TB_WKF_TASK_NODE [NODES] ON NODES.SITE_ID=task.CURRENT_SITE_ID 
                                    LEFT JOIN [UOF].dbo.TB_EB_USER [usr2]  ON NODES.ORIGINAL_SIGNER = [usr2].USER_GUID
                                    LEFT JOIN [UOF].dbo.[TB_EB_EMPL_DEP] ON [TB_EB_EMPL_DEP].USER_GUID=[usr2].USER_GUID AND [TB_EB_EMPL_DEP].ORDERS='0'
                                    LEFT JOIN [UOF].dbo.[TB_EB_JOB_TITLE] ON [TB_EB_EMPL_DEP].TITLE_ID=[TB_EB_JOB_TITLE].TITLE_ID


                                    WHERE
                                    1=1  
                                    AND  TASK_STATUS NOT IN ('2','3','4')
                                    AND ISNULL([NODES].SIGN_STATUS,999)<>0
                                    AND DATEDIFF(HOUR,CONVERT(datetime,START_TIME),GETDATE())>=36
                                    AND DATEDIFF(DAY,CONVERT(datetime,START_TIME),GETDATE())<=365
                                    AND FORM_NAME NOT IN (SELECT  [FORM_NAME]  FROM [UOF].[dbo].[Z_NOT_MQ_FORM_NAME] )

                                    )  AS TEMP 
                                    WHERE ISNULL(APPLICANT_EMAIL,'')<>''
                                    
                                    GROUP BY APPLICANT_NAME,APPLICANT_EMAIL
                                    ORDER BY APPLICANT_NAME,APPLICANT_EMAIL

                                   

                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(DS, "DS");
                sqlConn.Close();



                if (DS.Tables["DS"].Rows.Count > 0)
                {

                    return DS.Tables["DS"];
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
                                    FROM [UOF].dbo.TB_WKF_TASK task
                                    INNER JOIN [UOF].dbo.TB_WKF_FORM_VERSION formVer ON task.FORM_VERSION_ID = formVer.FORM_VERSION_ID
                                    INNER JOIN [UOF].dbo.TB_WKF_FORM form  ON  formVer.FORM_ID = form.FORM_ID 
                                    LEFT JOIN [UOF].dbo.TB_EB_USER [usr]  ON task.USER_GUID = usr.USER_GUID
                                    LEFT JOIN [UOF].dbo.TB_WKF_TASK_NODE [NODES] ON NODES.SITE_ID=task.CURRENT_SITE_ID 
                                    LEFT JOIN [UOF].dbo.TB_EB_USER [usr2]  ON NODES.ORIGINAL_SIGNER = [usr2].USER_GUID
                                    LEFT JOIN [UOF].dbo.[TB_EB_EMPL_DEP] ON [TB_EB_EMPL_DEP].USER_GUID=[usr2].USER_GUID AND [TB_EB_EMPL_DEP].ORDERS='0'
                                    LEFT JOIN [UOF].dbo.[TB_EB_JOB_TITLE] ON [TB_EB_EMPL_DEP].TITLE_ID=[TB_EB_JOB_TITLE].TITLE_ID


                                    WHERE
                                    1=1  
                                    AND  TASK_STATUS NOT IN ('2','3','4')
                                    AND ISNULL([NODES].SIGN_STATUS,999)<>0
                                 
                                    AND DATEDIFF(HOUR,CONVERT(datetime,START_TIME),GETDATE())>=36
                                    AND DATEDIFF(DAY,CONVERT(datetime,START_TIME),GETDATE())<=365
                                    AND FORM_NAME NOT IN (SELECT  [FORM_NAME]  FROM [UOF].[dbo].[Z_NOT_MQ_FORM_NAME] )

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
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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

        public void SEND_UOF_TASK_APPLICATION_FORM(string APPLICANT_NAME, string APPLICANT_EMAIL, DataTable DT)
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
                        //MyMail.To.Add("tk290@tkfood.com.tw");                                //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                        MySMTP.Send(MyMail);

                        MyMail.Dispose(); //釋放資源


                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show("有錯誤");

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

            if (DTSEARCHUOF_GRAFFAIRS_1005 != null && DTSEARCHUOF_GRAFFAIRS_1005.Rows.Count >= 1)
            {
                foreach (DataRow DR in DTSEARCHUOF_GRAFFAIRS_1005.Rows)
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


                    string MESSAGES = GA003 + " 申請的請購單:" + GA002 + "，物品:" + GA005 + "，已由" + GA999 + " 在" + GA015 + "購買完成。";


                    if (!string.IsNullOrEmpty(USER_GUID))
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
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
        public void SEND_MESSAGE_UOF_GRAFFAIRS_1005(string USER_GUID, string MESS)
        {
            Guid NEW = Guid.NewGuid();
            string MESSAGE_GUID = NEW.ToString();
            string TOPIC = "系統通知 " + MESS;
            string MESSAGE_CONTENT = "系統通知 " + MESS;
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
        public void SEND_EMAIL_UOF_GRAFFAIRS_1005(string EMAILTO, string Subject, string Body)
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
                    //MessageBox.Show("有錯誤");

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

        public void PREPARE_UOF_TASK_TASK_GRAFFIR()
        {
            DataTable DT_OF_TASK_TASK_GRAFFIR = new DataTable();
            DataTable GRAFFIR_TO_EMAIL = new DataTable();
            DT_OF_TASK_TASK_GRAFFIR = FIND_OF_TASK_TASK_GRAFFIR();
            GRAFFIR_TO_EMAIL = FIND_GRAFFIR_TO_EMAIL();



            try
            {
                if (DT_OF_TASK_TASK_GRAFFIR != null && DT_OF_TASK_TASK_GRAFFIR.Rows.Count > 0 && GRAFFIR_TO_EMAIL != null && GRAFFIR_TO_EMAIL.Rows.Count > 0)
                {
                    SEND_UOF_TASK_FORM_GRAFFIR(GRAFFIR_TO_EMAIL, DT_OF_TASK_TASK_GRAFFIR);
                }
            }
            catch
            {

            }



        }

        public DataTable FIND_OF_TASK_TASK_GRAFFIR()
        {

            DataSet DS_FIND_OF_TASK_TASK_GRAFFIR = new DataSet();

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

                //   AND DOC_NBR = 'GA1005230100006'

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
                                    FROM [UOF].dbo.TB_WKF_TASK task
                                    INNER JOIN [UOF].dbo.TB_WKF_FORM_VERSION formVer ON task.FORM_VERSION_ID = formVer.FORM_VERSION_ID
                                    INNER JOIN [UOF].dbo.TB_WKF_FORM form  ON  formVer.FORM_ID = form.FORM_ID 
                                    LEFT JOIN [UOF].dbo.TB_EB_USER [usr]  ON task.USER_GUID = usr.USER_GUID
                                    LEFT JOIN [UOF].dbo.TB_WKF_TASK_NODE [NODES] ON NODES.SITE_ID=task.CURRENT_SITE_ID 
                                    LEFT JOIN [UOF].dbo.TB_EB_USER [usr2]  ON NODES.ORIGINAL_SIGNER = [usr2].USER_GUID
                                    LEFT JOIN [UOF].dbo.[TB_EB_EMPL_DEP] ON [TB_EB_EMPL_DEP].USER_GUID=[usr2].USER_GUID AND [TB_EB_EMPL_DEP].ORDERS='0'
                                    LEFT JOIN [UOF].dbo.[TB_EB_JOB_TITLE] ON [TB_EB_EMPL_DEP].TITLE_ID=[TB_EB_JOB_TITLE].TITLE_ID


                                    WHERE
                                    1=1  
                                    AND  TASK_STATUS NOT IN ('2','3','4')
                                    AND ISNULL([NODES].SIGN_STATUS,999)<>0
                                 
                                    AND DATEDIFF(HOUR,CONVERT(datetime,START_TIME),GETDATE())>=36
                                    AND DATEDIFF(DAY,CONVERT(datetime,START_TIME),GETDATE())<=365
                                    AND FORM_NAME NOT IN (SELECT  [FORM_NAME]  FROM [UOF].[dbo].[Z_NOT_MQ_FORM_NAME] )

                                    ) AS TEMP
                                    WHERE 1=1
                                    AND CURRENTNAME IN (SELECT  [CURRENT_NAMES]  FROM [UOF].[dbo].[Z_UOF_GRAFFIR_CURRENT_NAMES])
                                    GROUP BY APPLICANT_NAME,FORM_NAME,DOC_NBR,START_TIME,CURRENTNAME
                                    ORDER BY FORM_NAME,DOC_NBR

                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_FIND_OF_TASK_TASK_GRAFFIR.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(DS_FIND_OF_TASK_TASK_GRAFFIR, "DS_FIND_OF_TASK_TASK_GRAFFIR");
                sqlConn.Close();



                if (DS_FIND_OF_TASK_TASK_GRAFFIR.Tables["DS_FIND_OF_TASK_TASK_GRAFFIR"].Rows.Count > 0)
                {

                    return DS_FIND_OF_TASK_TASK_GRAFFIR.Tables["DS_FIND_OF_TASK_TASK_GRAFFIR"];
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

        public DataTable FIND_GRAFFIR_TO_EMAIL()
        {
            DataSet DS_FIND_GRAFFIR_TO_EMAIL = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();



            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                //   AND DOC_NBR = 'GA1005230100006'

                sbSql.AppendFormat(@"                                    
                                  SELECT [ID]
                                    ,[SENDTO]
                                    ,[MAIL]
                                    ,[NAME]
                                    ,[COMMENTS]
                                    FROM [TKMQ].[dbo].[MQSENDMAIL]
                                    WHERE [SENDTO]='PURTOCC'

                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_FIND_GRAFFIR_TO_EMAIL.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(DS_FIND_GRAFFIR_TO_EMAIL, "DS_FIND_GRAFFIR_TO_EMAIL");
                sqlConn.Close();



                if (DS_FIND_GRAFFIR_TO_EMAIL.Tables["DS_FIND_GRAFFIR_TO_EMAIL"].Rows.Count > 0)
                {

                    return DS_FIND_GRAFFIR_TO_EMAIL.Tables["DS_FIND_GRAFFIR_TO_EMAIL"];
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

        public void SEND_UOF_TASK_FORM_GRAFFIR(DataTable TO_EMAIL, DataTable DT)
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


                SUBJEST.AppendFormat(@"系統通知-請查收-每日-UOF表單中，總務未核單的簽核人及表單單號，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                //ERP 採購相關單別、單號未核準的明細
                //
                BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                    + "<br>" + "系統通知-請查收-每日-UOF表單中，總務未核單的簽核人及表單單號，謝謝"
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
                        foreach (DataRow DR in TO_EMAIL.Rows)
                        {
                            MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail
                        }

                        //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                        MySMTP.Send(MyMail);

                        MyMail.Dispose(); //釋放資源

                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show("有錯誤");

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

        public void SEND_LINE(string message)
        {
            //2025年3月31日結束本服務
            //LINE向用戶傳送通知的商品服務，建議可改用功能更豐富的Messaging API。
            //string token = "iJgYn1ZKgcTcCPKioCM4ispXQFu1gD7uegpufl7mkVV";
            //string message = "Hello, world! "+DateTime.Now.ToString("yyyyMMddHHmmss");

            DataTable dt = GetDataFromMSSQL("SELECT  [TOKEN] ,[KINDS] ,[COMMENTS] FROM [TKIT].[dbo].[TB_LINE_TOKEN] WHERE [KINDS]='IT'");

            string token = dt.Rows[0]["TOKEN"].ToString();

            string url = "https://notify-api.line.me/api/notify";
            try
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.DefaultConnectionLimit = 9999;
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12; // Use TLS 1.2, TLS 1.1, and TLS 1.0

                var request = (HttpWebRequest)WebRequest.Create(url);
                var postData = string.Format("message={0}", message);
                var data = Encoding.UTF8.GetBytes(postData);

                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = data.Length;
                request.Headers.Add("Authorization", "Bearer " + token);

                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072; // Use TLS 1.2
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true; // Bypass certificate validation

                using (var stream = request.GetRequestStream()) stream.Write(data, 0, data.Length);
                var response = (HttpWebResponse)request.GetResponse();
                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }



        public DataTable GetDataFromMSSQL(string sql)
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);


            try
            {
                // 建立資料庫連線物件
                using (SqlConnection connection = new SqlConnection(sqlsb.ConnectionString))
                {
                    // 建立資料庫查詢命令物件
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        // 建立資料庫資料適配器物件
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            // 建立 DataTable 物件
                            DataTable dataTable = new DataTable();

                            // 使用資料適配器物件填充 DataTable 物件
                            adapter.Fill(dataTable);

                            // 回傳 DataTable 物件
                            return dataTable;
                        }
                    }
                }
            }
            catch
            {
                return null;
            }

        }

        public void SEND_TEST_MAIL()
        {
            string MySMTPCONFIG = ConfigurationManager.AppSettings["MySMTP"];
            string NAME = ConfigurationManager.AppSettings["NAME"];
            string PW = ConfigurationManager.AppSettings["PW"];

            System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
            MyMail.From = new System.Net.Mail.MailAddress("tk290@tkfood.com.tw");

            //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
            //MyMail.Subject = "每日訂單-製令追踨表"+DateTime.Now.ToString("yyyy/MM/dd");
            MyMail.Subject = "Subject";
            //MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
            MyMail.Body = "Body";
            MyMail.IsBodyHtml = true; //是否使用html格式

            //加上附圖
            //string path = System.Environment.CurrentDirectory + @"/Images/emaillogo.jpg";
            //MyMail.AlternateViews.Add(GetEmbeddedImage(path, Body));

            System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
            MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);


            try
            {
                MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email，多筆mail
                                                      //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                MySMTP.Send(MyMail);

                MyMail.Dispose(); //釋放資源

                MessageBox.Show("寄送成功 ");


            }
            catch (Exception ex)
            {
                //MessageBox.Show("有錯誤 ");

                //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                //ex.ToString();
            }
        }

        public void SEND_TEST_MAIL_2()
        {
            string recipientEmail = "tk290@tkfood.com.tw";
            string senderEmail = "tk290@tkfood.com.tw";
            string subject = "Email Subject";
            string body = "Email Body";

            MailMessage message = new MailMessage(senderEmail, recipientEmail, subject, body);
            SmtpClient smtpClient = new SmtpClient("officemail.cloudmax.com.tw", 587);

            // 設定 smtpClient 的帳號密碼，如果需要的話
            smtpClient.Credentials = new System.Net.NetworkCredential("tkpublic@tkfood.com.tw", "@@tkmail629");

            smtpClient.EnableSsl = true; // 如果需要使用 SSL 連線，需要設定為 true

            try
            {
                smtpClient.Send(message); // 發送郵件

                MessageBox.Show("寄送成功 ");
            }
            catch (Exception ex)
            {
                //MessageBox.Show("有錯誤 ");
            }

        }

        //交辨未完成meail
        public void CHECK_TB_EIP_SCH_DEVOLVE()
        {
            //找出所有被交辨人  
            DataTable DT = FIND_TB_EIP_SCH_DEVOLVE_NAMES();

            if (DT != null && DT.Rows.Count >= 1)
            {
                SEND_EMAIL_TB_EIP_SCH_DEVOLVE(DT);
            }
        }

        //找出交辨的所有 被交辨人
        public DataTable FIND_TB_EIP_SCH_DEVOLVE_NAMES()
        {
            DataSet DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES = new DataSet();

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

                //   AND DOC_NBR = 'GA1005230100006'

                sbSql.AppendFormat(@"                          
                                
                                    SELECT 
                                    TB_EB_USER.NAME AS '被交辨人'
                                    ,TB_EB_USER.EMAIL
                                    FROM [UOF].dbo.TB_EIP_SCH_DEVOLVE
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_DEVOLVE_EXAMINE_LOG ON TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_WORK ON TB_EIP_SCH_WORK.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=TB_EIP_SCH_WORK.EXECUTE_USER
                                    LEFT JOIN [UOF].dbo.TB_EB_USER USER2 ON USER2.USER_GUID=TB_EIP_SCH_DEVOLVE.DIRECTOR

                                    WHERE 1=1
                                  
                                    AND TB_EIP_SCH_WORK.WORK_STATE  IN ('NotYetBegin','Proceeding')
                                    AND TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID NOT IN (SELECT [DEVOLVE_GUID]  FROM [UOF].[dbo].[Z_TB_EIP_SCH_DEVOLVE_IGNORES])
                                    

                                    GROUP BY TB_EB_USER.NAME ,TB_EB_USER.EMAIL
                                    ORDER BY TB_EB_USER.NAME ,TB_EB_USER.EMAIL

                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES, "DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES");
                sqlConn.Close();



                if (DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES.Tables["DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES"].Rows.Count > 0)
                {

                    return DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES.Tables["DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES"];
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
        //寄送mail給被交辨人
        public void SEND_EMAIL_TB_EIP_SCH_DEVOLVE(DataTable DT)
        {
            DataTable DTDETAILS = new DataTable();
            DataTable DTDETAILSALL = new DataTable();

            DTDETAILSALL = FIND_TB_EIP_SCH_DEVOLVE_DETAILS_ALL();



            foreach (DataRow DR in DT.Rows)
            {
                string NAME_COMMIT = DR["被交辨人"].ToString();
                // 如果被交辨人中有單引號，應該進行轉義
                NAME_COMMIT = NAME_COMMIT.Replace("'", "''");
                // 建立查詢字串
                string filterExpression = $"被交辨人 = '{NAME_COMMIT}'";
                // 使用 Select 方法查詢
                DataRow[] result = DTDETAILSALL.Select(filterExpression);

                if (result.Length > 0)
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


                        SUBJEST.AppendFormat(@"系統通知-請查收-每日-交辨事項未完成明細，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                        //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                        //ERP 採購相關單別、單號未核準的明細
                        //
                        BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                            + "<br>" + "系統通知-請查收-每日-交辨事項未完成明細，謝謝"
                            + " <br>"
                            );





                        if (result.Length > 0)
                        {
                            BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                            BODY.AppendFormat(@"<table> ");
                            BODY.AppendFormat(@"<tr >");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">逾時天數</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">被交辨人</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨項目</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨回覆狀況</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨預計要求完成日期</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨開始日期</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨人</th>");


                            BODY.AppendFormat(@"</tr> ");

                            foreach (DataRow row in result)
                            {
                                //Console.WriteLine($"ID: {row["ID"]}, Name: {row["Name"]}");


                                BODY.AppendFormat(@"<tr >");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + row["逾時天數"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + row["被交辨人"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + row["交辨項目"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + row["交辨回覆狀況"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + row["交辨預計要求完成日期"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + row["交辨開始日期"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + row["交辨人"].ToString() + "</td>");

                                BODY.AppendFormat(@"</tr> ");

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

                                MyMail.To.Add(DR["EMAIL"].ToString()); //設定收件者Email，多筆mail
                                                                       //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                                MySMTP.Send(MyMail);

                                MyMail.Dispose(); //釋放資源


                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show("有錯誤");

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

            }




            //foreach (DataRow DR in DT.Rows)
            //{
            //    DTDETAILS.Clear();
            //    DTDETAILS = FIND_TB_EIP_SCH_DEVOLVE_DETAILS(DR["被交辨人"].ToString());

            //    if(DTDETAILS!=null && DTDETAILS.Rows.Count>=1)
            //    {
            //        try
            //        {
            //            StringBuilder SUBJEST = new StringBuilder();
            //            StringBuilder BODY = new StringBuilder();

            //            ////加上附圖
            //            //string path = System.Environment.CurrentDirectory+@"/Images/emaillogo.jpg";
            //            //LinkedResource res = new LinkedResource(path);
            //            //res.ContentId = Guid.NewGuid().ToString();

            //            SUBJEST.Clear();
            //            BODY.Clear();


            //            SUBJEST.AppendFormat(@"系統通知-請查收-每日-交辨事項未完成明細，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            //            //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

            //            //ERP 採購相關單別、單號未核準的明細
            //            //
            //            BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
            //                + "<br>" + "系統通知-請查收-每日-交辨事項未完成明細，謝謝"
            //                + " <br>"
            //                );





            //            if (DTDETAILS.Rows.Count > 0)
            //            {
            //                BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

            //                BODY.AppendFormat(@"<table> ");
            //                BODY.AppendFormat(@"<tr >");
            //                BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">逾時天數</th>");
            //                BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">被交辨人</th>");
            //                BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨項目</th>");
            //                BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨回覆狀況</th>");
            //                BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨預計要求完成日期</th>");
            //                BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨開始日期</th>");
            //                BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨人</th>");


            //                BODY.AppendFormat(@"</tr> ");

            //                foreach (DataRow DR_DTDETAILS in DTDETAILS.Rows)
            //                {

            //                    BODY.AppendFormat(@"<tr >");
            //                    BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR_DTDETAILS["逾時天數"].ToString() + "</td>");
            //                    BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR_DTDETAILS["被交辨人"].ToString() + "</td>");
            //                    BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR_DTDETAILS["交辨項目"].ToString() + "</td>");
            //                    BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR_DTDETAILS["交辨回覆狀況"].ToString() + "</td>");
            //                    BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR_DTDETAILS["交辨預計要求完成日期"].ToString() + "</td>");
            //                    BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR_DTDETAILS["交辨開始日期"].ToString() + "</td>");
            //                    BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR_DTDETAILS["交辨人"].ToString() + "</td>");

            //                    BODY.AppendFormat(@"</tr> ");

            //                    //BODY.AppendFormat("<span></span>");
            //                    //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br> " + "品名     " + DR["TD005"].ToString() + "</span>");
            //                    //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購數量 " + DR["TD008"].ToString() + "</span>");
            //                    //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購單位 " + DR["TD009"].ToString() + "</span>");
            //                    //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>");
            //                }
            //                BODY.AppendFormat(@"</table> ");
            //            }


            //            try
            //            {
            //                string MySMTPCONFIG = ConfigurationManager.AppSettings["MySMTP"];
            //                string NAME = ConfigurationManager.AppSettings["NAME"];
            //                string PW = ConfigurationManager.AppSettings["PW"];

            //                System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
            //                MyMail.From = new System.Net.Mail.MailAddress("tk290@tkfood.com.tw");

            //                //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
            //                //MyMail.Subject = "每日訂單-製令追踨表"+DateTime.Now.ToString("yyyy/MM/dd");
            //                MyMail.Subject = SUBJEST.ToString();
            //                //MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
            //                MyMail.Body = BODY.ToString();
            //                MyMail.IsBodyHtml = true; //是否使用html格式

            //                //加上附圖
            //                //string path = System.Environment.CurrentDirectory + @"/Images/emaillogo.jpg";
            //                //MyMail.AlternateViews.Add(GetEmbeddedImage(path, Body));

            //                System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
            //                MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);


            //                try
            //                {

            //                    MyMail.To.Add(DR["EMAIL"].ToString()); //設定收件者Email，多筆mail
            //                    //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
            //                    MySMTP.Send(MyMail);

            //                    MyMail.Dispose(); //釋放資源


            //                }
            //                catch (Exception ex)
            //                {
            //                    //MessageBox.Show("有錯誤");

            //                    //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
            //                    //ex.ToString();
            //                }
            //            }
            //            catch
            //            {

            //            }
            //            finally
            //            {

            //            }


            //        }
            //        catch
            //        {

            //        }
            //        finally
            //        {

            //        }
            //    }
            //}
        }
        //找出被交辨的所有未完成的交辨事項
        public DataTable FIND_TB_EIP_SCH_DEVOLVE_DETAILS(string NAME)
        {

            DataSet DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS = new DataSet();

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

                //   AND DOC_NBR = 'GA1005230100006'

                sbSql.AppendFormat(@"                          
                                
                                    SELECT 
                                    (CASE WHEN  DATEDIFF(DAY, TB_EIP_SCH_WORK.END_TIME, getdate())>0 THEN DATEDIFF(DAY, TB_EIP_SCH_WORK.END_TIME, getdate()) ELSE 0 END) AS '逾時天數'
                                    ,USER2.NAME AS '交辨人'
                                    ,CONVERT(nvarchar,TB_EIP_SCH_WORK.END_TIME,111) AS '交辨預計要求完成日期'
                                    ,CONVERT(nvarchar,TB_EIP_SCH_WORK.CREATE_TIME,111) AS '交辨開始日期'
                                    ,TB_EIP_SCH_DEVOLVE.SUBJECT AS '校稿區內容'
                                    ,TB_EIP_SCH_WORK.SUBJECT AS '交辨項目'
                                    ,TB_EIP_SCH_WORK.EXECUTE_USER AS '被交辨人ID'
                                    ,TB_EIP_SCH_WORK.WORK_STATE AS 'WORK_STATE'
                                    ,(ISNULL(TB_EIP_SCH_WORK.PROCEEDING_DESC,'')+ISNULL(TB_EIP_SCH_WORK.COMPLETE_DESC,''))  AS '交辨回覆'
                                    ,TB_EB_USER.NAME AS '被交辨人'
                                    ,(CASE  WHEN TB_EIP_SCH_WORK.WORK_STATE='Completed' THEN '審稿完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Audit' THEN '交辨完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Proceeding' THEN '處理中' WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN '未開始' END) AS '交辨狀態'
                                    ,(CASE WHEN ISNULL(TB_EIP_SCH_WORK.COMPLETE_TIME,'')<>'' THEN CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,24),1,8) ELSE CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,24),1,8) END)  AS '回覆時間'
                                    ,TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.*
                                    ,TB_EB_USER.ACCOUNT
                                    ,TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID AS 'DEVOLVE_GUID'
                                    ,TB_EB_USER.EMAIL
                                    ,(CASE WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN '未回覆交辨' ELSE '已回覆交辨，但交辨人還未完成' END) AS '交辨回覆狀況'

                                    FROM [UOF].dbo.TB_EIP_SCH_DEVOLVE
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_DEVOLVE_EXAMINE_LOG ON TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_WORK ON TB_EIP_SCH_WORK.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=TB_EIP_SCH_WORK.EXECUTE_USER
                                    LEFT JOIN [UOF].dbo.TB_EB_USER USER2 ON USER2.USER_GUID=TB_EIP_SCH_DEVOLVE.DIRECTOR

                                    WHERE 1=1

                                    AND TB_EIP_SCH_WORK.WORK_STATE  IN ('NotYetBegin','Proceeding')
                                    AND TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID NOT IN (SELECT [DEVOLVE_GUID]  FROM [UOF].[dbo].[Z_TB_EIP_SCH_DEVOLVE_IGNORES])
                                    AND TB_EB_USER.NAME='{0}'
                                    ORDER BY CONVERT(nvarchar,TB_EIP_SCH_WORK.END_TIME,111) 


                                   ", NAME);

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS, "DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS");
                sqlConn.Close();



                if (DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS.Tables["DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS"].Rows.Count > 0)
                {

                    return DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS.Tables["DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS"];
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

        //找出被交辨的所有未完成的交辨事項
        public DataTable FIND_TB_EIP_SCH_DEVOLVE_DETAILS_ALL()
        {

            DataSet DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS = new DataSet();

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

                //   AND DOC_NBR = 'GA1005230100006'

                sbSql.AppendFormat(@"                          
                                
                                    SELECT 
                                    (CASE WHEN  DATEDIFF(DAY, TB_EIP_SCH_WORK.END_TIME, getdate())>0 THEN DATEDIFF(DAY, TB_EIP_SCH_WORK.END_TIME, getdate()) ELSE 0 END) AS '逾時天數'
                                    ,USER2.NAME AS '交辨人'
                                    ,CONVERT(nvarchar,TB_EIP_SCH_WORK.END_TIME,111) AS '交辨預計要求完成日期'
                                    ,CONVERT(nvarchar,TB_EIP_SCH_WORK.CREATE_TIME,111) AS '交辨開始日期'
                                    ,TB_EIP_SCH_DEVOLVE.SUBJECT AS '校稿區內容'
                                    ,TB_EIP_SCH_WORK.SUBJECT AS '交辨項目'
                                    ,TB_EIP_SCH_WORK.EXECUTE_USER AS '被交辨人ID'
                                    ,TB_EIP_SCH_WORK.WORK_STATE AS 'WORK_STATE'
                                    ,(ISNULL(TB_EIP_SCH_WORK.PROCEEDING_DESC,'')+ISNULL(TB_EIP_SCH_WORK.COMPLETE_DESC,''))  AS '交辨回覆'
                                    ,TB_EB_USER.NAME AS '被交辨人'
                                    ,(CASE  WHEN TB_EIP_SCH_WORK.WORK_STATE='Completed' THEN '審稿完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Audit' THEN '交辨完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Proceeding' THEN '處理中' WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN '未開始' END) AS '交辨狀態'
                                    ,(CASE WHEN ISNULL(TB_EIP_SCH_WORK.COMPLETE_TIME,'')<>'' THEN CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,24),1,8) ELSE CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,24),1,8) END)  AS '回覆時間'
                                    ,TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.*
                                    ,TB_EB_USER.ACCOUNT
                                    ,TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID AS 'DEVOLVE_GUID'
                                    ,TB_EB_USER.EMAIL
                                    ,(CASE WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN '未回覆交辨' ELSE '已回覆交辨，但交辨人還未完成' END) AS '交辨回覆狀況'

                                    FROM [UOF].dbo.TB_EIP_SCH_DEVOLVE
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_DEVOLVE_EXAMINE_LOG ON TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_WORK ON TB_EIP_SCH_WORK.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=TB_EIP_SCH_WORK.EXECUTE_USER
                                    LEFT JOIN [UOF].dbo.TB_EB_USER USER2 ON USER2.USER_GUID=TB_EIP_SCH_DEVOLVE.DIRECTOR

                                    WHERE 1=1

                                    AND TB_EIP_SCH_WORK.WORK_STATE  IN ('NotYetBegin','Proceeding')
                                    AND TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID NOT IN (SELECT [DEVOLVE_GUID]  FROM [UOF].[dbo].[Z_TB_EIP_SCH_DEVOLVE_IGNORES])
                                    ORDER BY CONVERT(nvarchar,TB_EIP_SCH_WORK.END_TIME,111) 


                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS, "DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS");
                sqlConn.Close();



                if (DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS.Tables["DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS"].Rows.Count > 0)
                {

                    return DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS.Tables["DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS"];
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

        //交辨未完成meail
        public void CHECK_TB_EIP_SCH_DEVOLVE_MANAGER()
        {
            //找出所有被交辨人的主管
            DataTable DT = FIND_TB_EIP_SCH_DEVOLVE_NAMES_MANAGER();

            if (DT != null && DT.Rows.Count >= 1)
            {
                SEND_EMAIL_TB_EIP_SCH_DEVOLVE_MANAGER(DT);
            }
        }

        //找出交辨的所有 被交辨人
        public DataTable FIND_TB_EIP_SCH_DEVOLVE_NAMES_MANAGER()
        {
            DataSet DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES = new DataSet();

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
                                    Z_TB_EIP_SCH_DEVOLVE_MANAGER.MANAGER AS '被交辨人主管'
                                    ,Z_TB_EIP_SCH_DEVOLVE_MANAGER.MANAGEREMAILS
                                    FROM [UOF].dbo.TB_EIP_SCH_DEVOLVE
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_DEVOLVE_EXAMINE_LOG ON TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_WORK ON TB_EIP_SCH_WORK.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=TB_EIP_SCH_WORK.EXECUTE_USER
                                    LEFT JOIN [UOF].dbo.TB_EB_USER USER2 ON USER2.USER_GUID=TB_EIP_SCH_DEVOLVE.DIRECTOR
                                    LEFT JOIN [UOF].dbo.Z_TB_EIP_SCH_DEVOLVE_MANAGER ON TB_EB_USER.USER_GUID=Z_TB_EIP_SCH_DEVOLVE_MANAGER.ID
                                    WHERE 1=1
                                    --AND TB_EIP_SCH_WORK.SUBJECT  LIKE '%校稿%'
                                    AND TB_EIP_SCH_WORK.WORK_STATE  IN ('NotYetBegin','Proceeding')
                                    AND TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID NOT IN (SELECT [DEVOLVE_GUID]  FROM [UOF].[dbo].[Z_TB_EIP_SCH_DEVOLVE_IGNORES])
                                    AND ISNULL(Z_TB_EIP_SCH_DEVOLVE_MANAGER.MANAGER,'')<>''

                                    AND (CASE WHEN  DATEDIFF(DAY, TB_EIP_SCH_WORK.END_TIME, getdate())>0 THEN DATEDIFF(DAY, TB_EIP_SCH_WORK.END_TIME, getdate()) ELSE 0 END)>=1
                                    GROUP BY Z_TB_EIP_SCH_DEVOLVE_MANAGER.MANAGER,Z_TB_EIP_SCH_DEVOLVE_MANAGER.MANAGEREMAILS
                                    ORDER BY Z_TB_EIP_SCH_DEVOLVE_MANAGER.MANAGER,Z_TB_EIP_SCH_DEVOLVE_MANAGER.MANAGEREMAILS


                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES, "DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES");
                sqlConn.Close();



                if (DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES.Tables["DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES"].Rows.Count > 0)
                {

                    return DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES.Tables["DS_FIND_TB_EIP_SCH_DEVOLVE_NAMES"];
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
        //寄送mail給被交辨人
        public void SEND_EMAIL_TB_EIP_SCH_DEVOLVE_MANAGER(DataTable DT)
        {
            DataTable DTDETAILS = new DataTable();
            DataTable DTDETAILS_ALL = new DataTable();
            DTDETAILS_ALL = FIND_TB_EIP_SCH_DEVOLVE_DETAILS_MANAGER_ALL();

            foreach (DataRow DR in DT.Rows)
            {
                string NAME_COMMIT = DR["被交辨人主管"].ToString();
                // 如果被交辨人中有單引號，應該進行轉義
                NAME_COMMIT = NAME_COMMIT.Replace("'", "''");
                // 建立查詢字串
                string filterExpression = $"被交辨人主管 = '{NAME_COMMIT}'";
                // 使用 Select 方法查詢
                DataRow[] result = DTDETAILS_ALL.Select(filterExpression);


                if (result.Length > 0)
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


                        SUBJEST.AppendFormat(@"系統通知-請查收-每日-交辨事項未完成明細(主管追踨)，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                        //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                        //ERP 採購相關單別、單號未核準的明細
                        //
                        BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                            + "<br>" + "系統通知-請查收-每日-交辨事項未完成明細(主管追踨)，謝謝"
                            + "<br>"
                            );





                        if (result.Length > 0)
                        {
                            BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                            BODY.AppendFormat(@"<table> ");
                            BODY.AppendFormat(@"<tr >");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">逾時天數</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">被交辨人</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨項目</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨項目</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨回覆狀況</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨開始日期</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨人</th>");


                            BODY.AppendFormat(@"</tr> ");

                            foreach (DataRow row in result)
                            {

                                BODY.AppendFormat(@"<tr >");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + row["逾時天數"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + row["被交辨人"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + row["交辨項目"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + row["交辨回覆狀況"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + row["交辨預計要求完成日期"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + row["交辨開始日期"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + row["交辨人"].ToString() + "</td>");

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

                                MyMail.To.Add(DR["MANAGEREMAILS"].ToString()); //設定收件者Email，多筆mail
                                                                               //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                                MySMTP.Send(MyMail);

                                MyMail.Dispose(); //釋放資源


                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show("有錯誤");

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



            }
            //foreach (DataRow DR in DT.Rows)
            //{             

            //    if(!string.IsNullOrEmpty(DR["被交辨人主管"].ToString()))
            //    {
            //        DTDETAILS.Clear();
            //        DTDETAILS = FIND_TB_EIP_SCH_DEVOLVE_DETAILS_MANAGER(DR["被交辨人主管"].ToString());

            //        if (DTDETAILS != null && DTDETAILS.Rows.Count >= 1)
            //        {
            //            try
            //            {
            //                StringBuilder SUBJEST = new StringBuilder();
            //                StringBuilder BODY = new StringBuilder();

            //                ////加上附圖
            //                //string path = System.Environment.CurrentDirectory+@"/Images/emaillogo.jpg";
            //                //LinkedResource res = new LinkedResource(path);
            //                //res.ContentId = Guid.NewGuid().ToString();

            //                SUBJEST.Clear();
            //                BODY.Clear();


            //                SUBJEST.AppendFormat(@"系統通知-請查收-每日-交辨事項未完成明細(主管追踨)，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            //                //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

            //                //ERP 採購相關單別、單號未核準的明細
            //                //
            //                BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
            //                    + "<br>" + "系統通知-請查收-每日-交辨事項未完成明細(主管追踨)，謝謝"
            //                    + " <br>"
            //                    );





            //                if (DTDETAILS.Rows.Count > 0)
            //                {
            //                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

            //                    BODY.AppendFormat(@"<table> ");
            //                    BODY.AppendFormat(@"<tr >");
            //                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">逾時天數</th>");
            //                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">被交辨人</th>");
            //                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨項目</th>");
            //                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨項目</th>");
            //                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨回覆狀況</th>");
            //                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨開始日期</th>");
            //                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">交辨人</th>");


            //                    BODY.AppendFormat(@"</tr> ");

            //                    foreach (DataRow DR_DTDETAILS in DTDETAILS.Rows)
            //                    {

            //                        BODY.AppendFormat(@"<tr >");
            //                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR_DTDETAILS["逾時天數"].ToString() + "</td>");
            //                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR_DTDETAILS["被交辨人"].ToString() + "</td>");
            //                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR_DTDETAILS["交辨項目"].ToString() + "</td>");
            //                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR_DTDETAILS["交辨回覆狀況"].ToString() + "</td>");
            //                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR_DTDETAILS["交辨預計要求完成日期"].ToString() + "</td>");
            //                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR_DTDETAILS["交辨開始日期"].ToString() + "</td>");
            //                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR_DTDETAILS["交辨人"].ToString() + "</td>");

            //                        BODY.AppendFormat(@"</tr> ");

            //                        //BODY.AppendFormat("<span></span>");
            //                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br> " + "品名     " + DR["TD005"].ToString() + "</span>");
            //                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購數量 " + DR["TD008"].ToString() + "</span>");
            //                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購單位 " + DR["TD009"].ToString() + "</span>");
            //                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>");
            //                    }
            //                    BODY.AppendFormat(@"</table> ");
            //                }


            //                try
            //                {
            //                    string MySMTPCONFIG = ConfigurationManager.AppSettings["MySMTP"];
            //                    string NAME = ConfigurationManager.AppSettings["NAME"];
            //                    string PW = ConfigurationManager.AppSettings["PW"];

            //                    System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
            //                    MyMail.From = new System.Net.Mail.MailAddress("tk290@tkfood.com.tw");

            //                    //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
            //                    //MyMail.Subject = "每日訂單-製令追踨表"+DateTime.Now.ToString("yyyy/MM/dd");
            //                    MyMail.Subject = SUBJEST.ToString();
            //                    //MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
            //                    MyMail.Body = BODY.ToString();
            //                    MyMail.IsBodyHtml = true; //是否使用html格式

            //                    //加上附圖
            //                    //string path = System.Environment.CurrentDirectory + @"/Images/emaillogo.jpg";
            //                    //MyMail.AlternateViews.Add(GetEmbeddedImage(path, Body));

            //                    System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
            //                    MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);


            //                    try
            //                    {

            //                        MyMail.To.Add(DR["MANAGEREMAILS"].ToString()); //設定收件者Email，多筆mail
            //                        //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
            //                        MySMTP.Send(MyMail);

            //                        MyMail.Dispose(); //釋放資源


            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        //MessageBox.Show("有錯誤");

            //                        //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
            //                        //ex.ToString();
            //                    }
            //                }
            //                catch
            //                {

            //                }
            //                finally
            //                {

            //                }


            //            }
            //            catch
            //            {

            //            }
            //            finally
            //            {

            //            }
            //        }
            //    }
            //}
        }
        //找出被交辨的所有未完成的交辨事項-經理
        public DataTable FIND_TB_EIP_SCH_DEVOLVE_DETAILS_MANAGER(string NAME)
        {

            DataSet DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS = new DataSet();

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

                //   AND DOC_NBR = 'GA1005230100006'

                sbSql.AppendFormat(@"                          
                                    SELECT 
                                    (CASE WHEN  DATEDIFF(DAY, TB_EIP_SCH_WORK.END_TIME, getdate())>0 THEN DATEDIFF(DAY, TB_EIP_SCH_WORK.END_TIME, getdate()) ELSE 0 END) AS '逾時天數'
                                    ,USER2.NAME AS '交辨人'
                                    ,CONVERT(nvarchar,TB_EIP_SCH_WORK.END_TIME,111) AS '交辨預計要求完成日期'
                                    ,CONVERT(nvarchar,TB_EIP_SCH_WORK.CREATE_TIME,111) AS '交辨開始日期'
                                    ,TB_EIP_SCH_DEVOLVE.SUBJECT AS '校稿區內容'
                                    ,TB_EIP_SCH_WORK.SUBJECT AS '交辨項目'
                                    ,TB_EIP_SCH_WORK.EXECUTE_USER AS '被交辨人ID'
                                    ,TB_EIP_SCH_WORK.WORK_STATE AS 'WORK_STATE'
                                    ,(ISNULL(TB_EIP_SCH_WORK.PROCEEDING_DESC,'')+ISNULL(TB_EIP_SCH_WORK.COMPLETE_DESC,''))  AS '交辨回覆'
                                    ,TB_EB_USER.NAME AS '被交辨人'
                                    ,(CASE  WHEN TB_EIP_SCH_WORK.WORK_STATE='Completed' THEN '審稿完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Audit' THEN '交辨完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Proceeding' THEN '處理中' WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN '未開始' END) AS '交辨狀態'
                                    ,(CASE WHEN ISNULL(TB_EIP_SCH_WORK.COMPLETE_TIME,'')<>'' THEN CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,24),1,8) ELSE CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,24),1,8) END)  AS '回覆時間'
                                    ,TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.*
                                    ,TB_EB_USER.ACCOUNT
                                    ,TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID AS 'DEVOLVE_GUID'
                                    ,TB_EB_USER.EMAIL
                                    ,Z_TB_EIP_SCH_DEVOLVE_MANAGER.MANAGER AS '被交辨人主管'
                                    ,Z_TB_EIP_SCH_DEVOLVE_MANAGER.MANAGEREMAILS
                                    ,(CASE WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN '未回覆交辨' ELSE '已回覆交辨，但交辨人還未完成' END) AS '交辨回覆狀況'

                                    FROM [UOF].dbo.TB_EIP_SCH_DEVOLVE
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_DEVOLVE_EXAMINE_LOG ON TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_WORK ON TB_EIP_SCH_WORK.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=TB_EIP_SCH_WORK.EXECUTE_USER
                                    LEFT JOIN [UOF].dbo.TB_EB_USER USER2 ON USER2.USER_GUID=TB_EIP_SCH_DEVOLVE.DIRECTOR
                                    LEFT JOIN [UOF].dbo.Z_TB_EIP_SCH_DEVOLVE_MANAGER ON TB_EB_USER.USER_GUID=Z_TB_EIP_SCH_DEVOLVE_MANAGER.ID
                                    WHERE 1=1

                                    AND TB_EIP_SCH_WORK.WORK_STATE  IN ('NotYetBegin','Proceeding')
                                    AND TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID NOT IN (SELECT [DEVOLVE_GUID]  FROM [UOF].[dbo].[Z_TB_EIP_SCH_DEVOLVE_IGNORES])
                                    AND (CASE WHEN  DATEDIFF(DAY, TB_EIP_SCH_WORK.END_TIME, getdate())>0 THEN DATEDIFF(DAY, TB_EIP_SCH_WORK.END_TIME, getdate()) ELSE 0 END)>=1
                                    AND Z_TB_EIP_SCH_DEVOLVE_MANAGER.MANAGER='{0}'
                                    ORDER BY CONVERT(nvarchar,TB_EIP_SCH_WORK.END_TIME,111) 

                                   ", NAME);

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS, "DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS");
                sqlConn.Close();



                if (DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS.Tables["DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS"].Rows.Count > 0)
                {

                    return DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS.Tables["DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS"];
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

        //找出被交辨的所有未完成的交辨事項-經理
        public DataTable FIND_TB_EIP_SCH_DEVOLVE_DETAILS_MANAGER_ALL()
        {

            DataSet DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS = new DataSet();

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

                //   AND DOC_NBR = 'GA1005230100006'

                sbSql.AppendFormat(@"                          
                                    SELECT 
                                    (CASE WHEN  DATEDIFF(DAY, TB_EIP_SCH_WORK.END_TIME, getdate())>0 THEN DATEDIFF(DAY, TB_EIP_SCH_WORK.END_TIME, getdate()) ELSE 0 END) AS '逾時天數'
                                    ,USER2.NAME AS '交辨人'
                                    ,CONVERT(nvarchar,TB_EIP_SCH_WORK.END_TIME,111) AS '交辨預計要求完成日期'
                                    ,CONVERT(nvarchar,TB_EIP_SCH_WORK.CREATE_TIME,111) AS '交辨開始日期'
                                    ,TB_EIP_SCH_DEVOLVE.SUBJECT AS '校稿區內容'
                                    ,TB_EIP_SCH_WORK.SUBJECT AS '交辨項目'
                                    ,TB_EIP_SCH_WORK.EXECUTE_USER AS '被交辨人ID'
                                    ,TB_EIP_SCH_WORK.WORK_STATE AS 'WORK_STATE'
                                    ,(ISNULL(TB_EIP_SCH_WORK.PROCEEDING_DESC,'')+ISNULL(TB_EIP_SCH_WORK.COMPLETE_DESC,''))  AS '交辨回覆'
                                    ,TB_EB_USER.NAME AS '被交辨人'
                                    ,(CASE  WHEN TB_EIP_SCH_WORK.WORK_STATE='Completed' THEN '審稿完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Audit' THEN '交辨完成' WHEN TB_EIP_SCH_WORK.WORK_STATE='Proceeding' THEN '處理中' WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN '未開始' END) AS '交辨狀態'
                                    ,(CASE WHEN ISNULL(TB_EIP_SCH_WORK.COMPLETE_TIME,'')<>'' THEN CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.COMPLETE_TIME,24),1,8) ELSE CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,111)+' '+ SUBSTRING(CONVERT(NVARCHAR,TB_EIP_SCH_WORK.PROCEEDING_TIME,24),1,8) END)  AS '回覆時間'
                                    ,TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.*
                                    ,TB_EB_USER.ACCOUNT
                                    ,TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID AS 'DEVOLVE_GUID'
                                    ,TB_EB_USER.EMAIL
                                    ,Z_TB_EIP_SCH_DEVOLVE_MANAGER.MANAGER AS '被交辨人主管'
                                    ,Z_TB_EIP_SCH_DEVOLVE_MANAGER.MANAGEREMAILS
                                    ,(CASE WHEN TB_EIP_SCH_WORK.WORK_STATE='NotYetBegin' THEN '未回覆交辨' ELSE '已回覆交辨，但交辨人還未完成' END) AS '交辨回覆狀況'

                                    FROM [UOF].dbo.TB_EIP_SCH_DEVOLVE
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_DEVOLVE_EXAMINE_LOG ON TB_EIP_SCH_DEVOLVE_EXAMINE_LOG.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EIP_SCH_WORK ON TB_EIP_SCH_WORK.DEVOLVE_GUID=TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID
                                    LEFT JOIN [UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID=TB_EIP_SCH_WORK.EXECUTE_USER
                                    LEFT JOIN [UOF].dbo.TB_EB_USER USER2 ON USER2.USER_GUID=TB_EIP_SCH_DEVOLVE.DIRECTOR
                                    LEFT JOIN [UOF].dbo.Z_TB_EIP_SCH_DEVOLVE_MANAGER ON TB_EB_USER.USER_GUID=Z_TB_EIP_SCH_DEVOLVE_MANAGER.ID
                                    WHERE 1=1

                                    AND TB_EIP_SCH_WORK.WORK_STATE  IN ('NotYetBegin','Proceeding')
                                    AND TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID NOT IN (SELECT [DEVOLVE_GUID]  FROM [UOF].[dbo].[Z_TB_EIP_SCH_DEVOLVE_IGNORES])
                                    AND (CASE WHEN  DATEDIFF(DAY, TB_EIP_SCH_WORK.END_TIME, getdate())>0 THEN DATEDIFF(DAY, TB_EIP_SCH_WORK.END_TIME, getdate()) ELSE 0 END)>=1
                                    AND ISNULL(Z_TB_EIP_SCH_DEVOLVE_MANAGER.MANAGER,'')<>''
                                    ORDER BY CONVERT(nvarchar,TB_EIP_SCH_WORK.END_TIME,111) 

                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS, "DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS");
                sqlConn.Close();



                if (DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS.Tables["DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS"].Rows.Count > 0)
                {

                    return DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS.Tables["DS_FIND_TB_EIP_SCH_DEVOLVE_DETAILS"];
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
        /// 本年新品的銷售報表
        /// </summary>
        public void PREPARESENDEMAIL_NEWSLAES(string path_File)
        {
            SETPATH();

            path_File = path_File + ".xlsx";
            //DATES = DateTime.Now.ToString("yyyyMMdd");
            //DirectoryNAME = @"C:\MQTEMP\" + DATES.ToString() + @"\";
            //path_File_NEWSLAES = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日新品銷售表" + DATES.ToString() + ".pdf";
            //path_File = path_File_NEWSLAES;

            //如果日期資料夾不存在就新增
            if (!Directory.Exists(DirectoryNAME))
            {
                //新增資料夾
                Directory.CreateDirectory(DirectoryNAME);
            }


            // SAVEREPORT_NEWSLAES(path_File_NEWSLAES);

            DataSet DS_NEWSLAES = ERP_NEWSLAES();

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


                SUBJEST.AppendFormat(@"系統通知-老楊食品-本季新品的銷售報表，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                //ERP 採購相關單別、單號未核準的明細
                //
                BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                    + "<br>" + "本季新品的銷售報表表的明細如下"

                    );


                if (DS_NEWSLAES != null && DS_NEWSLAES.Tables[0].Rows.Count > 0)
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                    BODY.AppendFormat(@"<table> ");
                    BODY.AppendFormat(@"<tr >");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品名</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">規格</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">單位</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">新品建立日期</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">第1天業務銷貨日</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">業務銷貨數量</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">業務銷貨金額</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">第1天業務銷退日</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">業務銷退數量</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">業務銷退金額</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">第1天POS銷售日</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">POS銷售數量</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">POS銷售金額</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">總銷售數量</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">總銷售未稅金額</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">平均單位成本</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">總成本</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">總毛利</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">毛利率</th>");

                    BODY.AppendFormat(@"</tr> ");

                    foreach (DataRow DR in DS_NEWSLAES.Tables[0].Rows)
                    {

                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["品號"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["品名"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["規格"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["單位"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["新品建立日期"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["第1天業務銷貨日"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["業務銷貨數量"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["業務銷貨金額"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["第1天業務銷退日"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["業務銷退數量"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["業務銷退金額"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["第1天POS銷售日"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["POS銷售數量"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["POS銷售金額"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["總銷售數量"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["總銷售未稅金額"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["平均單位成本"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["總成本"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["總毛利"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["毛利率"].ToString() + "</td>");

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



                SENDEMAIL_NEWSALES(SUBJEST, BODY, path_File);

            }
            catch
            {

            }
            finally
            {

            }
        }

        /// <summary>
        /// 寄送本年新品的銷售報表
        /// </summary>
        /// <param name="Subject"></param>
        /// <param name="Body"></param>
        public void SENDEMAIL_NEWSALES(StringBuilder Subject, StringBuilder Body, string Attachments)
        {
            try
            {

                DataSet DSFINDPURCHECKMAILTO = FINDPURCHECKMAILTO("NEWSALES");

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

                Attachment attch = new Attachment(Attachments);
                MyMail.Attachments.Add(attch);
                if (DSFINDPURCHECKMAILTO.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow DR in DSFINDPURCHECKMAILTO.Tables[0].Rows)
                    {
                        try
                        {
                            MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail
                                                                  //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email                          
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }

                MySMTP.Send(MyMail);

                MyMail.Dispose(); //釋放資源
            }
            catch
            {

            }
            finally
            {

            }


        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataSet ERP_NEWSLAES()
        {
            DataSet DS_NEWSLAES = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();


            // 取得當前日期和時間
            DateTime currentDate = DateTime.Now;
            // 減去 3 個月的時間
            DateTime threeMonthsAgo = currentDate.AddMonths(-3);
            DateTime firstDay = new DateTime(threeMonthsAgo.Year, threeMonthsAgo.Month, 1);
            DateTime lastDay = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(1).AddDays(-1);
            //DateTime firstDayOfYear = new DateTime(DateTime.Now.Year, 1, 1);
            //DateTime lastDayOfYear = new DateTime(DateTime.Now.Year, 12, 31);



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
                                    SDATES AS '開始日'
                                    ,EDATES AS '結束日'
                                    ,MB001 AS '品號'
                                    ,MB002 AS '品名'
                                    ,MB003 AS '規格'
                                    ,MB004 AS '單位'
                                    ,CREATE_DATE AS '新品建立日期'
                                    ,TOPTG003 AS '第1天業務銷貨日'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTH008)), 1), '.00', '') AS '業務銷貨數量'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTH037)), 1), '.00', '') AS '業務銷貨金額'
                                    ,TOPTI003 AS '第1天業務銷退日'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTJ007)), 1), '.00', '') AS '業務銷退數量'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTJ033)), 1), '.00', '') AS '業務銷退金額'
                                    ,TOPTB001 AS '第1天POS銷售日'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTB019)), 1), '.00', '') AS 'POS銷售數量'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTB031)), 1), '.00', '') AS 'POS銷售金額'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(DECIMAL(16,4),PERCOSTS)), 1), '.00', '') AS '平均單位成本'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(SUMTH008-SUMTJ007+SUMTB019))), 1), '.00', '')  AS '總銷售數量'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(SUMTH037-SUMTJ033+SUMTB031))), 1), '.00', '')  AS '總銷售未稅金額'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(PERCOSTS*(SUMTH008-SUMTJ007+SUMTB019)))), 1), '.00', '')  AS '總成本'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(SUMTH037-SUMTJ033+SUMTB031-(PERCOSTS*(SUMTH008-SUMTJ007+SUMTB019))))), 1), '.00', '')  AS '總毛利'
                                    ,CONVERT(NVARCHAR,CONVERT(DECIMAL(16,2),(CASE WHEN (SUMTH037-SUMTJ033+SUMTB031-(PERCOSTS*(SUMTH008-SUMTJ007+SUMTB019)))<>0 AND (SUMTH037-SUMTJ007+SUMTB031)<>0  THEN (SUMTH037-SUMTJ033+SUMTB031-(PERCOSTS*(SUMTH008-SUMTJ007+SUMTB019)))/(SUMTH037+SUMTB031) ELSE  0 END )*100))+'%'  AS '毛利率'
                                    FROM 
                                    (
                                    SELECT *
                                    ,ISNULL(
                                    (SELECT CASE WHEN SUM(LA024)<>0 AND SUM(LA016)<>0 THEN SUM(LA024)/SUM(LA016) ELSE 0 END
                                    FROM [TK].dbo.SASLA
                                    WHERE LA005=MB001
                                    AND CONVERT(NVARCHAR,LA015,112)>=SDATES
                                    AND CONVERT(NVARCHAR,LA015,112)<='{1}')
                                    ,0) AS PERCOSTS
                                    FROM (
                                    SELECT '{0}' SDATES,'{1}' AS EDATES,MB001,MB002,MB003,MB004,CREATE_DATE
                                    ,ISNULL((SELECT TOP 1 ISNULL(TG003,'') FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG003>='{0}' AND TH004=MB001 ORDER BY TG003 ),'') AS TOPTG003
                                    ,ISNULL((SELECT SUM((CASE WHEN TH009=MD002 THEN ((TH008+TH024)*MD004/MD003) ELSE (TH008+TH024) END)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH LEFT JOIN [TK].dbo.INVMD ON MD001=TH004 WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG003>='{0}' AND TH004=MB001),0) AS SUMTH008
                                    ,ISNULL((SELECT SUM(TH037) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG003>='{0}' AND TH004=MB001),0) AS SUMTH037

                                    ,ISNULL((SELECT TOP 1 ISNULL(TI003,'') FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI003>='{0}' AND TJ004=MB001 ORDER BY TI003 ),'') AS TOPTI003
                                    ,ISNULL((SELECT SUM((CASE WHEN TJ008=MD002 THEN (TJ007*MD004/MD003) ELSE TJ007 END)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ LEFT JOIN [TK].dbo.INVMD ON MD001=TJ004 WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI003>='{0}' AND TJ004=MB001),0) AS SUMTJ007
                                    ,ISNULL((SELECT SUM(TJ033) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI003>='{0}' AND TJ004=MB001),0) AS SUMTJ033

                                    ,ISNULL((SELECT TOP 1 ISNULL(TB001,'') FROM [TK].dbo.POSTB WHERE TB010=MB001 AND TB001>='{0}' ORDER BY TB001),'') AS TOPTB001
                                    ,ISNULL((SELECT SUM(TB019) FROM [TK].dbo.POSTB WHERE TB010=MB001 AND TB001>='{0}'),0) AS SUMTB019
                                    ,ISNULL((SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB010=MB001 AND TB001>='{0}'),0) AS SUMTB031
                                    FROM [TK].dbo.INVMB
                                    WHERE 1=1
                                    AND MB001 LIKE '4%'
                                    AND MB002 NOT LIKE '%試吃%'
                                    AND CREATE_DATE>='{0}'
                                    ) AS TEMP
                                    ) AS TEMP2
                                    ORDER BY (SUMTH037+SUMTB031) DESC



    
                                   ", firstDay.ToString("yyyyMMdd"), lastDay.ToString("yyyyMMdd"));

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_NEWSLAES.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(DS_NEWSLAES, "DS_NEWSLAES");
                sqlConn.Close();



                if (DS_NEWSLAES.Tables["DS_NEWSLAES"].Rows.Count > 0)
                {
                    return DS_NEWSLAES;
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
        public void SETFILE_NEWSLAES(string pathFile)
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


            SEARCH_NEWSLAES(pathFile);

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}

        }

        public void SEARCH_NEWSLAES(string pathFile)
        {
            DataSet ds1 = new DataSet();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            // 取得當前日期和時間
            DateTime currentDate = DateTime.Now;
            // 減去 3 個月的時間
            DateTime threeMonthsAgo = currentDate.AddMonths(-3);
            DateTime firstDay = new DateTime(threeMonthsAgo.Year, threeMonthsAgo.Month, 1);
            DateTime lastDay = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(1).AddDays(-1);
            //DateTime firstDayOfYear = new DateTime(DateTime.Now.Year, 1, 1);
            //DateTime lastDayOfYear = new DateTime(DateTime.Now.Year, 12, 31);

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
                                    MB001 AS '品號'
                                    ,MB002 AS '品名'
                                    ,MB003 AS '規格'
                                    ,MB004 AS '單位'
                                    ,CREATE_DATE AS '新品建立日期'
                                    ,TOPTG003 AS '第1天業務銷貨日'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTH008)), 1), '.00', '') AS '業務銷貨數量'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTH037)), 1), '.00', '') AS '業務銷貨金額'
                                    ,TOPTI003 AS '第1天業務銷退日'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTJ007)), 1), '.00', '') AS '業務銷退數量'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTJ033)), 1), '.00', '') AS '業務銷退金額'
                                    ,TOPTB001 AS '第1天POS銷售日'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTB019)), 1), '.00', '') AS 'POS銷售數量'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTB031)), 1), '.00', '') AS 'POS銷售金額'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(DECIMAL(16,4),PERCOSTS)), 1), '.00', '') AS '平均單位成本'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(SUMTH008-SUMTJ007+SUMTB019))), 1), '.00', '')  AS '總銷售數量'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(SUMTH037-SUMTJ033+SUMTB031))), 1), '.00', '')  AS '總銷售未稅金額'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(PERCOSTS*(SUMTH008-SUMTJ007+SUMTB019)))), 1), '.00', '')  AS '總成本'
                                    ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(SUMTH037-SUMTJ033+SUMTB031-(PERCOSTS*(SUMTH008-SUMTJ007+SUMTB019))))), 1), '.00', '')  AS '總毛利'
                                    ,CONVERT(NVARCHAR,CONVERT(DECIMAL(16,2),(CASE WHEN (SUMTH037-SUMTJ033+SUMTB031-(PERCOSTS*(SUMTH008-SUMTJ007+SUMTB019)))<>0 AND (SUMTH037-SUMTJ007+SUMTB031)<>0  THEN (SUMTH037-SUMTJ033+SUMTB031-(PERCOSTS*(SUMTH008-SUMTJ007+SUMTB019)))/(SUMTH037+SUMTB031) ELSE  0 END )*100))+'%'  AS '毛利率'
                                    FROM 
                                    (
                                    SELECT *
                                    ,ISNULL(
                                    (SELECT CASE WHEN SUM(LA024)<>0 AND SUM(LA016)<>0 THEN SUM(LA024)/SUM(LA016) ELSE 0 END
                                    FROM [TK].dbo.SASLA
                                    WHERE LA005=MB001
                                    AND CONVERT(NVARCHAR,LA015,112)>='{0}'
                                    AND CONVERT(NVARCHAR,LA015,112)<='{1}')
                                    ,0) AS PERCOSTS
                                    FROM (
                                    SELECT '{0}' SDATES,'{1}' AS EDATES,MB001,MB002,MB003,MB004,CREATE_DATE
                                    ,ISNULL((SELECT TOP 1 ISNULL(TG003,'') FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG003>='{0}' AND TH004=MB001 ORDER BY TG003 ),'') AS TOPTG003
                                    ,ISNULL((SELECT SUM((CASE WHEN TH009=MD002 THEN ((TH008+TH024)*MD004/MD003) ELSE (TH008+TH024) END)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH LEFT JOIN [TK].dbo.INVMD ON MD001=TH004 WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG003>='{0}' AND TH004=MB001),0) AS SUMTH008
                                    ,ISNULL((SELECT SUM(TH037) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG003>='{0}' AND TH004=MB001),0) AS SUMTH037

                                    ,ISNULL((SELECT TOP 1 ISNULL(TI003,'') FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI003>='{0}' AND TJ004=MB001 ORDER BY TI003 ),'') AS TOPTI003
                                    ,ISNULL((SELECT SUM((CASE WHEN TJ008=MD002 THEN (TJ007*MD004/MD003) ELSE TJ007 END)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ LEFT JOIN [TK].dbo.INVMD ON MD001=TJ004 WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI003>='{0}' AND TJ004=MB001),0) AS SUMTJ007
                                    ,ISNULL((SELECT SUM(TJ033) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI003>='{0}' AND TJ004=MB001),0) AS SUMTJ033

                                    ,ISNULL((SELECT TOP 1 ISNULL(TB001,'') FROM [TK].dbo.POSTB WHERE TB010=MB001 AND TB001>='{0}' ORDER BY TB001),'') AS TOPTB001
                                    ,ISNULL((SELECT SUM(TB019) FROM [TK].dbo.POSTB WHERE TB010=MB001 AND TB001>='{0}'),0) AS SUMTB019
                                    ,ISNULL((SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB010=MB001 AND TB001>='{0}'),0) AS SUMTB031
                                    FROM [TK].dbo.INVMB
                                    WHERE 1=1
                                    AND MB001 LIKE '4%'
                                    AND MB002 NOT LIKE '%試吃%'
                                    AND ISNULL(MB002,'')<>''
                                    AND CREATE_DATE>='{0}'
                                    ) AS TEMP
                                    ) AS TEMP2
                                    ORDER BY (SUMTH037+SUMTB031) DESC
                                    ", firstDay.ToString("yyyyMMdd"), lastDay.ToString("yyyyMMdd"));

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter1.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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

        public void SETFILE_POSINV(string pathFile)
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


            SEARCH_POSINV(pathFile);

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}

        }

        public void SEARCH_POSINV(string pathFile)
        {
            DataSet ds1 = new DataSet();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DateTime firstDayOfYear = new DateTime(DateTime.Now.Year, 1, 1);
            DateTime lastDayOfYear = new DateTime(DateTime.Now.Year, 12, 31);
            string TODAY = DateTime.Now.ToString("yyyyMMdd");

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
                                   --20230725 查INVLA

                                    SELECT 
                                    LA009 AS '庫別代號'
                                    ,MC002 AS '庫別'
                                    ,LA001 AS '品號'
                                    ,MB002 AS '品名'
                                    ,MB003 AS '規格'
                                    ,MB004 AS '單位'
                                    ,LA016 AS '有效日'
                                    ,NUMS AS '庫存數量'
                                    ,(CASE WHEN  ISDATE(生產日期)=1 THEN 生產日期 WHEN  ISDATE(進貨日期)=1 THEN 進貨日期 WHEN  ISDATE(託外生產日期)=1 THEN 託外生產日期 ELSE 0 END )  AS '生產-進貨日期'
                                    ,(CASE WHEN  ISDATE(生產日期)=1 THEN DATEDIFF(DAY,生產日期,'{0}') WHEN  ISDATE(進貨日期)=1 THEN DATEDIFF(DAY,進貨日期,'{0}') WHEN  ISDATE(託外生產日期)=1 THEN DATEDIFF(DAY,託外生產日期,'{0}') ELSE 0 END ) AS '在倉日期'
                                    ,(DATEDIFF(DAY,'{0}',LA016))  AS '有效天數'
                                    FROM 
                                    (
                                    SELECT LA009,LA001,LA016,SUM(LA005*LA011) AS NUMS
                                    ,ISNULL((SELECT TOP 1 TG040 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016 AND TG022='Y' ORDER BY TF003 ASC),'') AS '生產日期'
                                    ,ISNULL((SELECT TOP 1 TG003 FROM [TK].dbo.PURTG, [TK].dbo.PURTH WHERE TG001=TH001 AND TG002=TH002 AND  TH004=LA001 AND TH010=LA016 AND TG013='Y' ORDER BY TH036 ASC),'') AS '進貨日期'
                                    ,ISNULL((SELECT TOP 1 TH003 FROM [TK].dbo.MOCTH,[TK].dbo.MOCTI WHERE TH001=TI001 AND TH002=TI002 AND TI004=LA001 AND TI010=LA016 AND TI037='Y' ORDER BY TH003 ASC),'') AS '託外生產日期'

                                    FROM [TK].dbo.INVLA 
                                    WHERE  (LA009 IN ('20001','21001','30001','30002','30003','30004')) 
                                    AND( LA001 LIKE '4%' OR LA001 LIKE '5%')
                                    AND ISDATE(LA016)=1
                                    GROUP BY  LA009,LA001,LA016
                                    HAVING SUM(LA005*LA011)>0
                                    ) AS TEMP
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=LA001
                                    LEFT JOIN [TK].dbo.CMSMC ON MC001=LA009
                                    WHERE MB002 NOT LIKE '%暫停%'
                                    ORDER BY LA009,LA001,LA016


                                    ", TODAY);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter1.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;

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
        /// <summary>
        /// 營銷各庫庫存通知
        /// </summary>
        public void PREPARESENDEMAIL_POSINV(string path_File)
        {
            DataSet DS_POSINV = ERP_POSINV();

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


                SUBJEST.AppendFormat(@"系統通知-老楊食品-營銷各庫庫存通知報表，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                //ERP 採購相關單別、單號未核準的明細
                //
                BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                    + "<br>" + "營銷各庫庫存通知的明細如下(含附件)"

                    );


                if (DS_POSINV != null && DS_POSINV.Tables[0].Rows.Count > 0)
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                    BODY.AppendFormat(@"<table> ");
                    BODY.AppendFormat(@"<tr >");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">庫別代號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">庫別</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品號</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品名</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">規格</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">單位</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">有效日</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">庫存數量</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">生產-進貨日期</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">在倉日期</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">有效天數</th>");
                    BODY.AppendFormat(@"</tr> ");

                    foreach (DataRow DR in DS_POSINV.Tables[0].Rows)
                    {

                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["庫別代號"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["庫別"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["品號"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["品名"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["規格"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["單位"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["有效日"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["庫存數量"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["生產-進貨日期"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["在倉日期"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["有效天數"].ToString() + "</td>");
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



                SENDEMAIL_POSINV(SUBJEST, BODY, path_File);

            }
            catch
            {

            }
            finally
            {

            }
        }

        /// <summary>
        /// 寄送營銷各庫庫存通知
        /// </summary>
        /// <param name="Subject"></param>
        /// <param name="Body"></param>
        public void SENDEMAIL_POSINV(StringBuilder Subject, StringBuilder Body, string Attachments)
        {
            DataSet DSMAILTO = FINDPURCHECKMAILTO("POSINV");
            DataSet DSMAILTOCC = FINDPURCHECKMAILTO("POSINVCC");

            try
            {
                if (DSMAILTO.Tables[0].Rows.Count > 0)
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
                        Attachment attch = new Attachment(Attachments + ".xlsx");
                        MyMail.Attachments.Add(attch);

                        //設定收件者Email，多筆mail
                        foreach (DataRow DR in DSMAILTO.Tables[0].Rows)
                        {
                            MyMail.To.Add(DR["MAIL"].ToString());
                        }
                        //設定收件者Email，多筆mail CC
                        foreach (DataRow DR in DSMAILTOCC.Tables[0].Rows)
                        {
                            MyMail.CC.Add(DR["MAIL"].ToString());
                        }

                        //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                        MySMTP.Send(MyMail);

                        MyMail.Dispose(); //釋放資源


                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show("有錯誤");

                        //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                        //ex.ToString();
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

        public DataSet ERP_POSINV()
        {
            DataSet DS_NEWSLAES = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            DateTime firstDayOfYear = new DateTime(DateTime.Now.Year, 1, 1);
            DateTime lastDayOfYear = new DateTime(DateTime.Now.Year, 12, 31);
            string TODAY = DateTime.Now.ToString("yyyyMMdd");

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
                                   --20230725 查INVLA

                                    SELECT 
                                    LA009 AS '庫別代號'
                                    ,MC002 AS '庫別'
                                    ,LA001 AS '品號'
                                    ,MB002 AS '品名'
                                    ,MB003 AS '規格'
                                    ,MB004 AS '單位'
                                    ,LA016 AS '有效日'
                                    ,NUMS AS '庫存數量'
                                    ,(CASE WHEN  ISDATE(生產日期)=1 THEN 生產日期 WHEN  ISDATE(進貨日期)=1 THEN 進貨日期 WHEN  ISDATE(託外生產日期)=1 THEN 託外生產日期 ELSE 0 END )  AS '生產-進貨日期'
                                    ,(CASE WHEN  ISDATE(生產日期)=1 THEN DATEDIFF(DAY,生產日期,'{0}') WHEN  ISDATE(進貨日期)=1 THEN DATEDIFF(DAY,進貨日期,'{0}') WHEN  ISDATE(託外生產日期)=1 THEN DATEDIFF(DAY,託外生產日期,'{0}') ELSE 0 END ) AS '在倉日期'
                                    ,(DATEDIFF(DAY,'{0}',LA016))  AS '有效天數'
                                    FROM 
                                    (
                                    SELECT LA009,LA001,LA016,SUM(LA005*LA011) AS NUMS
                                    ,ISNULL((SELECT TOP 1 TG040 FROM [TK].dbo.MOCTF,[TK].dbo.MOCTG WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016 AND TG022='Y' ORDER BY TF003 ASC),'') AS '生產日期'
                                    ,ISNULL((SELECT TOP 1 TG003 FROM [TK].dbo.PURTG, [TK].dbo.PURTH WHERE TG001=TH001 AND TG002=TH002 AND  TH004=LA001 AND TH010=LA016 AND TG013='Y' ORDER BY TH036 ASC),'') AS '進貨日期'
                                    ,ISNULL((SELECT TOP 1 TH003 FROM [TK].dbo.MOCTH,[TK].dbo.MOCTI WHERE TH001=TI001 AND TH002=TI002 AND TI004=LA001 AND TI010=LA016 AND TI037='Y' ORDER BY TH003 ASC),'') AS '託外生產日期'

                                    FROM [TK].dbo.INVLA 
                                    WHERE  (LA009 IN ('20001','21001','30001','30002','30003','30004')) 
                                    AND( LA001 LIKE '4%' OR LA001 LIKE '5%')
                                    AND ISDATE(LA016)=1
                                    GROUP BY  LA009,LA001,LA016
                                    HAVING SUM(LA005*LA011)>0
                                    ) AS TEMP
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=LA001
                                    LEFT JOIN [TK].dbo.CMSMC ON MC001=LA009
                                    WHERE MB002 NOT LIKE '%暫停%'
                                    ORDER BY LA009,LA001,LA016
                                    ORDER BY LA009,LA001,LA016
                                    ORDER BY LA009,LA001,LA016
                                    ORDER BY LA009,LA001,LA016


                                    ", TODAY);

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_NEWSLAES.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(DS_NEWSLAES, "DS_NEWSLAES");
                sqlConn.Close();



                if (DS_NEWSLAES.Tables["DS_NEWSLAES"].Rows.Count > 0)
                {
                    return DS_NEWSLAES;
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

        public void SETFILE_COPTCD(string pathFile)
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


            SEARCH_COPTCD(pathFile);

            //if (!File.Exists(pathFile + ".xlsx"))
            //{
            //    //SEARCH()

            //}

        }

        public void SEARCH_COPTCD(string pathFile)
        {
            DataSet ds1 = new DataSet();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DateTime FirstDay = DateTime.Now.AddDays(-DateTime.Now.Day + 1);
            //2個月內
            DateTime LastDay = DateTime.Now.AddMonths(2).AddDays(-DateTime.Now.AddMonths(1).Day);
            //本月
            //DateTime LastDay = DateTime.Now.AddMonths(1).AddDays(-DateTime.Now.AddMonths(1).Day);

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
                                    SELECT *
                                    FROM 
                                    (
                                    SELECT '1國內' KINDS,TC001 AS '訂單單別',TC002 AS '訂單單號',MA002 AS '客戶簡稱'
                                    ,CASE WHEN TC016='1' THEN '應稅內含' WHEN TC016='2' THEN '應稅外加' END  AS '課稅別'
                                    ,ME002 AS '部門',TD005 AS '品名',TD008 AS 	'訂單數量',TD009 AS '已交數量',TD024 AS	'贈品數量',TD025 AS	'贈品已交量',(TD008-TD009) AS  '未出數量',TD010 AS 	'單位',TD011 AS  '單價',(TD008-TD009)*TD011 AS '未出貨金額',TD013 AS'預交日'
                                    ,(TD009)*TD011 AS '已出貨金額'                                
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                             
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='1國內')
                                    AND TC001 NOT IN ('A223')

                                    UNION ALL
                                    SELECT '1國內' KINDS,TC001 AS '訂單單別',TC002 AS '訂單單號',MA002 AS '客戶簡稱'
                                    ,CASE WHEN TC016='1' THEN '應稅內含' WHEN TC016='2' THEN '應稅外加' END  AS '課稅別'
                                    ,ME002 AS '部門',TD005 AS '品名',TD008 AS 	'訂單數量',TD009 AS '已交數量',TD024 AS	'贈品數量',TD025 AS	'贈品已交量',(TD008-TD009) AS  '未出數量',TD010 AS 	'單位',TD011 AS  '單價',(TD008-TD009)*TD011 AS '未出貨金額',TD013 AS'預交日'
                                    ,(TD009)*TD011 AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                        
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='1國內')
                                    AND TC001  IN ('A223')
                                    AND TC004 NOT IN ('2248500100')
                                    AND TC004 NOT IN ('2248500100')

                                    UNION ALL
                                    SELECT '1國內' KINDS,'A221' AS '訂單單別','小計' AS '訂單單號','' AS '客戶簡稱'
                                    ,'' AS '課稅別'
                                    ,'' AS '部門','' AS '品名',0 AS 	'訂單數量',0 AS '已交數量',0 AS	'贈品數量',0 AS	'贈品已交量',0 AS  '未出數量','' AS 	'單位',0 AS  '單價',CONVERT(INT,SUM((TD008-TD009)*TD011)) AS '未出貨金額','' AS'預交日'
                                    ,CONVERT(INT,SUM((TD009)*TD011)) AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                             
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='1國內')
                                    AND TC001 NOT IN ('A223')

                                    UNION ALL
                                    SELECT '1國內' KINDS,'A223' AS '訂單單別','小計' AS '訂單單號','' AS '客戶簡稱'
                                    ,'' AS '課稅別'
                                    ,'' AS '部門','' AS '品名',0 AS 	'訂單數量',0 AS '已交數量',0 AS	'贈品數量',0 AS	'贈品已交量',0 AS  '未出數量','' AS 	'單位',0 AS  '單價',CONVERT(INT,SUM((TD008-TD009)*TD011)) AS '未出貨金額','' AS'預交日'
                                    ,CONVERT(INT,SUM((TD009)*TD011)) AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                        
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='1國內')
                                    AND TC001  IN ('A223')
                                    AND TC004 NOT IN ('2248500100')
                                    AND TC004 NOT IN ('2248500100')

                                    UNION ALL
                                    SELECT '1國內' KINDS,'A223' AS '訂單單別','特計' AS '訂單單號',MA002 AS '客戶簡稱'
                                    ,'' AS '課稅別'
                                    ,'' AS '部門','' AS '品名',0 AS 	'訂單數量',0 AS '已交數量',0 AS	'贈品數量',0 AS	'贈品已交量',0 AS  '未出數量','' AS 	'單位',0 AS  '單價',CONVERT(INT,SUM((TD008-TD009)*TD011)) AS '未出貨金額','' AS'預交日'
                                    ,CONVERT(INT,SUM((TD009)*TD011)) AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                        
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='1國內')
                                    AND TC001  IN ('A223')
                                    AND TC004 NOT IN ('2248500100')
                                    AND TC004 NOT IN ('2248500100')
                                    AND MA002 IN ('橘平屋')
                                    GROUP BY MA002

                                    UNION ALL
                                    SELECT '1國內' KINDS,'A223' AS '訂單單別','特計' AS '訂單單號',MA002 AS '客戶簡稱'
                                    ,'' AS '課稅別'
                                    ,'' AS '部門','' AS '品名',0 AS 	'訂單數量',0 AS '已交數量',0 AS	'贈品數量',0 AS	'贈品已交量',0 AS  '未出數量','' AS 	'單位',0 AS  '單價',CONVERT(INT,SUM((TD008-TD009)*TD011)) AS '未出貨金額','' AS'預交日'
                                    ,CONVERT(INT,SUM((TD009)*TD011)) AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                        
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='1國內')
                                    AND TC001  IN ('A223')
                                    AND TC004 NOT IN ('2248500100')
                                    AND TC004 NOT IN ('2248500100')
                                    AND MA002 IN ('大林廠計畫訂單客戶-業務')
                                    GROUP BY MA002

                                    UNION ALL
                                    SELECT '1國內' KINDS,'A223' AS '訂單單別','特計' AS '訂單單號','其他' AS '客戶簡稱'
                                    ,'' AS '課稅別'
                                    ,'' AS '部門','' AS '品名',0 AS 	'訂單數量',0 AS '已交數量',0 AS	'贈品數量',0 AS	'贈品已交量',0 AS  '未出數量','' AS 	'單位',0 AS  '單價',CONVERT(INT,SUM((TD008-TD009)*TD011)) AS '未出貨金額','' AS'預交日'
                                    ,CONVERT(INT,SUM((TD009)*TD011)) AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                        
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='1國內')
                                    AND TC001  IN ('A223')
                                    AND TC004 NOT IN ('2248500100')
                                    AND TC004 NOT IN ('2248500100')
                                    AND MA002 NOT IN ('橘平屋','大林廠計畫訂單客戶-業務')


                                    UNION ALL
                                    SELECT  '2國外' KINDS,TC001 AS '訂單單別',TC002 AS '訂單單號',MA002 AS '客戶簡稱'
                                    ,CASE WHEN TC016='1' THEN '應稅內含' WHEN TC016='2' THEN '應稅外加' END  AS '課稅別'
                                    ,ME002 AS '部門',TD005 AS '品名',TD008 AS 	'訂單數量',TD009 AS '已交數量',TD024 AS	'贈品數量',TD025 AS	'贈品已交量',(TD008-TD009) AS  '未出數量',TD010 AS 	'單位',TD011 AS  '單價',(TD008-TD009)*TD011 AS '未出貨金額',TD013 AS'預交日'
                                    ,(TD009)*TD011 AS '已出貨金額'                                
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                             
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='2國外')
                                    AND TC001 NOT IN ('A223')

                                    UNION ALL
                                    SELECT '2國外' KINDS,TC001 AS '訂單單別',TC002 AS '訂單單號',MA002 AS '客戶簡稱'
                                    ,CASE WHEN TC016='1' THEN '應稅內含' WHEN TC016='2' THEN '應稅外加' END  AS '課稅別'
                                    ,ME002 AS '部門',TD005 AS '品名',TD008 AS 	'訂單數量',TD009 AS '已交數量',TD024 AS	'贈品數量',TD025 AS	'贈品已交量',(TD008-TD009) AS  '未出數量',TD010 AS 	'單位',TD011 AS  '單價',(TD008-TD009)*TD011 AS '未出貨金額',TD013 AS'預交日'
                                    ,(TD009)*TD011 AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                        
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='2國外')
                                    AND TC001  IN ('A223')
                                    AND TC004 NOT IN ('2248500100')
                                    AND TC004 NOT IN ('2248500100')

                                    UNION ALL
                                    SELECT '2國外' KINDS,'A222' AS '訂單單別','小計' AS '訂單單號','' AS '客戶簡稱'
                                    ,'' AS '課稅別'
                                    ,'' AS '部門','' AS '品名',0 AS 	'訂單數量',0 AS '已交數量',0 AS	'贈品數量',0 AS	'贈品已交量',0 AS  '未出數量','' AS 	'單位',0 AS  '單價',CONVERT(INT,SUM((TD008-TD009)*TD011)) AS '未出貨金額','' AS'預交日'
                                    ,CONVERT(INT,SUM((TD009)*TD011)) AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                             
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='2國外')
                                    AND TC001 NOT IN ('A223')


                                    ) AS TEMP 
                                    ORDER BY KINDS,訂單單別,訂單單號,未出貨金額 DESC



                                    ", FirstDay.ToString("yyyyMMdd"), LastDay.ToString("yyyyMMdd"));

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter1.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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

        /// <summary>
        /// 訂單明細及金額報表
        /// </summary>
        public void PREPARESENDEMAIL_COPTCD(string path_File)
        {
            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();
            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"系統通知-老楊食品-訂單明細及金額報表，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                + "<br>" + "訂單明細及金額的明細如附件");
            BODY.AppendFormat(" "
                         + "<br>" + "謝謝"

                         + "</span><br>");



            SENDEMAIL_COPTCD(SUBJEST, BODY, path_File);

            //DataSet DS = ERP_COPTCD();

            //try
            //{
            //    StringBuilder SUBJEST = new StringBuilder();
            //    StringBuilder BODY = new StringBuilder();

            //    ////加上附圖
            //    //string path = System.Environment.CurrentDirectory+@"/Images/emaillogo.jpg";
            //    //LinkedResource res = new LinkedResource(path);
            //    //res.ContentId = Guid.NewGuid().ToString();

            //    SUBJEST.Clear();
            //    BODY.Clear();


            //    SUBJEST.AppendFormat(@"系統通知-老楊食品-訂單明細及金額報表，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            //    //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

            //    //ERP 採購相關單別、單號未核準的明細
            //    //
            //    BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
            //        + "<br>" + "訂單明細及金額的明細如下(含附件)"

            //        );


            //    if (DS != null && DS.Tables[0].Rows.Count > 0)
            //    {
            //        BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

            //        BODY.AppendFormat(@"<table> ");
            //        BODY.AppendFormat(@"<tr >");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">KINDS</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">訂單單別</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">訂單單號</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">客戶簡稱</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">課稅別</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">部門</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品名</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">訂單數量</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">已交數量</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">贈品數量</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">贈品已交量</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">未出數量</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">單位</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">單價</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">未出貨金額</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">預交日</th>");
            //        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">已出貨金額</th>");
            //        BODY.AppendFormat(@"</tr> ");

            //        foreach (DataRow DR in DS.Tables[0].Rows)
            //        {

            //            BODY.AppendFormat(@"<tr >");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["KINDS"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["訂單單別"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["訂單單號"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["客戶簡稱"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["課稅別"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["部門"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["品名"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["訂單數量"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["已交數量"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["贈品數量"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["贈品已交量"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["未出數量"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["單位"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["單價"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["未出貨金額"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["預交日"].ToString() + "</td>");
            //            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["已出貨金額"].ToString() + "</td>");
            //            BODY.AppendFormat(@"</tr> ");

            //            //BODY.AppendFormat("<span></span>");
            //            //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br> " + "品名     " + DR["TD005"].ToString() + "</span>");
            //            //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購數量 " + DR["TD008"].ToString() + "</span>");
            //            //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購單位 " + DR["TD009"].ToString() + "</span>");
            //            //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>");
            //        }
            //        BODY.AppendFormat(@"</table> ");
            //    }
            //    else
            //    {
            //        BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "本日無資料");
            //    }

            //    BODY.AppendFormat(" "
            //                 + "<br>" + "謝謝"

            //                 + "</span><br>");



            //    SENDEMAIL_COPTCD(SUBJEST, BODY, path_File);

            //}
            //catch
            //{

            //}
            //finally
            //{

            //}
        }

        /// <summary>
        /// 訂單明細及金額報表
        /// </summary>
        /// <param name="Subject"></param>
        /// <param name="Body"></param>
        public void SENDEMAIL_COPTCD(StringBuilder Subject, StringBuilder Body, string Attachments)
        {
            DataSet DSMAILTO = FINDPURCHECKMAILTO("COPTCD");

            try
            {
                if (DSMAILTO.Tables[0].Rows.Count > 0)
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
                        Attachment attch = new Attachment(Attachments + ".xlsx");
                        MyMail.Attachments.Add(attch);

                        //設定收件者Email，多筆mail
                        foreach (DataRow DR in DSMAILTO.Tables[0].Rows)
                        {
                            MyMail.To.Add(DR["MAIL"].ToString());
                        }


                        //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                        MySMTP.Send(MyMail);

                        MyMail.Dispose(); //釋放資源


                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show("有錯誤");

                        //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                        //ex.ToString();
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

        public DataSet ERP_COPTCD()
        {
            DataSet DS = new DataSet();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DateTime FirstDay = DateTime.Now.AddDays(-DateTime.Now.Day + 1);
            //2個月內
            DateTime LastDay = DateTime.Now.AddMonths(2).AddDays(-DateTime.Now.AddMonths(1).Day);
            //本月
            //DateTime LastDay = DateTime.Now.AddMonths(1).AddDays(-DateTime.Now.AddMonths(1).Day);


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
                                    SELECT *
                                    FROM 
                                    (
                                    SELECT '1國內' KINDS,TC001 AS '訂單單別',TC002 AS '訂單單號',MA002 AS '客戶簡稱'
                                    ,CASE WHEN TC016='1' THEN '應稅內含' WHEN TC016='2' THEN '應稅外加' END  AS '課稅別'
                                    ,ME002 AS '部門',TD005 AS '品名',TD008 AS 	'訂單數量',TD009 AS '已交數量',TD024 AS	'贈品數量',TD025 AS	'贈品已交量',(TD008-TD009) AS  '未出數量',TD010 AS 	'單位',TD011 AS  '單價',(TD008-TD009)*TD011 AS '未出貨金額',TD013 AS'預交日'
                                    ,(TD009)*TD011 AS '已出貨金額'                                
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                             
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='1國內')
                                    AND TC001 NOT IN ('A223')

                                    UNION ALL
                                    SELECT '1國內' KINDS,TC001 AS '訂單單別',TC002 AS '訂單單號',MA002 AS '客戶簡稱'
                                    ,CASE WHEN TC016='1' THEN '應稅內含' WHEN TC016='2' THEN '應稅外加' END  AS '課稅別'
                                    ,ME002 AS '部門',TD005 AS '品名',TD008 AS 	'訂單數量',TD009 AS '已交數量',TD024 AS	'贈品數量',TD025 AS	'贈品已交量',(TD008-TD009) AS  '未出數量',TD010 AS 	'單位',TD011 AS  '單價',(TD008-TD009)*TD011 AS '未出貨金額',TD013 AS'預交日'
                                    ,(TD009)*TD011 AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                        
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='1國內')
                                    AND TC001  IN ('A223')
                                    AND TC004 NOT IN ('2248500100')
                                    AND TC004 NOT IN ('2248500100')

                                    UNION ALL
                                    SELECT '1國內' KINDS,'A221' AS '訂單單別','小計' AS '訂單單號','' AS '客戶簡稱'
                                    ,'' AS '課稅別'
                                    ,'' AS '部門','' AS '品名',0 AS 	'訂單數量',0 AS '已交數量',0 AS	'贈品數量',0 AS	'贈品已交量',0 AS  '未出數量','' AS 	'單位',0 AS  '單價',CONVERT(INT,SUM((TD008-TD009)*TD011)) AS '未出貨金額','' AS'預交日'
                                    ,CONVERT(INT,SUM((TD009)*TD011)) AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                             
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='1國內')
                                    AND TC001 NOT IN ('A223')

                                    UNION ALL
                                    SELECT '1國內' KINDS,'A223' AS '訂單單別','小計' AS '訂單單號','' AS '客戶簡稱'
                                    ,'' AS '課稅別'
                                    ,'' AS '部門','' AS '品名',0 AS 	'訂單數量',0 AS '已交數量',0 AS	'贈品數量',0 AS	'贈品已交量',0 AS  '未出數量','' AS 	'單位',0 AS  '單價',CONVERT(INT,SUM((TD008-TD009)*TD011)) AS '未出貨金額','' AS'預交日'
                                    ,CONVERT(INT,SUM((TD009)*TD011)) AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                        
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='1國內')
                                    AND TC001  IN ('A223')
                                    AND TC004 NOT IN ('2248500100')
                                    AND TC004 NOT IN ('2248500100')

                                    UNION ALL
                                    SELECT '1國內' KINDS,'A223' AS '訂單單別','特計' AS '訂單單號',MA002 AS '客戶簡稱'
                                    ,'' AS '課稅別'
                                    ,'' AS '部門','' AS '品名',0 AS 	'訂單數量',0 AS '已交數量',0 AS	'贈品數量',0 AS	'贈品已交量',0 AS  '未出數量','' AS 	'單位',0 AS  '單價',CONVERT(INT,SUM((TD008-TD009)*TD011)) AS '未出貨金額','' AS'預交日'
                                    ,CONVERT(INT,SUM((TD009)*TD011)) AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                        
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='1國內')
                                    AND TC001  IN ('A223')
                                    AND TC004 NOT IN ('2248500100')
                                    AND TC004 NOT IN ('2248500100')
                                    AND MA002 IN ('橘平屋')
                                    GROUP BY MA002

                                    UNION ALL
                                    SELECT '1國內' KINDS,'A223' AS '訂單單別','特計' AS '訂單單號',MA002 AS '客戶簡稱'
                                    ,'' AS '課稅別'
                                    ,'' AS '部門','' AS '品名',0 AS 	'訂單數量',0 AS '已交數量',0 AS	'贈品數量',0 AS	'贈品已交量',0 AS  '未出數量','' AS 	'單位',0 AS  '單價',CONVERT(INT,SUM((TD008-TD009)*TD011)) AS '未出貨金額','' AS'預交日'
                                    ,CONVERT(INT,SUM((TD009)*TD011)) AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                        
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='1國內')
                                    AND TC001  IN ('A223')
                                    AND TC004 NOT IN ('2248500100')
                                    AND TC004 NOT IN ('2248500100')
                                    AND MA002 IN ('大林廠計畫訂單客戶-業務')
                                    GROUP BY MA002

                                    UNION ALL
                                    SELECT '1國內' KINDS,'A223' AS '訂單單別','特計' AS '訂單單號','其他' AS '客戶簡稱'
                                    ,'' AS '課稅別'
                                    ,'' AS '部門','' AS '品名',0 AS 	'訂單數量',0 AS '已交數量',0 AS	'贈品數量',0 AS	'贈品已交量',0 AS  '未出數量','' AS 	'單位',0 AS  '單價',CONVERT(INT,SUM((TD008-TD009)*TD011)) AS '未出貨金額','' AS'預交日'
                                    ,CONVERT(INT,SUM((TD009)*TD011)) AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                        
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='1國內')
                                    AND TC001  IN ('A223')
                                    AND TC004 NOT IN ('2248500100')
                                    AND TC004 NOT IN ('2248500100')
                                    AND MA002 NOT IN ('橘平屋','大林廠計畫訂單客戶-業務')


                                    UNION ALL
                                    SELECT  '2國外' KINDS,TC001 AS '訂單單別',TC002 AS '訂單單號',MA002 AS '客戶簡稱'
                                    ,CASE WHEN TC016='1' THEN '應稅內含' WHEN TC016='2' THEN '應稅外加' END  AS '課稅別'
                                    ,ME002 AS '部門',TD005 AS '品名',TD008 AS 	'訂單數量',TD009 AS '已交數量',TD024 AS	'贈品數量',TD025 AS	'贈品已交量',(TD008-TD009) AS  '未出數量',TD010 AS 	'單位',TD011 AS  '單價',(TD008-TD009)*TD011 AS '未出貨金額',TD013 AS'預交日'
                                    ,(TD009)*TD011 AS '已出貨金額'                                
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                             
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='2國外')
                                    AND TC001 NOT IN ('A223')

                                    UNION ALL
                                    SELECT '2國外' KINDS,TC001 AS '訂單單別',TC002 AS '訂單單號',MA002 AS '客戶簡稱'
                                    ,CASE WHEN TC016='1' THEN '應稅內含' WHEN TC016='2' THEN '應稅外加' END  AS '課稅別'
                                    ,ME002 AS '部門',TD005 AS '品名',TD008 AS 	'訂單數量',TD009 AS '已交數量',TD024 AS	'贈品數量',TD025 AS	'贈品已交量',(TD008-TD009) AS  '未出數量',TD010 AS 	'單位',TD011 AS  '單價',(TD008-TD009)*TD011 AS '未出貨金額',TD013 AS'預交日'
                                    ,(TD009)*TD011 AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                        
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='2國外')
                                    AND TC001  IN ('A223')
                                    AND TC004 NOT IN ('2248500100')
                                    AND TC004 NOT IN ('2248500100')

                                    UNION ALL
                                    SELECT '2國外' KINDS,'A222' AS '訂單單別','小計' AS '訂單單號','' AS '客戶簡稱'
                                    ,'' AS '課稅別'
                                    ,'' AS '部門','' AS '品名',0 AS 	'訂單數量',0 AS '已交數量',0 AS	'贈品數量',0 AS	'贈品已交量',0 AS  '未出數量','' AS 	'單位',0 AS  '單價',CONVERT(INT,SUM((TD008-TD009)*TD011)) AS '未出貨金額','' AS'預交日'
                                    ,CONVERT(INT,SUM((TD009)*TD011)) AS '已出貨金額'
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.COPMA,[TK].dbo.CMSME
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MA001
                                    AND TC005=ME001
                                    AND TC027 IN ('Y','N')                             
                                    AND TD013>='{0}' AND TD013<='{1}'
                                    AND TC005 IN (SELECT [DEPNO] FROM [TKBUSINESS].[dbo].[TBREPORTSKINDS] WHERE [KINDS]='2國外')
                                    AND TC001 NOT IN ('A223')


                                    ) AS TEMP 
                                    ORDER BY KINDS,訂單單別,訂單單號,未出貨金額 DESC



                                    ", FirstDay.ToString("yyyyMMdd"), LastDay.ToString("yyyyMMdd"));

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(DS, "DS");
                sqlConn.Close();



                if (DS.Tables["DS"].Rows.Count > 0)
                {
                    return DS;
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

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();

            report1.Load(@"REPORT\國內、外業務部業績日報表V7.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
            table.Connection.CommandTimeout = TIMEOUT_LIMITS;
            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            DateTime now = DateTime.Now;
            // 取得本月第一天日期
            DateTime firstDayOfMonth = new DateTime(now.Year, now.Month, 1);
            // 取得本月最後一天日期
            int daysInMonth = DateTime.DaysInMonth(now.Year, now.Month);
            DateTime lastDayOfMonth = new DateTime(now.Year, now.Month, daysInMonth);


            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                           --20220712 業務員日 報表
                            --20210910 業務員日 報表
                            --200050 張釋予
                            --140078 蔡顏鴻
                            --100005 何姍怡
                            --160155 洪櫻芬
                            --170007 林杏育
                            --120003 葉枋俐
                            SELECT 
                            DATES
                            ,國內張釋予銷貨
                            ,國內張釋予銷退
                            ,國內蔡顏鴻銷貨
                            ,國內蔡顏鴻銷退
                            ,國內何姍怡銷貨
                            ,國內何姍怡銷退
                            ,國內洪櫻芬銷貨
                            ,國內洪櫻芬銷退
                            ,官網及現銷銷貨
                            ,官網及現銷銷退
                            ,全聯銷貨
                            ,國外洪櫻芬銷貨
                            ,國外洪櫻芬銷退
                            ,國外葉枋俐銷貨
                            ,國外葉枋俐銷退
                            ,(國內張釋予銷貨+國內張釋予銷退+國內蔡顏鴻銷貨+國內蔡顏鴻銷退+國內何姍怡銷貨+國內何姍怡銷退+國內洪櫻芬銷貨+國內洪櫻芬銷退+全聯銷貨) AS '國內業務合計'
                            ,(國外洪櫻芬銷貨+國外洪櫻芬銷退+國外葉枋俐銷貨+國外葉枋俐銷退) AS '國外業務合計'
                            ,(國內張釋予銷貨+國內張釋予銷退+國內蔡顏鴻銷貨+國內蔡顏鴻銷退+國內何姍怡銷貨+國內何姍怡銷退+國內洪櫻芬銷貨+國內洪櫻芬銷退+全聯銷貨+國外洪櫻芬銷貨+國外洪櫻芬銷退+國外葉枋俐銷貨+國外葉枋俐銷退) AS '總計'
                            ,(SELECT ISNULL(INTARGETMONEYS,0) FROM [TK].[dbo].[ZTARGETMONEYS] WHERE YEARSMOTNS=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6)) AS '國內月目標業績'
                            ,(SELECT ISNULL([OUTTARGETMONEYS],0) FROM [TK].[dbo].[ZTARGETMONEYS] WHERE YEARSMOTNS=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6)) AS '國外月目標業績'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND SUBSTRING(TG003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6) AND TG023='Y' AND (TG004 LIKE '1%' OR TG004 LIKE '2%' OR TG004 LIKE 'A2%' OR TG004 LIKE 'B2%') AND (TG004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%'))  AND TG006 IN ('200050','140078','100005','160155','170007') ) AS '國內月總銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND SUBSTRING(TI003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6) AND TI019='Y' AND (TI004 LIKE '1%' OR TI004 LIKE '2%' OR TI004 LIKE 'A2%' OR TI004 LIKE 'B2%') AND (TI004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TI006 IN ('200050','140078','100005','160155','170007') ) AS '國內月總銷退'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND SUBSTRING(TG003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6) AND TG023='Y' AND (TG004 LIKE '3%' OR TG004 LIKE 'A3%' OR TG004 LIKE 'B3%') AND TG006 IN ('160155','120003')) AS '國外月總銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND SUBSTRING(TI003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6)  AND TI019='Y'AND (TI004 LIKE '3%' OR TI004 LIKE 'A3%' OR TI004 LIKE 'B3%') AND TI006 IN ('160155','120003')) AS '國外月總銷退'
                            ,(((SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND SUBSTRING(TG003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6) AND TG023='Y' AND (TG004 LIKE '1%' OR TG004 LIKE '2%' OR TG004 LIKE 'A2%' OR TG004 LIKE 'B2%') AND (TG004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%'))  AND TG006 IN ('200050','140078','100005','160155','170007') )+(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND SUBSTRING(TI003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6) AND TI019='Y' AND (TI004 LIKE '1%' OR TI004 LIKE '2%' OR TI004 LIKE 'A2%' OR TI004 LIKE 'B2%') AND (TI004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TI006 IN ('200050','140078','100005','160155','170007') ))/(SELECT ISNULL(INTARGETMONEYS,0) FROM [TK].[dbo].[ZTARGETMONEYS] WHERE YEARSMOTNS=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6))) AS '國內月累績達成率'
                            ,(((SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND SUBSTRING(TG003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6) AND TG023='Y' AND (TG004 LIKE '3%' OR TG004 LIKE 'A3%' OR TG004 LIKE 'B3%') AND TG006 IN ('160155','120003'))+(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND SUBSTRING(TI003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6)  AND TI019='Y'AND (TI004 LIKE '3%' OR TI004 LIKE 'A3%' OR TI004 LIKE 'B3%') AND TI006 IN ('160155','120003')))/(SELECT ISNULL([OUTTARGETMONEYS],0) FROM [TK].[dbo].[ZTARGETMONEYS] WHERE YEARSMOTNS=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6))) AS '國外月累績達成率'
                            FROM (
                            SELECT CONVERT(nvarchar,DATES,112) AS DATES
                            ,[RTSALEMONEYS]  AS '全聯銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG003=CONVERT(nvarchar,DATES,112) AND TG023='Y' AND (TG004 LIKE '1%' OR TG004 LIKE '2%' OR TG004 LIKE 'A2%' OR TG004 LIKE 'B2%') AND (TG004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%'))  AND TG006='200050') AS '國內張釋予銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI003=CONVERT(nvarchar,DATES,112) AND TI019='Y' AND (TI004 LIKE '1%' OR TI004 LIKE '2%' OR TI004 LIKE 'A2%' OR TI004 LIKE 'B2%') AND (TI004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TI006='200050') AS '國內張釋予銷退'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG003=CONVERT(nvarchar,DATES,112) AND TG023='Y' AND (TG004 LIKE '1%' OR TG004 LIKE '2%' OR TG004 LIKE 'A2%' OR TG004 LIKE 'B2%') AND (TG004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TG006='140078') AS '國內蔡顏鴻銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI003=CONVERT(nvarchar,DATES,112) AND TI019='Y' AND (TI004 LIKE '1%' OR TI004 LIKE '2%' OR TI004 LIKE 'A2%' OR TI004 LIKE 'B2%') AND (TI004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TI006='140078') AS '國內蔡顏鴻銷退'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG003=CONVERT(nvarchar,DATES,112) AND TG023='Y' AND (TG004 LIKE '1%' OR TG004 LIKE '2%' OR TG004 LIKE 'A2%' OR TG004 LIKE 'B2%') AND (TG004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TG006='100005') AS '國內何姍怡銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI003=CONVERT(nvarchar,DATES,112) AND TI019='Y' AND (TI004 LIKE '1%' OR TI004 LIKE '2%' OR TI004 LIKE 'A2%' OR TI004 LIKE 'B2%') AND (TI004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TI006='100005') AS '國內何姍怡銷退'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG003=CONVERT(nvarchar,DATES,112) AND TG023='Y' AND (TG004 LIKE '1%' OR TG004 LIKE '2%' OR TG004 LIKE 'A2%' OR TG004 LIKE 'B2%') AND (TG004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TG006='160155') AS '國內洪櫻芬銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI003=CONVERT(nvarchar,DATES,112) AND TI019='Y' AND (TI004 LIKE '1%' OR TI004 LIKE '2%' OR TI004 LIKE 'A2%' OR TI004 LIKE 'B2%') AND (TI004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TI006='160155') AS '國內洪櫻芬銷退'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG003=CONVERT(nvarchar,DATES,112) AND TG023='Y' AND (TG004 LIKE '1%' OR TG004 LIKE '2%' OR TG004 LIKE 'A2%' OR TG004 LIKE 'B2%') AND (TG004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TG001 IN ('A230','A233','A234','A235','A23A','A23E') AND TG006 NOT IN ('200050','140078','100005','160155')) AS '官網及現銷銷貨' 
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI003=CONVERT(nvarchar,DATES,112) AND TI019='Y' AND (TI004 LIKE '1%' OR TI004 LIKE '2%' OR TI004 LIKE 'A2%' OR TI004 LIKE 'B2%') AND (TI004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TI001 IN ('A243','A246','A247','A248','A249') AND TI006 NOT IN ('200050','140078','100005','160155')) AS '官網及現銷銷退' 
                            ,'-' AS '-'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND (CASE WHEN ISDATE(COPTG.UDF01)=1 THEN COPTG.UDF01 ELSE TG003 END =CONVERT(nvarchar,DATES,112)) AND TG023='Y' AND (TG004 LIKE '3%' OR TG004 LIKE 'A3%' OR TG004 LIKE 'B3%') AND TG006='160155') AS '國外洪櫻芬銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI003=CONVERT(nvarchar,DATES,112)  AND TI019='Y'AND (TI004 LIKE '3%' OR TI004 LIKE 'A3%' OR TI004 LIKE 'B3%') AND TI006='160155') AS '國外洪櫻芬銷退'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG003=CONVERT(nvarchar,DATES,112) AND TG023='Y' AND (TG004 LIKE '3%' OR TG004 LIKE 'A3%' OR TG004 LIKE 'B3%') AND TG006='120003') AS '國外葉枋俐銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI003=CONVERT(nvarchar,DATES,112)  AND TI019='Y'AND (TI004 LIKE '3%' OR TI004 LIKE 'A3%' OR TI004 LIKE 'B3%') AND TI006='120003') AS '國外葉枋俐銷退'
                            FROM [TK].dbo.ZDATES
                            WHERE CONVERT(nvarchar,DATES,112)>='{0}' AND CONVERT(nvarchar,DATES,112)<='{1}'
                            ) AS TEMP
                            ORDER BY DATES
                            ", firstDayOfMonth.ToString("yyyyMMdd"), lastDayOfMonth.ToString("yyyyMMdd"));


            return SB;

        }
        public void SENDEMAIL_DAILY_SALES_MONEY()
        {
            DataSet dsSALESMONEYS = new DataSet();
            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();

            SETPATH();

            DATES = DateTime.Now.ToString("yyyyMMdd");
            DirectoryNAME = @"C:\MQTEMP\" + DATES.ToString() + @"\";
            pathFile_SALES_MONEYS = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日業務單位業績日報表" + DATES.ToString() + ".pdf";
            //pathFile_SALES_MONEYS = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日業務單位業績日報表" + DATES.ToString() + ".jpg";

            //如果日期資料夾不存在就新增
            if (!Directory.Exists(DirectoryNAME))
            {
                //新增資料夾
                Directory.CreateDirectory(DirectoryNAME);
            }


            SAVEREPORT(pathFile_SALES_MONEYS);

            dsSALESMONEYS = SERACHMAILSALESMONEYS();

            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"每日-國內外業務業績日報-" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear All, ");
            BODY.AppendFormat(Environment.NewLine + "檢附截至目前各業務每日業績，請參考附件，謝謝");
            BODY.AppendFormat(Environment.NewLine);
            BODY.AppendFormat(Environment.NewLine);
            BODY.AppendFormat(Environment.NewLine + "--");
            BODY.AppendFormat(Environment.NewLine + "業務部 ｜ 連佳瑋");
            BODY.AppendFormat(Environment.NewLine + "");
            BODY.AppendFormat(Environment.NewLine + "622 011 嘉義縣大林鎮大埔美園區五路3號");
            BODY.AppendFormat(Environment.NewLine + "No. 3, Dapumeiyuanqu 5th Rd., Dalin Township, Chiayi County 622 011, Taiwan");
            BODY.AppendFormat(Environment.NewLine + "TEL/ 05-295 6520 #4011    FAX/ 05-295 6519    E-MAIL/ tk660@tkfood.com.tw");
            BODY.AppendFormat(Environment.NewLine + "官網/ www.tkfood.com.tw    FB/ www.facebook.com/tkfood");

            //BODY.AppendFormat("<br /><br />");
            //BODY.AppendFormat("<img src=\"cid:image001\" alt=\"图片描述\" style=\"width:400px;\" />");
            //BODY.AppendFormat("<br /><br />Thank you for your attention.");

            //string emailBody = BODY.ToString();
            //// 创建 HTML 视图
            //AlternateView htmlView = AlternateView.CreateAlternateViewFromString(emailBody, null, MediaTypeNames.Text.Html);

            //// 使用本地图片路径添加图片附件
            //string imagePath = pathFile_SALES_MONEYS;  // 本地图片路径
            //LinkedResource imageResource = new LinkedResource(imagePath, MediaTypeNames.Image.Jpeg);
            //imageResource.ContentId = "image001";  // 设置 Content-ID
            //imageResource.TransferEncoding = TransferEncoding.Base64;
            //htmlView.LinkedResources.Add(imageResource);

            //// 将 HTML 视图添加到邮件
            //System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
            //MyMail.AlternateViews.Add(htmlView);

            string MySMTPCONFIG = ConfigurationManager.AppSettings["MySMTP"];
            string NAME = ConfigurationManager.AppSettings["NAME"];
            string PW = ConfigurationManager.AppSettings["PW"];

            System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
            MyMail.From = new System.Net.Mail.MailAddress("tk660@tkfood.com.tw");

            //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
            //MyMail.Subject = "每日訂單-製令追踨表"+DateTime.Now.ToString("yyyy/MM/dd");
            MyMail.Subject = SUBJEST.ToString();
            //MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
            MyMail.Body = BODY.ToString();
            //MyMail.IsBodyHtml = true; //是否使用html格式

            System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
            MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);

            Attachment attch = new Attachment(pathFile_SALES_MONEYS);
            MyMail.Attachments.Add(attch);


            try
            {
                foreach (DataRow od in dsSALESMONEYS.Tables[0].Rows)
                {

                    MyMail.To.Add(od["MAIL"].ToString()); //設定收件者Email，多筆mail
                }

                //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email

                MySMTP.Send(MyMail);

                MyMail.Dispose(); //釋放資源


            }
            catch (Exception ex)
            {
                ADDLOG(DateTime.Now, SUBJEST.ToString(), ex.ToString());
                //ex.ToString();
            }
        }


        public void SAVEREPORT(string pathFileSALESMONEYS)
        {
            string FILENAME = pathFileSALESMONEYS;
            //string FILENAME = @"C:\MQTEMP\20210915\每日業務單位業績日報表20210915.pdf";
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();

            report1.Load(@"REPORT\國內、外業務部業績日報表V7.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            //adapter1.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
            table.Connection.CommandTimeout = TIMEOUT_LIMITS;
            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));


            // prepare a report
            report1.Prepare();
            // create an instance of HTML export filter
            FastReport.Export.Pdf.PDFExport export = new FastReport.Export.Pdf.PDFExport();
            //FastReport.Export.Image.ImageExport ImageExport = new FastReport.Export.Image.ImageExport();
            // show the export options dialog and do the export
            report1.Export(export, FILENAME);

        }

        public DataSet SERACHMAILSALESMONEYS()
        {
            SqlDataAdapter adapterSALESMONEYS = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilderSALESMONEYS = new SqlCommandBuilder();
            DataSet dsSALESMONEYS = new DataSet();

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
                                    WHERE [SENDTO]='SALESMONEYS'  
                                    ");

                adapterSALESMONEYS = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderSALESMONEYS = new SqlCommandBuilder(adapterSALESMONEYS);
                sqlConn.Open();
                dsSALESMONEYS.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapterSALESMONEYS.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapterSALESMONEYS.Fill(dsSALESMONEYS, "dsSALESMONEYS");
                sqlConn.Close();



                if (dsSALESMONEYS.Tables["dsSALESMONEYS"].Rows.Count >= 1)
                {
                    return dsSALESMONEYS;
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


        public void SETFASTREPORT_QC_CHECK()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL_QC_CEHCK();
            Report report1 = new Report();

            report1.Load(@"REPORT\\溫溼度警報.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["TKA01"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
            table.Connection.CommandTimeout = TIMEOUT_LIMITS;
            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL_QC_CEHCK()
        {
            DateTime now = DateTime.Now;
            now = now.AddDays(-1);
            string SDAYS = now.ToString("yyyyMMdd");


            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                          
                            SELECT CONVERT(NVARCHAR,[開始時間],112) AS '日期'
                            ,[Machine].[區域]
                            ,[alarm_table].[機台名稱],[alarm_table].[警報名稱],COUNT([alarm_table].[NO]) AS '警報次數'
                            ,CONVERT(decimal(16,2),COUNT([alarm_table].[NO])*3/60) AS '警報持續時間(分)'
                            FROM [TK_FOOD].[dbo].[alarm_table]
                            LEFT JOIN [TK_FOOD].[dbo].[Machine] ON [Machine].[機台名稱]= [alarm_table].[機台名稱]
                            WHERE CONVERT(NVARCHAR,[開始時間],112)='{0}'
                            GROUP BY CONVERT(NVARCHAR,[開始時間],112),[Machine].[區域],[alarm_table].[機台名稱],[警報名稱]
                            ORDER BY COUNT([alarm_table].[NO]) DESC

                            ", SDAYS);


            return SB;

        }

        public void SENDEMAIL_DAILY_QC_CHECK()
        {
            DataSet ds = new DataSet();
            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();

            SETPATH();

            DATES = DateTime.Now.ToString("yyyyMMdd");
            DirectoryNAME = @"C:\MQTEMP\" + DATES.ToString() + @"\";
            //pathFile_QC_CHECK = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日溫溼度警報" + DATES.ToString() + ".pdf";
            pathFile_QC_CHECK = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日溫溼度警報" + DATES.ToString() + ".jpg";

            //如果日期資料夾不存在就新增
            if (!Directory.Exists(DirectoryNAME))
            {
                //新增資料夾
                Directory.CreateDirectory(DirectoryNAME);
            }


            SAVEREPORT_QC_CHECK(pathFile_QC_CHECK);

            ds = SERACHMAIL_QC_CHECK();

            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"每日-每日溫溼度警報-" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear All, ");
            BODY.AppendFormat(Environment.NewLine + "檢附每日溫溼度警報，請參考附件，謝謝");
            BODY.AppendFormat("<br /><br />");
            BODY.AppendFormat("<img src=\"cid:image001\" alt=\"图片描述\" style=\"width:400px;\" />");
            BODY.AppendFormat("<br /><br />Thank you for your attention.");

            string emailBody = BODY.ToString();
            // 创建 HTML 视图
            AlternateView htmlView = AlternateView.CreateAlternateViewFromString(emailBody, null, MediaTypeNames.Text.Html);

            // 使用本地图片路径添加图片附件
            string imagePath = pathFile_QC_CHECK;  // 本地图片路径
            LinkedResource imageResource = new LinkedResource(imagePath, MediaTypeNames.Image.Jpeg);
            imageResource.ContentId = "image001";  // 设置 Content-ID
            imageResource.TransferEncoding = TransferEncoding.Base64;
            htmlView.LinkedResources.Add(imageResource);

            // 将 HTML 视图添加到邮件
            System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
            MyMail.AlternateViews.Add(htmlView);

            string MySMTPCONFIG = ConfigurationManager.AppSettings["MySMTP"];
            string NAME = ConfigurationManager.AppSettings["NAME"];
            string PW = ConfigurationManager.AppSettings["PW"];

           
            MyMail.From = new System.Net.Mail.MailAddress("tk290@tkfood.com.tw");

            //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
            //MyMail.Subject = "每日訂單-製令追踨表"+DateTime.Now.ToString("yyyy/MM/dd");
            MyMail.Subject = SUBJEST.ToString();
            //MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
            MyMail.Body = BODY.ToString();
            //MyMail.IsBodyHtml = true; //是否使用html格式

            System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
            MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);

            Attachment attch = new Attachment(pathFile_QC_CHECK);
            MyMail.Attachments.Add(attch);


            try
            {
                foreach (DataRow od in ds.Tables[0].Rows)
                {

                    MyMail.To.Add(od["MAIL"].ToString()); //設定收件者Email，多筆mail
                }

                //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email

                MySMTP.Send(MyMail);

                MyMail.Dispose(); //釋放資源


            }
            catch (Exception ex)
            {
                ADDLOG(DateTime.Now, SUBJEST.ToString(), ex.ToString());
                //ex.ToString();
            }
        }


        public void SAVEREPORT_QC_CHECK(string pathFile)
        {
            string FILENAME = pathFile;
            //string FILENAME = @"C:\MQTEMP\20210915\每日業務單位業績日報表20210915.pdf";
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL_QC_CEHCK();
            Report report1 = new Report();

            report1.Load(@"REPORT\溫溼度警報.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["TKA01"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
            table.Connection.CommandTimeout = TIMEOUT_LIMITS;
            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));


            // prepare a report
            report1.Prepare();
            // create an instance of HTML export filter
            //FastReport.Export.Pdf.PDFExport export = new FastReport.Export.Pdf.PDFExport();
            FastReport.Export.Image.ImageExport ImageExport = new FastReport.Export.Image.ImageExport();
            // show the export options dialog and do the export
            report1.Export(ImageExport, FILENAME);

        }

        public DataSet SERACHMAIL_QC_CHECK()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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
                                    WHERE [SENDTO]='QCCHECK'  
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds;
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

        public void SAVEREPORT_NEWSLAES(string path_File_NEWSLAES)
        {
            string FILENAME = path_File_NEWSLAES;
            //string FILENAME = @"C:\MQTEMP\20210915\每日業務單位業績日報表20210915.pdf";
            StringBuilder SQL1 = new StringBuilder();

            StringBuilder SB = new StringBuilder();
            string SDAYS = DateTime.Now.ToString("yyyy") + "0101";
            string EDAYS = DateTime.Now.ToString("yyyyMMdd");
            string ADDYEARSs = DateTime.Now.ToString("yyyy");
            string ADD_DAYS = DateTime.Now.ToString("yyyy") + "0101";

            SQL1.AppendFormat(@"                               
                            SELECT  
                            MB001 AS '品號'
                            ,MB002 AS '品名'
                            ,MB003 AS '規格'
                            ,MB004 AS '單位'
                            ,CREATE_DATE AS '新品建立日期'
                            ,TOPTG003 AS '第1天業務銷貨日'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTH008)), 1), '.00', '') AS '累計-業務銷貨數量'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTH037)), 1), '.00', '') AS '累計-業務銷貨金額'
                            ,TOPTI003 AS '第1天業務銷退日'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTJ007)), 1), '.00', '') AS '累計-業務銷退數量'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTJ033)), 1), '.00', '') AS '累計-業務銷退金額'
                            ,TOPTB001 AS '第1天POS銷售日'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTB019)), 1), '.00', '') AS '累計-POS銷售數量'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTB031)), 1), '.00', '') AS '累計-POS銷售金額'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(DECIMAL(16,4),單位成本)), 1), '.00', '') AS '平均單位成本'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(SUMTH008-SUMTJ007+SUMTB019))), 1), '.00', '')  AS '累計-總銷售數量'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(SUMTH037-SUMTJ033+SUMTB031))), 1), '.00', '')  AS '累計-總銷售未稅金額'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(單位成本*(SUMTH008-SUMTJ007+SUMTB019)))), 1), '.00', '')  AS '累計-總成本'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(SUMTH037-SUMTJ033+SUMTB031-(單位成本*(SUMTH008-SUMTJ007+SUMTB019))))), 1), '.00', '')  AS '累計-總毛利'
                            ,CONVERT(NVARCHAR,CONVERT(DECIMAL(16,2),(CASE WHEN (SUMTH037-SUMTJ033+SUMTB031-(單位成本*(SUMTH008-SUMTJ007+SUMTB019)))<>0 AND (SUMTH037-SUMTJ007+SUMTB031)<>0  THEN (SUMTH037-SUMTJ033+SUMTB031-(單位成本*(SUMTH008-SUMTJ007+SUMTB019)))/(SUMTH037+SUMTB031) ELSE  0 END )*100))+'%'  AS '累計-毛利率'
                            FROM 
                            (
                            SELECT *
                            ,ISNULL(
                            (SELECT CASE WHEN SUM(LA024)<>0 AND SUM(LA016)<>0 THEN SUM(LA024)/SUM(LA016) ELSE 0 END
                            FROM [TK].dbo.SASLA
                            WHERE LA005=MB001
                            AND CONVERT(NVARCHAR,LA015,112)>='{0}'
                            AND CONVERT(NVARCHAR,LA015,112)<='{1}')
                            ,0) AS PERCOSTS
                            FROM (
                            SELECT '{0}' SDATES,'{1}' AS EDATES,MB001,MB002,MB003,MB004,CREATE_DATE
                            ,ISNULL((SELECT TOP 1 ISNULL(TG003,'') FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG003>='{0}' AND TH004=MB001 ORDER BY TG003 ),'') AS TOPTG003
                            ,ISNULL((SELECT SUM((CASE WHEN TH009=MD002 THEN ((TH008+TH024)*MD004/MD003) ELSE (TH008+TH024) END)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH LEFT JOIN [TK].dbo.INVMD ON MD001=TH004  AND TH009=MD002  WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG003>='{0}'  AND TG003<='{1}' AND TH004=MB001),0) AS SUMTH008
                            ,ISNULL((SELECT SUM(TH037) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG003>='{0}'  AND TG003<='{1}'  AND TH004=MB001),0) AS SUMTH037

                            ,ISNULL((SELECT TOP 1 ISNULL(TI003,'') FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI003>='{0}' AND TJ004=MB001 ORDER BY TI003 ),'') AS TOPTI003
                            ,ISNULL((SELECT SUM((CASE WHEN TJ008=MD002 THEN (TJ007*MD004/MD003) ELSE TJ007 END)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ LEFT JOIN [TK].dbo.INVMD ON MD001=TJ004  AND TJ008=MD002 WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI003>='{0}'  AND TI003<='{1}'  AND TJ004=MB001),0) AS SUMTJ007
                            ,ISNULL((SELECT SUM(TJ033) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI003>='{0}'  AND TI003<='{1}' AND TJ004=MB001),0) AS SUMTJ033

                            ,ISNULL((SELECT TOP 1 ISNULL(TB001,'') FROM [TK].dbo.POSTB WHERE TB010=MB001 AND TB001>='{0}' ORDER BY TB001),'') AS TOPTB001
                            ,ISNULL((SELECT SUM(TB019) FROM [TK].dbo.POSTB WHERE TB010=MB001 AND TB001>='{0}' AND TB001<='{1}' ),0) AS SUMTB019
                            ,ISNULL((SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB010=MB001 AND TB001>='{0}' AND TB001<='{1}'),0) AS SUMTB031
                            FROM [TK].dbo.INVMB
                            WHERE 1=1
                            AND MB001 LIKE '4%'
                            AND MB002 NOT LIKE '%試吃%'
                            AND MB002 NOT LIKE '%空%'
                            AND ISNULL(MB002,'')<>''
                            AND CREATE_DATE>='{2}'
                            ) AS TEMP
                            LEFT JOIN
                            (
                            SELECT *
                            FROM 
                            (
                            SELECT TA001 AS '品號',MB002 AS '品名',MB003 AS '規格',MB004 AS '單位'
                            ,CONVERT(DECIMAL(16,2),AVG((ME007+ME008+ME009+ME010)/(生產入庫數+ME005))) 單位成本
                            ,CONVERT(DECIMAL(16,2),AVG((ME007)/(生產入庫數+ME005))) 單位材料成本, CONVERT(DECIMAL(16,2),AVG((ME008)/(生產入庫數+ME005))) 單位人工成本,CONVERT(DECIMAL(16,2),AVG((ME009)/(生產入庫數+ME005))) 單位製造成本,CONVERT(DECIMAL(16,2),AVG((ME010)/(生產入庫數+ME005))) 單位加工成本
                            FROM 
                            (
                            SELECT TA002,TA001,SUM(TA012) '生產入庫數',SUM(TA016-TA019) AS '本階人工成本',SUM(TA017-TA020) AS '本階製造費用'
                            FROM [TK].dbo.CSTTA
                            WHERE TA002 LIKE '{3}%'
                            GROUP BY TA002,TA001
                            ) AS TEMP
                            LEFT JOIN [TK].dbo.CSTME ON ME001=TA001 AND ME002=TA002
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TA001
                            WHERE 1=1
                            AND (生產入庫數+ME005)>0      
                            GROUP BY  TA001 ,MB002 ,MB003,MB004                                 
                            ) AS TEMP2
                            ) AS TEMP3 ON TEMP3.品號=TEMP.MB001
                            ) AS TEMP2
                            ORDER BY (SUMTH037-SUMTJ033+SUMTB031) DESC,新品建立日期
                            ", SDAYS, EDAYS, ADD_DAYS, ADDYEARSs);
            Report report1 = new Report();

            report1.Load(@"REPORT\新品銷售資料.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
            table.Connection.CommandTimeout = TIMEOUT_LIMITS;
            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));


            // prepare a report
            report1.Prepare();
            // create an instance of HTML export filter
            FastReport.Export.Pdf.PDFExport export = new FastReport.Export.Pdf.PDFExport();
            // show the export options dialog and do the export
            report1.Export(export, FILENAME);

        }

        /// <summary>
        /// 更新已進貨的數量，用驗收數量>TOTALNUMS
        /// </summary>
        public void UDPATE_PURVERSIONSNUMS_TOTALNUMS()
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

                sbSql.AppendFormat(@"  
                                    UPDATE [TKPUR].[dbo].[PURVERSIONSNUMS]
                                    SET [TOTALNUMS]=(SELECT SUM(TH015) FROM[TK].dbo.PURTH,[TK].dbo.PURTG WHERE TG001=TH001 AND TG002=TH002 AND TG013='Y' AND TH004=MB001 ) 
                                    WHERE [TOTALNUMS]<>(SELECT SUM(TH015) FROM[TK].dbo.PURTH,[TK].dbo.PURTG WHERE TG001=TH001 AND TG002=TH002 AND TG013='Y' AND TH004=MB001 ) 

                                    ");

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

        public void NEW_PURVERSIONSNUMS()
        {
            try
            {
                DataTable DT = SEARCH_PURVERSIONSNUMS();
                if (DT != null && DT.Rows.Count >= 1)
                {
                    SEND_PURVERSIONSNUMS(DT);
                }
            }
            catch
            {

            }
            finally { }

        }

        public DataTable SEARCH_PURVERSIONSNUMS()
        {
            DataTable DT = new DataTable();
            SqlDataAdapter Adapter1 = new SqlDataAdapter();
            SqlCommandBuilder SqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet DS = new DataSet();

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
                                    SELECT *
                                    FROM [TKPUR].[dbo].[PURVERSIONSNUMS]
                                    WHERE ISCLOSE='N'
                                    AND TOTALNUMS>=TARGETNUMS 
                                    ");

                Adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                SqlCmdBuilder1 = new SqlCommandBuilder(Adapter1);
                sqlConn.Open();
                DS.Clear();
                // 設置查詢的超時時間，以秒為單位
                Adapter1.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                Adapter1.Fill(DS, "DS");
                sqlConn.Close();



                if (DS.Tables["DS"].Rows.Count >= 1)
                {
                    return DS.Tables["DS"];
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

        public void SEND_PURVERSIONSNUMS(DataTable DT)
        {
            try
            {
                if (DT != null && DT.Rows.Count >= 1)
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


                        SUBJEST.AppendFormat(@"系統通知-老楊食品-版費可退回明細 ，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                        //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                        //ERP 採購相關單別、單號未核準的明細
                        //
                        BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                            + "<br>" + "版費可退回明細如下"

                            );


                        if (DT != null && DT.Rows.Count > 0)
                        {
                            BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                            BODY.AppendFormat(@"<table> ");
                            BODY.AppendFormat(@"<tr >");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">版型</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品號</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品名</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">可退還的版費</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">目標進貨量</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">已進貨量</th>");
                            BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">是否結案</th>");

                            BODY.AppendFormat(@"</tr> ");

                            foreach (DataRow DR in DT.Rows)
                            {

                                BODY.AppendFormat(@"<tr >");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["NAMES"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["MB001"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["MB002"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["BACKMONEYS"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TARGETNUMS"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TOTALNUMS"].ToString() + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["ISCLOSE"].ToString() + "</td>");

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



                        SENDE_TO_PURTYPES(SUBJEST, BODY);

                    }
                    catch
                    {

                    }
                    finally
                    {

                    }
                }
            }
            catch
            {

            }
            finally { }
        }

        public void SENDEMAIL_DAILY_TKWH_CALENDAR()
        {
            DataSet DS_EMAIL_CALENDAR = new DataSet();
            DataTable DT_CALENDAR1 = new DataTable();
            DataTable DT_CALENDAR2 = new DataTable();
            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();

            SETPATH();

            DATES = DateTime.Now.ToString("yyyyMMdd");

            DS_EMAIL_CALENDAR = SERACH_MAIL_CALENDAR();

            SUBJEST.Clear();
            BODY.Clear();

            //// 创建一个DataTable来存储日期和内容
            //DataTable eventsTable = new DataTable();
            //eventsTable.Columns.Add("Date", typeof(DateTime));
            //eventsTable.Columns.Add("Event", typeof(string));

            //// 假设这是从某个数据源获取的日期和内容，这里只是简单地手动添加了一些示例数据
            //eventsTable.Rows.Add(new DateTime(2024, 4, 16), "会议");
            //eventsTable.Rows.Add(new DateTime(2024, 4, 17), "生日聚会1");
            //eventsTable.Rows.Add(new DateTime(2024, 4, 17), "生日聚会2");
            //eventsTable.Rows.Add(new DateTime(2024, 4, 18), "项目截止日期");

            // 创建电子邮件消息
            //MailMessage mail = new MailMessage();
            //mail.From = new MailAddress(senderEmail);
            //mail.To.Add(new MailAddress(recipientEmail));

            SUBJEST.AppendFormat(@"系統通知-每日派車- {0}", DATES);

            // 构建HTML内容
            StringBuilder htmlBody = new StringBuilder();
            htmlBody.Append("<html><body>");
            htmlBody.Append("<h1>每日派車行事歷</h1>");

            //改成2個月
            // 获取指定月份的第一天和最后一天
            for (int MONTHSCOUNT = 0; MONTHSCOUNT <= 1; MONTHSCOUNT++)
            {
                DateTime firstDayOfMonth1 = new DateTime(DateTime.Now.Year, DateTime.Now.Month + MONTHSCOUNT, 1);
                DateTime lastDayOfMonth1 = firstDayOfMonth1.AddMonths(1).AddDays(-1);

                //第1個月
                DT_CALENDAR1 = SERACH_DT_CALENDAR(firstDayOfMonth1.ToString("yyyyMMdd"), lastDayOfMonth1.ToString("yyyyMMdd"));


                // 确定第一个星期一的日期
                DateTime firstMonday = firstDayOfMonth1.AddDays((7 - (int)firstDayOfMonth1.DayOfWeek + (int)DayOfWeek.Monday) % 7);


                htmlBody.Append("<table border='1' cellpadding='5' cellspacing='0'>");

                // 添加固定的表头，从星期一到星期日
                htmlBody.Append("<tr>");
                for (int i = 0; i < 7; i++)
                {
                    htmlBody.Append("<th>" + firstMonday.AddDays(i).ToString("ddd") + "</th>");
                }
                htmlBody.Append("</tr>");

                // 计算第一周之前的日期
                DateTime currentDay = firstMonday;

                DayOfWeek dayOfWeek = firstDayOfMonth1.DayOfWeek;
                //int dayOfWeekNumber = (int)dayOfWeek-2;
                int dayOfWeekNumber = ((int)firstDayOfMonth1.DayOfWeek - 2 + 7) % 7;
                while (dayOfWeekNumber >= 0 && dayOfWeekNumber <= 5)
                {
                    htmlBody.Append("<td></td>"); // 空单元格
                    dayOfWeekNumber--;
                }


                // 遍历当前月份的每一天
                for (int day = 1; day <= lastDayOfMonth1.Day; day++)
                {
                    // 检查日期是否在DataTable中存在对应的内容
                    DateTime currentDate = new DateTime(firstDayOfMonth1.Year, firstDayOfMonth1.Month, day);

                    if (DT_CALENDAR1 != null && DT_CALENDAR1.Rows.Count >= 1)
                    {
                        var rows = DT_CALENDAR1.AsEnumerable().Where(row => row.Field<string>("EVENTDATE") == currentDate.ToString("yyyyMMdd"));
                        List<string> events = rows.Select(row => row.Field<string>("EVENTS")).ToList();
                        string eventText = string.Join("<br>", events);

                        // 每周开始时添加新行
                        if (currentDate.DayOfWeek == DayOfWeek.Monday)
                        {
                            htmlBody.Append("</tr><tr>");
                        }

                        // 添加单元格
                        htmlBody.Append("<td valign='top'>" + currentDate.Month + "/" + currentDate.Day + "<br>" + eventText + "</td>");
                    }

                }

                // 补齐最后一周的空单元格
                while (currentDay.DayOfWeek != DayOfWeek.Monday)
                {
                    htmlBody.Append("<td></td>");
                    currentDay = currentDay.AddDays(1);
                }

                htmlBody.Append("</tr></table>");

                htmlBody.Append("<br><br>");
            }




            htmlBody.Append("</body></html>");

            //mail.Body = htmlBody.ToString();
            //mail.IsBodyHtml = true;

            string MySMTPCONFIG = ConfigurationManager.AppSettings["MySMTP"];
            string NAME = ConfigurationManager.AppSettings["NAME"];
            string PW = ConfigurationManager.AppSettings["PW"];

            System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
            MyMail.From = new System.Net.Mail.MailAddress("tk290@tkfood.com.tw");

            //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
            //MyMail.Subject = "每日訂單-製令追踨表"+DateTime.Now.ToString("yyyy/MM/dd");
            MyMail.Subject = SUBJEST.ToString();
            //MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
            MyMail.Body = htmlBody.ToString();
            MyMail.IsBodyHtml = true; //是否使用html格式

            System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
            MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);

            //Attachment attch = new Attachment(pathFile_SALES_MONEYS);
            //MyMail.Attachments.Add(attch);


            try
            {
                //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email，多筆mail

                foreach (DataRow od in DS_EMAIL_CALENDAR.Tables[0].Rows)
                {

                    MyMail.To.Add(od["MAIL"].ToString()); //設定收件者Email，多筆mail
                }

                MySMTP.Send(MyMail);

                MyMail.Dispose(); //釋放資源


            }
            catch (Exception ex)
            {
                ADDLOG(DateTime.Now, SUBJEST.ToString(), ex.ToString());
                //ex.ToString();
            }
        }

        public DataTable SERACH_DT_CALENDAR(string SDAYS, string EDAYS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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
                                   SELECT
                                     [ID]
                                    , CONVERT(NVARCHAR,[EVENTDATE],112) AS EVENTDATE
                                    ,[CAR]
                                    ,[EVENT]
                                    ,[CAR]+':'+[EVENT] AS 'EVENTS'
                                    FROM [TKWAREHOUSE].[dbo].[CALENDAR]
                                    WHERE CONVERT(NVARCHAR,[EVENTDATE],112)>='{0}' AND CONVERT(NVARCHAR,[EVENTDATE],112)<='{1}'
                                    ORDER BY [EVENTDATE],[ID]
                                    ", SDAYS, EDAYS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"];
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

        public DataSet SERACH_MAIL_CALENDAR()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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
                                    WHERE [SENDTO]='CALENDAR'  
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds;
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

        public void SENDEMAIL_TB_SALES_PROMOTIONS()
        {
            DataTable DS_EMAIL_TO_EMAIL = new DataTable();
            DataTable DT_TB_SALES_PROMOTIONS = new DataTable();

            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();

            try
            {
                DS_EMAIL_TO_EMAIL = SERACH_MAIL_TB_SALES_PROMOTIONS();
                DT_TB_SALES_PROMOTIONS = SERACH_TB_SALES_PROMOTIONS();

                if (DT_TB_SALES_PROMOTIONS != null && DT_TB_SALES_PROMOTIONS.Rows.Count >= 1)
                {
                    SUBJEST.Clear();
                    BODY.Clear();


                    SUBJEST.AppendFormat(@"系統通知-請查收-每週-業務活動記錄，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                    //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                    //ERP 採購相關單別、單號未核準的明細
                    //
                    BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                        + "<br>" + "系統通知-請查收-每週-業務活動記錄，謝謝"
                        + " <br>"
                        );





                    if (DT_TB_SALES_PROMOTIONS.Rows.Count > 0)
                    {
                        BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                        BODY.AppendFormat(@"<table> ");
                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">通路</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">活動時間</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">產品規格</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">出貨日</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">活動類型</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">活動內容及價格</th>");


                        BODY.AppendFormat(@"</tr> ");

                        foreach (DataRow DR in DT_TB_SALES_PROMOTIONS.Rows)
                        {

                            BODY.AppendFormat(@"<tr >");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["SALESTO"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["SDATES"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["PRODUCTS"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["SHIPDATES"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["KINDS"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["CONTEXTS"].ToString() + "</td>");

                            BODY.AppendFormat(@"</tr> ");


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
                            foreach (DataRow DR in DS_EMAIL_TO_EMAIL.Rows)
                            {
                                MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail
                            }

                            //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                            MySMTP.Send(MyMail);

                            MyMail.Dispose(); //釋放資源

                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show("有錯誤");

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



            }
            catch
            {

            }
            finally
            {

            }
        }

        public DataTable SERACH_TB_SALES_PROMOTIONS()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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
                                    SELECT  
                                    [ID]
                                    ,[ISCLOSEED]
                                    ,[SALESTO]
                                    ,[SDATES]
                                    ,[PRODUCTS]
                                    ,[SHIPDATES]
                                    ,[KINDS]
                                    ,[CONTEXTS]
                                    FROM [TKBUSINESS].[dbo].[TB_SALES_PROMOTIONS]
                                    WHERE [ISCLOSEED]IN ('N')
                                    ORDER BY [SDATES]
                                                                       
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"];
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

        public DataTable SERACH_MAIL_TB_SALES_PROMOTIONS()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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
                                    SELECT 
                                    [ID]
                                    ,[SENDTO]
                                    ,[MAIL]
                                    ,[NAME]
                                    ,[COMMENTS]
                                    FROM [TKMQ].[dbo].[MQSENDMAIL]
                                    WHERE [SENDTO]='TB_SALES_PROMOTIONS'
                                                                       
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"];
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

        public void SENDEMAIL_PURNOTIN()
        {
            DataTable DS_EMAIL_TO_EMAIL = new DataTable();
            DataTable DT_PURNOTIN = new DataTable();

            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();

            try
            {
                DS_EMAIL_TO_EMAIL = SERACH_MAIL_PURNOTIN();
                DT_PURNOTIN = SERACH_PURNOTIN();

                if (DT_PURNOTIN != null && DT_PURNOTIN.Rows.Count >= 1)
                {
                    SUBJEST.Clear();
                    BODY.Clear();


                    SUBJEST.AppendFormat(@"系統通知-請查收-每日-預計採購未到貨明細，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                    //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                    //ERP 採購相關單別、單號未核準的明細
                    //
                    BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                        + "<br>" + "系統通知-請查收-每日-預計採購未到貨明細，謝謝"
                        + " <br>"
                        );





                    if (DT_PURNOTIN.Rows.Count > 0)
                    {
                        BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                        BODY.AppendFormat(@"<table> ");
                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">預交日</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">廠商</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購單別</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購單號</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">序號</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品號</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品名</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">規格</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購數量</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">單位</th>");

                        BODY.AppendFormat(@"</tr> ");

                        foreach (DataRow DR in DT_PURNOTIN.Rows)
                        {

                            BODY.AppendFormat(@"<tr >");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["預交日"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["廠商"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["採購單別"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["採購單號"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["序號"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["品號"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["品名"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["規格"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["採購數量"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["單位"].ToString() + "</td>");
                            BODY.AppendFormat(@"</tr> ");


                        }
                        BODY.AppendFormat(@"</table> ");
                    }
                    else
                    {
                        BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "本日無資料");
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
                            foreach (DataRow DR in DS_EMAIL_TO_EMAIL.Rows)
                            {
                                MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail
                            }

                            //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                            MySMTP.Send(MyMail);

                            MyMail.Dispose(); //釋放資源

                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show("有錯誤");

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



            }
            catch
            {

            }
            finally
            {

            }
        }

        public DataTable SERACH_PURNOTIN()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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
                                    SELECT 
                                    TD012 AS '預交日'
                                    ,TC004 AS '供應廠商'
                                    ,MA002 AS '廠商'
                                    ,TD001 AS '採購單別'
                                    ,TD002 AS '採購單號'
                                    ,TD003 AS '序號'
                                    ,TD004 AS '品號'
                                    ,TD005 AS '品名'
                                    ,TD006 AS '規格'
                                    ,TD008 AS '採購數量'
                                    ,TD009 AS '單位'
                                    FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.PURMA
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND MA001=TC004
                                    AND TC014='Y'
                                    AND REPLACE(TD001+TD002+TD003,' ','') NOT IN (SELECT REPLACE(TH011+TH012+TH013,' ','') FROM [TK].dbo.PURTH)
                                    AND TD008>0
                                    AND TD012=CONVERT(NVARCHAR,GETDATE(),112)
                                    ORDER BY TC004

                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"];
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

        public DataTable SERACH_MAIL_PURNOTIN()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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
                                    SELECT 
                                    [ID]
                                    ,[SENDTO]
                                    ,[MAIL]
                                    ,[NAME]
                                    ,[COMMENTS]
                                    FROM [TKMQ].[dbo].[MQSENDMAIL]
                                    WHERE [SENDTO]='PURNOTIN'
                                                                       
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"];
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

        public void SENDEMAIL_TBPURCHECKFAX()
        {
            DataTable DS_EMAIL_TO_EMAIL = new DataTable();
            DataTable DT_DATAS = new DataTable();

            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();

            try
            {
                DS_EMAIL_TO_EMAIL = SERACH_MAIL_TBPURCHECKFAX();
                DT_DATAS = SERACH_TBPURCHECKFAX();

                if (DT_DATAS != null && DT_DATAS.Rows.Count >= 1)
                {
                    SUBJEST.Clear();
                    BODY.Clear();


                    SUBJEST.AppendFormat(@"系統通知-請查收-每日-預計採購未傳真明細，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                    //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                    //ERP 採購相關單別、單號未核準的明細
                    //
                    BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                        + "<br>" + "系統通知-請查收-每日-預計採購未傳真明細，謝謝"
                        + " <br>"
                        );





                    if (DT_DATAS.Rows.Count > 0)
                    {
                        BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                        BODY.AppendFormat(@"<table> ");
                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購單別</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購單號</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">廠商</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">預交日</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">序號</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品號</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品名</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">規格</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購數量</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">單位</th>");

                        BODY.AppendFormat(@"</tr> ");

                        foreach (DataRow DR in DT_DATAS.Rows)
                        {

                            BODY.AppendFormat(@"<tr >");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["採購單別"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["採購單號"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["廠商"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["預交日"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["序號"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["品號"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["品名"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["規格"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["採購數量"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["單位"].ToString() + "</td>");
                            BODY.AppendFormat(@"</tr> ");


                        }
                        BODY.AppendFormat(@"</table> ");
                    }
                    else
                    {
                        BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "本日無資料");
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
                            foreach (DataRow DR in DS_EMAIL_TO_EMAIL.Rows)
                            {
                                MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail
                            }

                            //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                            MySMTP.Send(MyMail);

                            MyMail.Dispose(); //釋放資源

                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show("有錯誤");

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



            }
            catch
            {

            }
            finally
            {

            }
        }

        public DataTable SERACH_TBPURCHECKFAX()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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
                                   SELECT 
                                    TD001 AS '採購單別'
                                    ,TD002 AS '採購單號'
                                    ,TC004 AS '供應廠商'
                                    ,MA002 AS '廠商'
                                    ,TD012 AS '預交日'
                                    ,TD003 AS '序號'
                                    ,TD004 AS '品號'
                                    ,TD005 AS '品名'
                                    ,TD006 AS '規格'
                                    ,TD008 AS '採購數量'
                                    ,TD009 AS '單位'
                                    FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.PURMA
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND MA001=TC004
                                    AND TC002 LIKE '%'+CONVERT(NVARCHAR,GETDATE(),112)+'%'
                                    AND REPLACE([TC001]+[TC002],' ','' ) NOT IN (SELECT REPLACE([TC001]+[TC002],' ','' ) FROM [TKPUR].[dbo].[TBPURCHECKFAX] )
                                    ORDER BY TD001,TD002,TD003
                                    
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"];
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

        public DataTable SERACH_MAIL_TBPURCHECKFAX()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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
                                    SELECT 
                                    [ID]
                                    ,[SENDTO]
                                    ,[MAIL]
                                    ,[NAME]
                                    ,[COMMENTS]
                                    FROM [TKMQ].[dbo].[MQSENDMAIL]
                                    WHERE [SENDTO]='TBPURCHECKFAX'
                                                                       
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"];
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

        public void SENDEMAIL_TB_DEVE_NEWLISTS()
        {
            DataTable DS_EMAIL_TO_EMAIL = new DataTable();
            DataTable DT_DATAS = new DataTable();

            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();

            try
            {
                DS_EMAIL_TO_EMAIL = SERACH_MAIL_TB_DEVE_NEWLISTS();
                DT_DATAS = SERACH_TB_DEVE_NEWLISTS();

                if (DT_DATAS != null && DT_DATAS.Rows.Count >= 1)
                {
                    SUBJEST.Clear();
                    BODY.Clear();


                    SUBJEST.AppendFormat(@"系統通知-請查收-每週-本月研發已開發的樣品，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                    //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                    //ERP 採購相關單別、單號未核準的明細
                    //
                    BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                        + "<br>" + "系統通知-請查收-每週-本月研發已開發的樣品，謝謝"
                        + " <br>"
                        );





                    if (DT_DATAS.Rows.Count > 0)
                    {
                        BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                        BODY.AppendFormat(@"<table> ");
                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">編號</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">商品</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">規格</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">業務</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">註記</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">差異特色</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">打樣日期</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">表單編號</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">業務回覆</th>");

                        BODY.AppendFormat(@"</tr> ");

                        foreach (DataRow DR in DT_DATAS.Rows)
                        {

                            BODY.AppendFormat(@"<tr >");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["編號"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["商品"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["規格"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["業務"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["註記"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["差異特色"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["打樣日期"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["DOC_NBR"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["F9FieldValue"].ToString() + "</td>");
                            BODY.AppendFormat(@"</tr> ");


                        }
                        BODY.AppendFormat(@"</table> ");
                    }
                    else
                    {
                        BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "本日無資料");
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
                            foreach (DataRow DR in DS_EMAIL_TO_EMAIL.Rows)
                            {
                                MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail
                            }

                            //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                            MySMTP.Send(MyMail);

                            MyMail.Dispose(); //釋放資源

                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show("有錯誤");

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



            }
            catch
            {

            }
            finally
            {

            }
        }

        public DataTable SERACH_TB_DEVE_NEWLISTS()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            string yyyyMM = DateTime.Now.ToString("yyyyMM");
            string YY = yyyyMM.Substring(2, 2);
            string MM = yyyyMM.Substring(4, 2);
            string NOLIKE = YY + '-' + MM;

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
                                   SELECT 
                                    [NO] AS '編號'
                                    ,[NAMES] AS '商品'
                                    ,[SPECS] AS '規格'
                                    ,[SALES] AS '業務'
                                    ,[COMMENTS] AS '註記'
                                    ,[INGREDIENTS] AS '差異特色'
                                    ,CONVERT(NVARCHAR,[GETDATES],112)  AS '打樣日期'
                                    ,[REPLY] AS '業務回覆'
                                    ,[SALESID] AS '業務ID'
                                    ,[COSTS] AS '成本'
                                    ,[MOQS] AS 'MOQ'
                                    ,[MANUPRODS] AS '一天產能量'
                                    ,CONVERT(NVARCHAR,[CARESTEDATES],112) AS '建立日期'
                                    ,[ID]
                                    ,(SELECT TOP 1 [DOC_NBR] FROM [192.168.1.223].[UOF].[dbo].[View_TKRS_TB_DEVE_NEWLISTS] WHERE [View_TKRS_TB_DEVE_NEWLISTS].[F01FieldValue]=[TB_DEVE_NEWLISTS].NO  COLLATE Chinese_Taiwan_Stroke_BIN ORDER BY [DOC_NBR] DESC) AS 'DOC_NBR'
                                    ,(SELECT TOP 1 [F01FieldValue] FROM [192.168.1.223].[UOF].[dbo].[View_TKRS_TB_DEVE_NEWLISTS] WHERE [View_TKRS_TB_DEVE_NEWLISTS].[F01FieldValue]=[TB_DEVE_NEWLISTS].NO  COLLATE Chinese_Taiwan_Stroke_BIN ORDER BY [DOC_NBR] DESC) AS 'F01FieldValue'
                                    ,(SELECT TOP 1 [F09FieldValue] FROM [192.168.1.223].[UOF].[dbo].[View_TKRS_TB_DEVE_NEWLISTS] WHERE [View_TKRS_TB_DEVE_NEWLISTS].[F01FieldValue]=[TB_DEVE_NEWLISTS].NO  COLLATE Chinese_Taiwan_Stroke_BIN ORDER BY [DOC_NBR] DESC) AS 'F9FieldValue'


                                    FROM [TKRESEARCH].[dbo].[TB_DEVE_NEWLISTS]

 
                                    WHERE 1=1
                                    AND [NO] LIKE '{0}%'
                                    ORDER BY [NO]
                                    
                                    ", NOLIKE);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"];
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

        public DataTable SERACH_MAIL_TB_DEVE_NEWLISTS()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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
                                    SELECT 
                                    [ID]
                                    ,[SENDTO]
                                    ,[MAIL]
                                    ,[NAME]
                                    ,[COMMENTS]
                                    FROM [TKMQ].[dbo].[MQSENDMAIL]
                                    WHERE [SENDTO]='TB_DEVE_NEWLISTS'
                                                                       
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
                adapter.Fill(ds, "ds");
                sqlConn.Close();



                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"];
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

        public void SENDEMAIL_DAILY_MOCMANULINE()
        {
            DataSet dsSALESMONEYS = new DataSet();
            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();

            SETPATH();

            DATES = DateTime.Now.ToString("yyyyMMdd");
            DirectoryNAME = @"C:\MQTEMP\" + DATES.ToString() + @"\";
            pathFileMOCMANULINE = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日預排製令表" + DATES.ToString() + ".pdf";
            //如果日期資料夾不存在就新增
            if (!Directory.Exists(DirectoryNAME))
            {
                //新增資料夾
                Directory.CreateDirectory(DirectoryNAME);
            }


            SAVEREPORT_DAILY_MOCMANULINE(pathFileMOCMANULINE);

            DataTable ds_MAIL_DAILY_MOCMANULINE = SERACH_MAIL_MOCMANULINE();

            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"系統通知-每日預排製令表" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日預排製令表，請查收" + Environment.NewLine + " ");

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
            //MyMail.IsBodyHtml = true; //是否使用html格式

            System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
            MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);

            Attachment attch = new Attachment(pathFileMOCMANULINE);
            MyMail.Attachments.Add(attch);


            try
            {
                foreach (DataRow od in ds_MAIL_DAILY_MOCMANULINE.Rows)
                {

                    MyMail.To.Add(od["MAIL"].ToString()); //設定收件者Email，多筆mail
                }

                //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email

                MySMTP.Send(MyMail);

                MyMail.Dispose(); //釋放資源


            }
            catch (Exception ex)
            {
                ADDLOG(DateTime.Now, SUBJEST.ToString(), ex.ToString());
                //ex.ToString();
            }
        }

        public void SAVEREPORT_DAILY_MOCMANULINE(string pathFile)
        {
            string FILENAME = pathFile;
            //string FILENAME = @"C:\MQTEMP\20210915\每日業務單位業績日報表20210915.pdf";
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL_DAILY_MOCMANULINE();
            Report report1 = new Report();

            report1.Load(@"REPORT\每日預排製令表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            //adapter1.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
            table.Connection.CommandTimeout = 300;
            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));


            // prepare a report
            report1.Prepare();
            // create an instance of HTML export filter
            FastReport.Export.Pdf.PDFExport export = new FastReport.Export.Pdf.PDFExport();
            // show the export options dialog and do the export
            report1.Export(export, FILENAME);

        }

        public StringBuilder SETSQL_DAILY_MOCMANULINE()
        {
            DateTime now = DateTime.Now;
            // 取得本月第一天日期
            DateTime firstDayOfMonth = new DateTime(now.Year, now.Month, 1);
            // 取得本月最後一天日期
            int daysInMonth = DateTime.DaysInMonth(now.Year, now.Month);
            DateTime lastDayOfMonth = new DateTime(now.Year, now.Month, daysInMonth);


            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"  
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



            return SB;

        }

        public void SENDEMAIL_DAILY_QC_TEMP_CHECK()
        {
            DataSet ds = new DataSet();
            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();

            SETPATH();

            DATES = DateTime.Now.ToString("yyyyMMdd");
            DirectoryNAME = @"C:\MQTEMP\" + DATES.ToString() + @"\";
            //string pathFile_QC_TEMP_CHECK = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日溫溼度明細" + DATES.ToString() + ".pdf";
            string pathFile_QC_TEMP_CHECK = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日溫溼度明細" + DATES.ToString() + ".jpg";

            //如果日期資料夾不存在就新增
            if (!Directory.Exists(DirectoryNAME))
            {
                //新增資料夾
                Directory.CreateDirectory(DirectoryNAME);
            }


            SAVEREPORT_QC_TEMP_CHECK(pathFile_QC_TEMP_CHECK);

            ds = SERACHMAIL_QC_CHECK();


            System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();

            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"每日-每日溫溼度明細-" + DateTime.Now.ToString("yyyy/MM/dd"));
            BODY.AppendFormat("Dear All, ");
            BODY.AppendFormat(Environment.NewLine + "檢附每日溫溼度明細，請參考附件，謝謝");
            BODY.AppendFormat("<br /><br />");
            BODY.AppendFormat("<img src=\"cid:image001\" alt=\"图片描述\" style=\"width:400px;\" />");
            BODY.AppendFormat("<br /><br />Thank you for your attention.");

            string emailBody = BODY.ToString();
            // 创建 HTML 视图
            AlternateView htmlView = AlternateView.CreateAlternateViewFromString(emailBody, null, MediaTypeNames.Text.Html);

            // 使用本地图片路径添加图片附件
            string imagePath = pathFile_QC_TEMP_CHECK;  // 本地图片路径
            LinkedResource imageResource = new LinkedResource(imagePath, MediaTypeNames.Image.Jpeg);
            imageResource.ContentId = "image001";  // 设置 Content-ID
            imageResource.TransferEncoding = TransferEncoding.Base64;
            htmlView.LinkedResources.Add(imageResource);

            // 将 HTML 视图添加到邮件
            MyMail.AlternateViews.Add(htmlView);

            string MySMTPCONFIG = ConfigurationManager.AppSettings["MySMTP"];
            string NAME = ConfigurationManager.AppSettings["NAME"];
            string PW = ConfigurationManager.AppSettings["PW"];

           
            MyMail.From = new System.Net.Mail.MailAddress("tk290@tkfood.com.tw");

            //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
            //MyMail.Subject = "每日訂單-製令追踨表"+DateTime.Now.ToString("yyyy/MM/dd");
            MyMail.Subject = SUBJEST.ToString();
            //MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
            MyMail.Body = BODY.ToString();
            //MyMail.IsBodyHtml = true; //是否使用html格式

            System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
            MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);

            Attachment attch = new Attachment(pathFile_QC_TEMP_CHECK);
            MyMail.Attachments.Add(attch);


            try
            {
                foreach (DataRow od in ds.Tables[0].Rows)
                {

                    MyMail.To.Add(od["MAIL"].ToString()); //設定收件者Email，多筆mail
                }

                //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email

                MySMTP.Send(MyMail);

                MyMail.Dispose(); //釋放資源


            }
            catch (Exception ex)
            {
                ADDLOG(DateTime.Now, SUBJEST.ToString(), ex.ToString());
                //ex.ToString();
            }
        }
        public void SAVEREPORT_QC_TEMP_CHECK(string pathFile)
        {
            string FILENAME = pathFile;
            //string FILENAME = @"C:\MQTEMP\20210915\每日業務單位業績日報表20210915.pdf";
            StringBuilder SQL1 = new StringBuilder();

            DateTime now = DateTime.Now;
            now = now.AddDays(-1);
            string SDAYS = now.ToString("yyyyMMdd");



            SQL1.AppendFormat(@"
                            --20240828 查溫溼度

                            SELECT 
                            [區域],
                            DATEPART(YEAR, [日期時間]) AS '年',
                            DATEPART(MONTH, [日期時間]) AS '月',
                            DATEPART(DAY, [日期時間]) AS '日',
                            DATEPART(HOUR, [日期時間]) AS '時',
                            AVG(CONVERT(decimal(16,4),[控項_1])) AS '溫度',
                            AVG(CONVERT(decimal(16,4),[控項_4])) AS '溼度',
                            (CONVERT(NVARCHAR,DATEPART(YEAR, [日期時間])) +CONVERT(NVARCHAR,DATEPART(MONTH, [日期時間]))+CONVERT(NVARCHAR,DATEPART(DAY, [日期時間])) +CONVERT(NVARCHAR,DATEPART(HOUR, [日期時間]) )) AS 'DATETIMES'

                            FROM [TK_FOOD].[dbo].[log_table]
                            LEFT JOIN [TK_FOOD].[dbo].[Machine] ON [Machine].[機台名稱] = [log_table].[機台名稱]
                            WHERE [Machine].[機台名稱] IN ( '溫濕度13', '溫濕度14')
                            AND CONVERT(NVARCHAR,[日期時間],112)='{0}'
                            GROUP BY 
                            [區域],
                            DATEPART(YEAR, [日期時間]), 
                            DATEPART(MONTH, [日期時間]), 
                            DATEPART(DAY, [日期時間]), 
                            DATEPART(HOUR, [日期時間])
                            ORDER BY 
                            [區域],
                            DATEPART(HOUR, [日期時間])

                            "
                            , SDAYS);
            Report report1 = new Report();

            report1.Load(@"REPORT\溫溼度明細.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["TKA01"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
            table.Connection.CommandTimeout = TIMEOUT_LIMITS;
            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));


            // prepare a report 
            report1.Prepare();
            // create an instance of HTML export filter
            //FastReport.Export.Pdf.PDFExport export = new FastReport.Export.Pdf.PDFExport();
            FastReport.Export.Image.ImageExport ImageExport = new FastReport.Export.Image.ImageExport();

            // show the export options dialog and do the export
            report1.Export(ImageExport, FILENAME);

        }

        //針對昨天核單的 總務採購單，給申請人發出公告
        public void NEW_GRAFFAIRS_1005_TB_EIP_BULLETIN()
        {
            DataTable DTSEARCHUOF_GRAFFAIRS_1005 = SEARCHUOF_GRAFFAIRS_1005_NEW();
            string xmlString = "";
            string xmlString_UserSet = "";

            //空的UserSet
            // 創建 <UserSet> 標籤
            XElement userSetElement = new XElement("UserSet");
            // 創建 XDocument 並添加 <UserSet>
            XDocument xmlDoc = new XDocument(userSetElement);

            // 使用 StringWriter 和 XmlTextWriter 將 XmlDocument 轉換為字串            
            using (StringWriter stringWriter = new StringWriter())
            {
                using (XmlTextWriter xmlTextWriter = new XmlTextWriter(stringWriter))
                {
                    xmlTextWriter.Formatting = Formatting.Indented; // 如果你想要縮進的格式化
                    xmlDoc.WriteTo(xmlTextWriter);
                    xmlTextWriter.Flush();

                    // 取得 XML 的字串表示
                    xmlString = stringWriter.GetStringBuilder().ToString();

                    // 輸出字串
                   // Console.WriteLine(xmlString);
                }
            }


            ////申請人的UserSet
            //// 創建 <UserSet> 標籤
            ////何翔鈞 192f1ddd-f6ef-4725-81e0-dc15c15a10cf
            //XElement userSetElement_UserSet = new XElement("UserSet",
            //    new XElement("Element",
            //        new XAttribute("type", "user"),
            //        new XElement("userId", "b6f50a95-17ec-47f2-b842-4ad12512b431")
            //    )
            //);
            //// 創建 XDocument 並添加 <UserSet>
            //XDocument xmlDoc_UserSet = new XDocument(userSetElement_UserSet);
            //using (StringWriter stringWriter = new StringWriter())
            //{
            //    using (XmlTextWriter xmlTextWriter = new XmlTextWriter(stringWriter))
            //    {
            //        xmlTextWriter.Formatting = Formatting.Indented; // 如果你想要縮進的格式化
            //        xmlDoc_UserSet.WriteTo(xmlTextWriter);
            //        xmlTextWriter.Flush();

            //        // 取得 XML 的字串表示
            //        xmlString_UserSet = stringWriter.GetStringBuilder().ToString();

            //        // 輸出字串
            //        // Console.WriteLine(xmlString);
            //    }
            //}

            string BULLETIN_GUID = Guid.NewGuid().ToString();
            string ANNOUNCER = "192f1ddd-f6ef-4725-81e0-dc15c15a10cf";
            string CLASS_GUID = "2e6d7f89-abcb-426b-afd6-8191fff9a668"; //01.行政類公告
            string TOPIC = "測試公告";
            string CONTEXT = "測試公告";
            string EXPIRE_DATE = DateTime.Now.AddDays(7).ToString("yyyyMMdd");
            string RM_ID = Guid.NewGuid().ToString();
            string FILE_GROUP_ID = "";
            string CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODIFY_USER = "";
            string MODIFY_DATE = DateTime.Now.ToString("yyyyMMdd");
            string PRINT = "N";
            string PRINT_USER_SET = xmlString;
            string MARQUEE = "N";
            string SYNCMSG = "Y";
            string STATUS = "Publish";
            string ATTACHMENT = "N";
            string CREATE_USER = "b6f50a95-17ec-47f2-b842-4ad12512b431";
            string PUBLISH_DATE = DateTime.Now.ToString("yyyyMMdd");
            string IS_CLONE = "0";
            string OUTER_BULLETION_ID = "";
            string OUTER_CLASS_GUID = "";
            string ANNOUNCER_DEP = "1577d8f3-a244-4af7-8ba5-b29c5f92e7d0";
            string OUTER_BULLETION_READER = xmlString;
            string OUTER_BULLETION_ALLOW_PRINT = "0";
            string OUTER_BULLETION_PRINT_USER = xmlString;
            string AUTO_PUBLISH_CONTROL = "0";
            string AUTO_PUBLISH_DATE = "";
            string IS_DELETE_OUTER_BULL = "0";
            string IS_DISPLAY_READER = "1";
            string IS_DISPLAY_OUTER_READER = "1";
            string UPDATE_TASK_ID = "";
            string IS_READER_IN_INNER = "0";
            string IS_STICKY = "";
            string RECOMMEND_NUM = "";

           
            string ROLE_ID = "BulletinBrowser";
            string USER_SET = xmlString_UserSet;

            if (DTSEARCHUOF_GRAFFAIRS_1005 != null && DTSEARCHUOF_GRAFFAIRS_1005.Rows.Count >= 1)
            {
                BULLETIN_GUID = new Guid().ToString();

                foreach (DataRow DR in DTSEARCHUOF_GRAFFAIRS_1005.Rows)
                {
                    //申請人的UserSet
                    // 創建 <UserSet> 標籤                   
                    XElement userSetElement_UserSet = new XElement("UserSet",
                        new XElement("Element",
                            new XAttribute("type", "user"),
                            new XElement("userId", DR["USER_GUID"].ToString())
                            //new XElement("userId", "b6f50a95-17ec-47f2-b842-4ad12512b431")
                        )
                    );
                    // 創建 XDocument 並添加 <UserSet>
                    XDocument xmlDoc_UserSet = new XDocument(userSetElement_UserSet);
                    using (StringWriter stringWriter = new StringWriter())
                    {
                        using (XmlTextWriter xmlTextWriter = new XmlTextWriter(stringWriter))
                        {
                            xmlTextWriter.Formatting = Formatting.Indented; // 如果你想要縮進的格式化
                            xmlDoc_UserSet.WriteTo(xmlTextWriter);
                            xmlTextWriter.Flush();

                            // 取得 XML 的字串表示
                            xmlString_UserSet = stringWriter.GetStringBuilder().ToString();

                            // 輸出字串
                            // Console.WriteLine(xmlString);
                        }
                    }

                    //每筆總務採購單，都發1張公告
                    BULLETIN_GUID = Guid.NewGuid().ToString();
                    RM_ID = Guid.NewGuid().ToString();
                    TOPIC = "請購物品: "+DR["請購物品"].ToString()+" ，已採購";
                    CONTEXT = "請購物品: " + DR["請購物品"].ToString() + " ，已採購" + ", 總務請購單:" + DR["總務請購單"].ToString() + ", 總務採購單:" + DR["總務採購單"].ToString();
                    USER_SET = xmlString_UserSet;

                    //新增公告
                    ADD_UOF_TB_EIP_BULLETIN(
                     BULLETIN_GUID,
                     ANNOUNCER,
                     CLASS_GUID,
                     TOPIC,
                     CONTEXT,
                     EXPIRE_DATE,
                     RM_ID,
                     FILE_GROUP_ID,
                     CREATE_DATE,
                     MODIFY_USER,
                     MODIFY_DATE,
                     PRINT,
                     PRINT_USER_SET,
                     MARQUEE,
                     SYNCMSG,
                     STATUS,
                     ATTACHMENT,
                     CREATE_USER,
                     PUBLISH_DATE,
                     IS_CLONE,
                     OUTER_BULLETION_ID,
                     OUTER_CLASS_GUID,
                     ANNOUNCER_DEP,
                     OUTER_BULLETION_READER,
                     OUTER_BULLETION_ALLOW_PRINT,
                     OUTER_BULLETION_PRINT_USER,
                     AUTO_PUBLISH_CONTROL,
                     AUTO_PUBLISH_DATE,
                     IS_DELETE_OUTER_BULL,
                     IS_DISPLAY_READER,
                     IS_DISPLAY_OUTER_READER,
                     UPDATE_TASK_ID,
                     IS_READER_IN_INNER,
                     IS_STICKY,
                     RECOMMEND_NUM
                    );

                    //新增公告對象=申請人
                    ADD_UOF_TB_EB_SEC_ROLE_MEMBER(
                        RM_ID,
                        ROLE_ID,
                        USER_SET
                        );
                }
            }

            ////新增公告
            //ADD_UOF_TB_EIP_BULLETIN(
            // BULLETIN_GUID,
            // ANNOUNCER,
            // CLASS_GUID,
            // TOPIC,
            // CONTEXT,
            // EXPIRE_DATE,
            // RM_ID,
            // FILE_GROUP_ID,
            // CREATE_DATE,
            // MODIFY_USER,
            // MODIFY_DATE,
            // PRINT,
            // PRINT_USER_SET,
            // MARQUEE,
            // SYNCMSG,
            // STATUS,
            // ATTACHMENT,
            // CREATE_USER,
            // PUBLISH_DATE,
            // IS_CLONE,
            // OUTER_BULLETION_ID,
            // OUTER_CLASS_GUID,
            // ANNOUNCER_DEP,
            // OUTER_BULLETION_READER,
            // OUTER_BULLETION_ALLOW_PRINT,
            // OUTER_BULLETION_PRINT_USER,
            // AUTO_PUBLISH_CONTROL,
            // AUTO_PUBLISH_DATE,
            // IS_DELETE_OUTER_BULL,
            // IS_DISPLAY_READER,
            // IS_DISPLAY_OUTER_READER,
            // UPDATE_TASK_ID,
            // IS_READER_IN_INNER,
            // IS_STICKY,
            // RECOMMEND_NUM
            //);

            ////新增公告對象=申請人
            //ADD_UOF_TB_EB_SEC_ROLE_MEMBER(
            //    RM_ID,
            //    ROLE_ID,
            //    USER_SET
            //    );

        }

        public DataTable SEARCHUOF_GRAFFAIRS_1005_NEW()
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
                                
                                    SELECT 
                                    TEMPALL.KINDS
                                    ,TEMPALL.DOC_NBR AS '總務採購單'
                                    ,TEMPALL.GG002 AS '請購物品'
                                    ,TEMPALL.EXTERNAL_FORM_NBR AS '總務請購單'
                                    ,TB_WKF_TASK1.END_TIME AS '總務採購單結案時間'
                                    ,TB_EB_USER.NAME AS '請購人'
                                    ,TB_EB_USER.USER_GUID 
                                    ,TB_WKF_TASK1.TASK_RESULT
                                    ,TB_WKF_TASK1.TASK_STATUS
                                    FROM (
                                    SELECT 
	                                    '合併' AS 'KINDS',
	                                    DOC_NBR,
                                        C.value('(Cell[@fieldId=""GG002""]/@fieldValue)[1]', 'VARCHAR(100)') AS GG002,
                                        C.value('(Cell[@fieldId=""EXTERNAL_FORM_NBR""]/@fieldValue)[1]', 'VARCHAR(100)') AS EXTERNAL_FORM_NBR
                                    FROM
                                        [UOF].[dbo].TB_WKF_TASK
                                    CROSS APPLY
                                        CURRENT_DOC.nodes('//Row') AS T(C)
                                    WHERE 1 = 1
                                    AND TB_WKF_TASK.CURRENT_DOC.exist('//Row') = 1

                                    UNION ALL
                                    SELECT

                                        '單筆' AS 'KINDS',
                                        TB_WKF_TASK.DOC_NBR,
                                        MAX(CASE WHEN x1.field_id.value('@fieldId', 'VARCHAR(50)') = 'GA005'
                                                 THEN x1.field_id.value('@fieldValue', 'VARCHAR(200)') END) AS GA005_value,
                                        MAX(CASE WHEN x1.field_id.value('@fieldId', 'VARCHAR(50)') = 'GA002'
                                                 THEN x1.field_id.value('@fieldValue', 'VARCHAR(200)') END) AS GA002_value
                                    FROM
                                        [UOF].[dbo].TB_WKF_TASK
                                    CROSS APPLY
                                        TB_WKF_TASK.CURRENT_DOC.nodes('/Form/FormFieldValue/FieldItem') AS x1(field_id)
                                    WHERE 1 = 1
                                    AND TB_WKF_TASK.CURRENT_DOC.exist('//Row') = 0
                                    GROUP BY TB_WKF_TASK.DOC_NBR
                                    ) AS TEMPALL
                                    LEFT JOIN[UOF].[dbo].TB_WKF_TASK TB_WKF_TASK1 ON TB_WKF_TASK1.DOC_NBR=TEMPALL.DOC_NBR
                                    LEFT JOIN[UOF].[dbo].TB_WKF_TASK TB_WKF_TASK2 ON TB_WKF_TASK2.DOC_NBR= TEMPALL.EXTERNAL_FORM_NBR
                                    LEFT JOIN [UOF].[dbo].TB_EB_USER ON TB_EB_USER.USER_GUID= TB_WKF_TASK2.USER_GUID
                                    WHERE 1=1
                                    AND TB_WKF_TASK1.TASK_RESULT= '0' AND TB_WKF_TASK1.TASK_STATUS= '2'
                                    AND TEMPALL.DOC_NBR LIKE 'GA1005%'
                                    AND ISNULL(TEMPALL.EXTERNAL_FORM_NBR ,'')<>''
                                    AND CONVERT(NVARCHAR,TB_WKF_TASK1.END_TIME,112)='{0}'
                                    ORDER BY TB_EB_USER.NAME

                                 

                                   ", END_TIME);

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                DS_FIND_UOF_TASK_APPLICATION_FORM.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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

        public void ADD_UOF_TB_EIP_BULLETIN(
            string BULLETIN_GUID,
            string ANNOUNCER,
            string CLASS_GUID,
            string TOPIC,
            string CONTEXT,
            string EXPIRE_DATE,
            string RM_ID,
            string FILE_GROUP_ID,
            string CREATE_DATE,
            string MODIFY_USER,
            string MODIFY_DATE,
            string PRINT,
            string PRINT_USER_SET,
            string MARQUEE,
            string SYNCMSG,
            string STATUS,
            string ATTACHMENT,
            string CREATE_USER,
            string PUBLISH_DATE,
            string IS_CLONE,
            string OUTER_BULLETION_ID,
            string OUTER_CLASS_GUID,
            string ANNOUNCER_DEP,
            string OUTER_BULLETION_READER,
            string OUTER_BULLETION_ALLOW_PRINT,
            string OUTER_BULLETION_PRINT_USER,
            string AUTO_PUBLISH_CONTROL,
            string AUTO_PUBLISH_DATE,
            string IS_DELETE_OUTER_BULL,
            string IS_DISPLAY_READER,
            string IS_DISPLAY_OUTER_READER,
            string UPDATE_TASK_ID,
            string IS_READER_IN_INNER,
            string IS_STICKY,
            string RECOMMEND_NUM
            )
        {
            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();


                using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                {
                    string query = @"
                                    INSERT INTO [UOF].dbo.[TB_EIP_BULLETIN]
                                    (
                                        [BULLETIN_GUID],
                                        [ANNOUNCER],
                                        [CLASS_GUID],
                                        [TOPIC],
                                        [CONTEXT],
                                        [EXPIRE_DATE],
                                        [RM_ID],
                                        [FILE_GROUP_ID],
                                        [CREATE_DATE],
                                        [MODIFY_USER],
                                        [MODIFY_DATE],
                                        [PRINT],
                                        [PRINT_USER_SET],
                                        [MARQUEE],
                                        [SYNCMSG],
                                        [STATUS],
                                        [ATTACHMENT],
                                        [CREATE_USER],
                                        [PUBLISH_DATE],
                                        [IS_CLONE],
                                        [OUTER_BULLETION_ID],
                                        [OUTER_CLASS_GUID],
                                        [ANNOUNCER_DEP],
                                        [OUTER_BULLETION_READER],
                                        [OUTER_BULLETION_ALLOW_PRINT],
                                        [OUTER_BULLETION_PRINT_USER],
                                        [AUTO_PUBLISH_CONTROL],
                                        [AUTO_PUBLISH_DATE],
                                        [IS_DELETE_OUTER_BULL],
                                        [IS_DISPLAY_READER],
                                        [IS_DISPLAY_OUTER_READER],
                                        [UPDATE_TASK_ID],
                                        [IS_READER_IN_INNER],
                                        [IS_STICKY],
                                        [RECOMMEND_NUM]
                                    )
                                    VALUES
                                    (
                                        @BULLETIN_GUID,
                                        @ANNOUNCER,
                                        @CLASS_GUID,
                                        @TOPIC,
                                        @CONTEXT,
                                        @EXPIRE_DATE,
                                        @RM_ID,
                                        @FILE_GROUP_ID,
                                        @CREATE_DATE,
                                        @MODIFY_USER,
                                        @MODIFY_DATE,
                                        @PRINT,
                                        @PRINT_USER_SET,
                                        @MARQUEE,
                                        @SYNCMSG,
                                        @STATUS,
                                        @ATTACHMENT,
                                        @CREATE_USER,
                                        @PUBLISH_DATE,
                                        @IS_CLONE,
                                        @OUTER_BULLETION_ID,
                                        @OUTER_CLASS_GUID,
                                        @ANNOUNCER_DEP,
                                        @OUTER_BULLETION_READER,
                                        @OUTER_BULLETION_ALLOW_PRINT,
                                        @OUTER_BULLETION_PRINT_USER,
                                        @AUTO_PUBLISH_CONTROL,
                                        @AUTO_PUBLISH_DATE,
                                        @IS_DELETE_OUTER_BULL,
                                        @IS_DISPLAY_READER,
                                        @IS_DISPLAY_OUTER_READER,
                                        @UPDATE_TASK_ID,
                                        @IS_READER_IN_INNER,
                                        @IS_STICKY,
                                        @RECOMMEND_NUM
                                    )
                                    ";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        // 使用參數化的方式傳遞值
                        cmd.Parameters.AddWithValue("@BULLETIN_GUID", BULLETIN_GUID);
                        cmd.Parameters.AddWithValue("@ANNOUNCER", ANNOUNCER);
                        cmd.Parameters.AddWithValue("@CLASS_GUID", CLASS_GUID);
                        cmd.Parameters.AddWithValue("@TOPIC", TOPIC);
                        cmd.Parameters.AddWithValue("@CONTEXT", CONTEXT);
                        cmd.Parameters.AddWithValue("@EXPIRE_DATE", EXPIRE_DATE);
                        cmd.Parameters.AddWithValue("@RM_ID", RM_ID);
                        cmd.Parameters.AddWithValue("@FILE_GROUP_ID", FILE_GROUP_ID);
                        cmd.Parameters.AddWithValue("@CREATE_DATE", CREATE_DATE);
                        cmd.Parameters.AddWithValue("@MODIFY_USER", MODIFY_USER);
                        cmd.Parameters.AddWithValue("@MODIFY_DATE", MODIFY_DATE);
                        cmd.Parameters.AddWithValue("@PRINT", PRINT);
                        cmd.Parameters.AddWithValue("@PRINT_USER_SET", PRINT_USER_SET);
                        cmd.Parameters.AddWithValue("@MARQUEE", MARQUEE);
                        cmd.Parameters.AddWithValue("@SYNCMSG", SYNCMSG);
                        cmd.Parameters.AddWithValue("@STATUS", STATUS);
                        cmd.Parameters.AddWithValue("@ATTACHMENT", ATTACHMENT);
                        cmd.Parameters.AddWithValue("@CREATE_USER", CREATE_USER);
                        cmd.Parameters.AddWithValue("@PUBLISH_DATE", PUBLISH_DATE);
                        cmd.Parameters.AddWithValue("@IS_CLONE", IS_CLONE);
                        cmd.Parameters.AddWithValue("@OUTER_BULLETION_ID", OUTER_BULLETION_ID);
                        cmd.Parameters.AddWithValue("@OUTER_CLASS_GUID", OUTER_CLASS_GUID);
                        cmd.Parameters.AddWithValue("@ANNOUNCER_DEP", ANNOUNCER_DEP);
                        cmd.Parameters.AddWithValue("@OUTER_BULLETION_READER", OUTER_BULLETION_READER);
                        cmd.Parameters.AddWithValue("@OUTER_BULLETION_ALLOW_PRINT", OUTER_BULLETION_ALLOW_PRINT);
                        cmd.Parameters.AddWithValue("@OUTER_BULLETION_PRINT_USER", OUTER_BULLETION_PRINT_USER);
                        cmd.Parameters.AddWithValue("@AUTO_PUBLISH_CONTROL", AUTO_PUBLISH_CONTROL);
                        cmd.Parameters.AddWithValue("@AUTO_PUBLISH_DATE", AUTO_PUBLISH_DATE);
                        cmd.Parameters.AddWithValue("@IS_DELETE_OUTER_BULL", IS_DELETE_OUTER_BULL);
                        cmd.Parameters.AddWithValue("@IS_DISPLAY_READER", IS_DISPLAY_READER);
                        cmd.Parameters.AddWithValue("@IS_DISPLAY_OUTER_READER", IS_DISPLAY_OUTER_READER);
                        cmd.Parameters.AddWithValue("@UPDATE_TASK_ID", UPDATE_TASK_ID);
                        cmd.Parameters.AddWithValue("@IS_READER_IN_INNER", IS_READER_IN_INNER);
                        cmd.Parameters.AddWithValue("@IS_STICKY", IS_STICKY);
                        cmd.Parameters.AddWithValue("@RECOMMEND_NUM", RECOMMEND_NUM);

                        // 開啟連接並執行命令
                        conn.Open();
                        cmd.ExecuteNonQuery();
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

        public void ADD_UOF_TB_EB_SEC_ROLE_MEMBER(
            string RM_ID,
            string ROLE_ID,
            string USER_SET            
            )
        {
            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();


                using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                {
                    string query = @"
                                    INSERT INTO  [UOF].[dbo].[TB_EB_SEC_ROLE_MEMBER]
                                    (
                                        [RM_ID]
                                        ,[ROLE_ID]
                                        ,[USER_SET]
                                    )
                                    VALUES
                                    (
                                        @RM_ID,
                                        @ROLE_ID,
                                        @USER_SET                                       
                                    )
                                    ";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        // 使用參數化的方式傳遞值
                        cmd.Parameters.AddWithValue("@RM_ID", RM_ID);
                        cmd.Parameters.AddWithValue("@ROLE_ID", ROLE_ID);
                        cmd.Parameters.AddWithValue("@USER_SET", USER_SET);
                       
                        // 開啟連接並執行命令
                        conn.Open();
                        cmd.ExecuteNonQuery();
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
            SENDEMAIL_DAILY_MOCMANULINE();
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
        private void button22_Click(object sender, EventArgs e)
        {
            //通知副總，總務未簽核的表單
            PREPARE_UOF_TASK_TASK_GRAFFIR();


        }
        private void button23_Click(object sender, EventArgs e)
        {
            SEND_LINE("Hello, world! " + DateTime.Now.ToString("yyyyMMddHHmmss"));
        }
        private void button24_Click(object sender, EventArgs e)
        {
            SEND_TEST_MAIL();
            //SEND_TEST_MAIL_2();
           

        }
        private void button25_Click(object sender, EventArgs e)
        {
            CHECK_TB_EIP_SCH_DEVOLVE();
        }
        private void button26_Click(object sender, EventArgs e)
        {
            CHECK_TB_EIP_SCH_DEVOLVE_MANAGER();
        }
        private void button27_Click(object sender, EventArgs e)
        {
            SETPATH();
            SETFILE_NEWSLAES(path_File_NEWSLAES);
            CLEAREXCEL();

            PREPARESENDEMAIL_NEWSLAES(path_File_NEWSLAES);

            MessageBox.Show("OK");
        }

        private void button28_Click(object sender, EventArgs e)
        {
            SETPATH();
            SETFILE_POSINV(path_File_POSINV);
            CLEAREXCEL();

            PREPARESENDEMAIL_POSINV(path_File_POSINV);
            MessageBox.Show("OK");
        }
        private void button29_Click(object sender, EventArgs e)
        {
            //path_File_COPTCD
            //每日訂單明細表
            SETPATH();
            SETFILE_COPTCD(path_File_COPTCD);
            CLEAREXCEL();

            PREPARESENDEMAIL_COPTCD(path_File_COPTCD);
            PREPARESENDEMAIL_COPTCD(path_File_COPTCD);
            PREPARESENDEMAIL_COPTCD(path_File_COPTCD);
            MessageBox.Show("OK");
        }
        private void button30_Click(object sender, EventArgs e)
        {
            SENDEMAIL_DAILY_SALES_MONEY();

            MessageBox.Show("完成");
        }

        private void button31_Click(object sender, EventArgs e)
        {
            //SETFASTREPORT();
            SENDEMAIL_DAILY_SALES_MONEY();
        }

        private void button32_Click(object sender, EventArgs e)
        {
            //SETFASTREPORT_QC_CHECK();
            SENDEMAIL_DAILY_QC_CHECK();
        }
        private void button33_Click(object sender, EventArgs e)
        {
            SENDEMAIL_DAILY_QC_CHECK();

            MessageBox.Show("完成");
        }
        private void button34_Click(object sender, EventArgs e)
        {
            UDPATE_PURVERSIONSNUMS_TOTALNUMS();
            NEW_PURVERSIONSNUMS();

            MessageBox.Show("完成");
        }

        private void button35_Click(object sender, EventArgs e)
        {
            SENDEMAIL_DAILY_TKWH_CALENDAR();

            MessageBox.Show("完成");
        }
        private void button36_Click(object sender, EventArgs e)
        {
            SENDEMAIL_TB_SALES_PROMOTIONS();
            MessageBox.Show("完成");
        }
        private void button37_Click(object sender, EventArgs e)
        {
            SENDEMAIL_PURNOTIN();
           
            MessageBox.Show("完成");
        }
        private void button38_Click(object sender, EventArgs e)
        {
            SENDEMAIL_TBPURCHECKFAX();
            MessageBox.Show("完成");
        }
        private void button39_Click(object sender, EventArgs e)
        {
            SENDEMAIL_TB_DEVE_NEWLISTS();
            MessageBox.Show("完成");
        }
        private void button40_Click(object sender, EventArgs e)
        {
            SENDEMAIL_DAILY_QC_TEMP_CHECK();

            MessageBox.Show("完成");
        }
        private void button41_Click(object sender, EventArgs e)
        {
            //針對昨天核單的 總務採購單，給申請人發出公告
            NEW_GRAFFAIRS_1005_TB_EIP_BULLETIN();

            MessageBox.Show("完成");
        }
        #endregion


    }
}
