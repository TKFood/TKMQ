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

namespace TKMQ
{
    public partial class FrmMQMAIL : Form
    {


        public FrmMQMAIL()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SENDMAIL()
        {
            string MySMTPCONFIG = ConfigurationManager.AppSettings["MySMTP"];
            string NAME = ConfigurationManager.AppSettings["NAME"];
            string PW = ConfigurationManager.AppSettings["PW"];

            System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
            MyMail.From = new System.Net.Mail.MailAddress("tk290@tkfood.com.tw");
            MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
            //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
            MyMail.Subject = "Email Test";
            MyMail.Body = "<h1>HIHI</h1>"; //設定信件內容
            MyMail.IsBodyHtml = true; //是否使用html格式
            System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
            MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);
            try
            {
                MySMTP.Send(MyMail);
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
            if (Directory.Exists(@"C:\MQTEMP\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"C:\MQTEMP\");
            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SENDMAIL();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETFILE();
        }
        #endregion


    }
}
