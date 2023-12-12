using EmailSendNotification.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Text.RegularExpressions;
using System.Net.Mail;

namespace EmailSendNotification
{
   public partial  class Service1 : ServiceBase
    {
        private Timer timer = null;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            timer = new Timer();
            this.timer.Interval = 60000;
            this.timer.AutoReset = false;
            this.timer.Elapsed += new ElapsedEventHandler(this.timer_tick);
            //timer.Enabled = true;
            timer.Start();
            WriteToLog("Service Started");
        }

        protected override void OnStop()
        {
            WriteToLog("Service Stopped");
        }

        public static void WriteToLog(string text)
        {
            string sTemp = ConfigurationManager.AppSettings["Path"] + "_" + DateTime.Now.ToString("dd_MM") + ".txt";
            FileStream Fs = new FileStream(sTemp, FileMode.OpenOrCreate | FileMode.Append);
            StreamWriter st = new StreamWriter(Fs);
            string dttemp = DateTime.Now.ToString("[dd:MM:yyyy] [HH:mm:ss:ffff]");
            st.WriteLine(dttemp + "\t" + text);
            st.Close();
        }

        private void timer_tick(object sender, ElapsedEventArgs e)
        {
            MainFunction();
            timer.Start();
            WriteToLog("Job Done");
        }
        public void MainFunction()
        {
            WriteToLog("MainFunction Started");
            GetSendingEmail();
        }

        public  void GetSendingEmail()
        {
            WebServices _ser = new WebServices();
            EmailDetails lstList = _ser.GetSendEmail();
            IsEmailDetails isEmailDetails = new IsEmailDetails();
            List<IsEmail> Input = new List<IsEmail>();
            string tomail = string.Empty;
            string ccmail = string.Empty;
            string bccmail = string.Empty;
            string body = string.Empty;
            string subject = string.Empty;
            bool result = false;

           

            List<EmailSending> lst = lstList.EmailSending;

            try
            {
                WriteToLog("Getting Email details");
                if (lst != null && lst.Count > 0)
                {

                    for (int i = 0; i < lst.Count; i++)
                    {
                        tomail = lst[i].ManagerEmail + ", " + lst[i].UserEmail;

                     



                        body = lst[i].Body;
                        subject = lst[i].Subject;
                        result = SendEmail(body, subject, tomail, ccmail, bccmail);

                        tomail = string.Empty;
                        ccmail = string.Empty;
                        bccmail = string.Empty;
                        body = string.Empty;
                        subject = string.Empty;
                        if (result == true)
                        {
                            WriteToLog("Mail Sent");

                            Input.Add(new IsEmail { id = Convert.ToInt32(lst[i].SendEmailID) }); //templateid = Convert.ToInt32(lst[i].TemplateId) });
                            isEmailDetails.Input = Input;
                            result = _ser.PushSendEmail(isEmailDetails);
                            WriteToLog("Updated sent emails");
                        }
                        else
                        {
                            Input.Add(new IsEmail { id = Convert.ToInt32(lst[i].SendEmailID) }); //templateid = Convert.ToInt32(lst[i].TemplateId) });
                            WriteToLog("mail not send " + Convert.ToString(lst[i].SendEmailID));
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                WriteToLog(ex.Message);
                result = false;
            }

            WriteToLog("Get Sending Email " + result.ToString());
        }


        public  bool SendEmail(string body, string subject, string tomail, string ccmail, string bccmail) //, List<Attachments> att)
        {
            bool result = false;
            try
            {

                MailMessage mail = new MailMessage();
                //System.Net.Mail.Attachment attachment;

                if (!string.IsNullOrEmpty(tomail))
                {
                    string[] toid = tomail.Split(',');
                    for (int i = 0; i < toid.Length; i++)
                    {
                        //if (!toid[i].Contains("daimler"))
                        //{
                        mail.To.Add(toid[i]);
                        //}

                    }
                }

                if (!string.IsNullOrEmpty(ccmail))
                {
                    string[] ccid = ccmail.Split(',');
                    for (int i = 0; i < ccid.Length; i++)
                    {
                        //if (!ccid[i].Contains("daimler"))
                        //{
                        mail.CC.Add(ccid[i]);
                        //}
                    }
                }
                if (!string.IsNullOrEmpty(bccmail))
                {
                    string[] bccid = bccmail.Split(',');
                    for (int i = 0; i < bccid.Length; i++)
                    {
                        //if (!bccid[i].Contains("daimler"))
                        //{
                        mail.Bcc.Add(bccid[i]);
                        //}
                    }
                }

                mail.Subject = subject;
                mail.IsBodyHtml = true;
                mail.Body = body;
                mail.From = new MailAddress(ConfigurationManager.AppSettings["FromId"]);
                mail.Priority = MailPriority.Normal;
                SmtpClient SMTPServer = new SmtpClient(ConfigurationManager.AppSettings["Host"], Convert.ToInt32(ConfigurationManager.AppSettings["Port"]));

                //Added for Local
                //SMTPServer.UseDefaultCredentials = false;
                //SMTPServer.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["FromId"], ConfigurationManager.AppSettings["Password"]);
                //SMTPServer.EnableSsl = true;
                //End


                SMTPServer.Send(mail);
                result = true;
            }
            catch (Exception ex)
            {
                WriteToLog("Send Email " + ex.Message);
                result = false;
            }
            return result;
        }


    }


}
