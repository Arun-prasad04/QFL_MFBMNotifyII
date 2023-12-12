using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Json;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace EmailSendNotification.Services
{
    partial class WebServices : WebClient
    {
        public WebServices()
        {
            InitializeComponent();
        }

        public int Timeout { get; set; }

        protected override WebRequest GetWebRequest(Uri address)
        {
            var request = base.GetWebRequest(address);
            request.Timeout = Timeout;
            return request;
        }

        public string GenerateToken()
        {
            string token = "FromEmail";
            return token;
        }

        public EmailDetails GetSendEmail()
        {
            string url = ConfigurationManager.AppSettings["API"] + "QFL.svc/GetSendEmail";
            EmailDetails lst = new EmailDetails();
            try
            {

                var client = new WebServices();
                client.Timeout = 900000;
                client.Encoding = Encoding.UTF8;
                client.Headers["Content-type"] = "application/json";
                client.Headers["Authorization"] = GenerateToken();
                client.Headers["cache-control"] = "no-cache";
                client.Proxy.Credentials = CredentialCache.DefaultCredentials;
                var jsonString = client.DownloadString(url);
                lst = JsonConvert.DeserializeObject<EmailDetails>(jsonString);
            }
            catch (Exception ex)
            {
                WriteToLog("Service Consume Error");
                WriteToLog(ex.Message);
            }
            return lst;
        }

        public void WriteToLog(string text)
        {
            string sTemp = ConfigurationManager.AppSettings["Path"] + "_" + DateTime.Now.ToString("dd_MM") + ".txt";
            FileStream Fs = new FileStream(sTemp, FileMode.OpenOrCreate | FileMode.Append);
            StreamWriter st = new StreamWriter(Fs);
            string dttemp = DateTime.Now.ToString("[dd:MM:yyyy] [HH:mm:ss:ffff]");
            st.WriteLine(dttemp + "\t" + text);
            st.Close();
        }

        public bool PushSendEmail(IsEmailDetails Input)
        {
            string url = ConfigurationManager.AppSettings["API"] + "QFL.svc/UpdateSendEmail";
            string jsonString = string.Empty;
            bool result = false;
            try
            {
                var client = new WebServices();
                client.Timeout = 900000;
                client.Encoding = Encoding.UTF8;
                client.Credentials = System.Net.CredentialCache.DefaultCredentials;
                client.Headers["Content-type"] = "application/json";
                client.Headers["Authorization"] = GenerateToken();
                client.Headers["cache-control"] = "no-cache";
                client.Proxy.Credentials = CredentialCache.DefaultCredentials;
                MemoryStream stream = new MemoryStream();
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(IsEmailDetails));
                serializer.WriteObject(stream, Input);
                byte[] data = client.UploadData(url, "POST", stream.ToArray());
                stream = new MemoryStream(data);
                jsonString = Encoding.UTF8.GetString(stream.ToArray());
                result = true;
                stream.Close();

            }
            catch (Exception ex)
            {
                WriteToLog("Service Consume Error");
                WriteToLog(ex.Message);
            }
            return result;
        }


    }

    public class EmailDetails
    {
        public List<EmailSending> EmailSending { get; set; }

    }
    public class EmailSending
    {
        public string SendEmailID { get; set; }
        public string ManagerEmail { get; set; }
        public string UserEmail { get; set; }
        public string ModifiedBy { get; set; }
        public string ModifiedOn { get; set; }

        public string Subject { get; set; }
        public string Body { get; set; }
        public string Vinnumber { get; set; }
        public string Model { get; set; }
        public string QGateName { get; set; }
        public string InspectionItem { get; set; }
        public string Defect { get; set; }
        public string Comments { get; set; }
        public string CompletedBy { get; set; }
        public string CompletedDate { get; set; }
        public int VinId { get; set; }
        public int QFLFeedBackWorkFlowId { get; set; }

    }

    public class IsEmailDetails
    {
        public List<IsEmail> Input { get; set; }

    }

    public class IsEmail
    {

        public int id { get; set; }
        public int templateid { get; set; }
    }
}
