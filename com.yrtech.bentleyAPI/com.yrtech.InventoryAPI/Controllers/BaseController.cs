using com.yrtech.InventoryAPI.Common;
using com.yrtech.InventoryAPI.DTO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Http;
using CDO;
using System.Web.Configuration;

namespace com.yrtech.InventoryAPI.Controllers
{
    public class BaseController : ApiController
    {

        public static byte[] Base64ToBytes(string base64Img)
        {
            if (!string.IsNullOrEmpty(base64Img))
            {
                byte[] bytes = Convert.FromBase64String(base64Img);
                return bytes;
            }
            return null;
        }

        public static Stream BytesToStream(byte[] dataBytes)
        {
            if (dataBytes == null)
            {
                return null;
            }
            MemoryStream ms = new MemoryStream(dataBytes);
            return ms;
        }

        public static string UploadBase64Pic(string filePath, string base64Img)
        {
            if (!string.IsNullOrEmpty(base64Img) && base64Img.Contains("data:image"))
            {
                if (string.IsNullOrEmpty(filePath))
                {
                    filePath = @"Bentley/"+WebConfigurationManager.AppSettings["Year"]+ @"/MarketAction/" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".jpg";
                }
                base64Img = base64Img.Trim().Replace("%", "").Replace(",", "").Replace(" ", "+");
                base64Img = base64Img.Substring(base64Img.IndexOf("base64") + 6);
                if (base64Img.Length % 4 > 0)
                {
                    base64Img = base64Img.PadRight(base64Img.Length + 4 - base64Img.Length % 4, '=');
                }
                Stream stream = BytesToStream(Base64ToBytes(base64Img));
                OSSClientHelper.UploadOSSFile(filePath, stream, stream.Length);
                Thread.Sleep(1);
            }
            else
            {
                filePath = base64Img;
            }
            return filePath;
        }

        public void SendEmail(string emailTo, string emailCC, string subjects, string body, string attachmentStream, string attachementFileName)
        {
            try
            {
                CDO.Message oMsg = new CDO.Message();
                Configuration conf = new ConfigurationClass();
                conf.Fields[CdoConfiguration.cdoSendUsingMethod].Value = CdoSendUsing.cdoSendUsingPort;
                conf.Fields[CdoConfiguration.cdoSMTPAuthenticate].Value = CdoProtocolsAuthentication.cdoBasic;
                conf.Fields[CdoConfiguration.cdoSMTPUseSSL].Value = true;
                conf.Fields[CdoConfiguration.cdoSMTPServer].Value = "smtp.163.com";//必填，而且要真实可用   
                conf.Fields[CdoConfiguration.cdoSMTPServerPort].Value = "465";//465特有
                conf.Fields[CdoConfiguration.cdoSendEmailAddress].Value = "<" + "keyvisionApproval@163.com" + ">";
                conf.Fields[CdoConfiguration.cdoSendUserName].Value = "keyvisionApproval@163.com";//真实的邮件地址   
                conf.Fields[CdoConfiguration.cdoSendPassword].Value = "asdqwe123";   //为邮箱密码，必须真实   
                conf.Fields.Update();

                oMsg.Configuration = conf;
  
                // oMsg.TextBody = System.Text.Encoding.UTF8;
                //Message.BodyEncoding = System.Text.Encoding.UTF8;
                oMsg.BodyPart.Charset = "utf-8";
                oMsg.HTMLBody = body;
                oMsg.Subject = subjects;
                oMsg.From = "keyvisionApproval@163.com";
                oMsg.To = emailTo;
                oMsg.CC = emailCC;
                //ADD attachment.
                //TODO: Change the path to the file that you want to attach.
                //oMsg.AddAttachment("C:\Hello.txt", "", "");
                //oMsg.AddAttachment("C:\Test.doc", "", "");
                oMsg.Send();
            }
            catch (System.Net.Mail.SmtpException ex)
            {
                throw ex;
            }
            //System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
            //client.Host = "smtp.163.com";//使用163的SMTP服务器发送邮件
            //client.Port = 465;
            ////client.UseDefaultCredentials = true   ;
            //client.UseDefaultCredentials = true;
            //client.Timeout = 5000;
            //client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
            //client.Credentials = new System.Net.NetworkCredential("keyvisionApproval@163.com", "asdqwe123");
            ////这里假定你已经拥有了一个163邮箱的账户，用户名为abc，密码为*******
            //System.Net.Mail.MailMessage Message = new System.Net.Mail.MailMessage();
            //Message.From = new System.Net.Mail.MailAddress("keyvisionApproval@163.com");//这里需要注意，163似乎有规定发信人的邮箱地址必须是163的，而且发信人的邮箱用户名必须和上面SMTP服务器认证时的用户名相同
            ////因为上面用的用户名abc作SMTP服务器认证，所以这里发信人的邮箱地址也应该写为abc@163.com
            ////Message.To.Add("123456@gmail.com");//将邮件发送给Gmail
            //Message.To.Add(emailTo);//将邮件发送给QQ邮箱
            //if (!string.IsNullOrEmpty(emailCC))
            //{
            //    Message.CC.Add(emailCC);
            //}
            //Message.Subject = subjects;
            //Message.Body = body;
            //// Message.Attachments.Add(new System.Net.Mail.Attachment(@"C:\Workspace\com.yrtech.Bentley\com.yrtech.bentleyAPI\com.yrtech.InventoryAPI\Content\Excel\LeadsReport.xlsx", System.Net.Mime.MediaTypeNames.Application.Octet));
            //Message.SubjectEncoding = System.Text.Encoding.UTF8;
            //Message.BodyEncoding = System.Text.Encoding.UTF8;
            //Message.Priority = System.Net.Mail.MailPriority.High;
            //Message.IsBodyHtml = true;
            //CommonHelper.log("Email:"+emailTo + " " + subjects);
            //Thread.Sleep(500);
            //client.Send(Message);
        }
    }
}
