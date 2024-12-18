using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.Text;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.Cms;

namespace Sendmes
{

    
    class Program
    {

        public static string[,] data = new string[,]
        {
            {"罗宇宸","23398026"},
            {"樊玥池","24201087"},
            {"郭晨影","24271038" },
            {"桑吉拉姆","23322061" },
            {"贾筱雨","24281099" },
            {"林婷婷","21211334" },
            {"乔梦豪","21291178" },
            {"王雨妍","24242043" },
            {"李书涵菀淼","23341140" },
            {"何梦君","22121718" },
            {"李好","23301068" },
            {"孔婧怡","23301099" },
            {"陈晓贤","23281068" },
            {"秦岳圣","23251328" },
            {"金芷瑶","22321076"},
            {"陈池","23261080" },
            {"热依莎","24251290" },
            {"楚严","23341106" },
        };
        //public static string[,] data = new string[,]
        //    {
        //       {"秦岳圣","23251328" },


        //    };






        static bool jdj_day(string args)
        {

            if (args.Length < 6)
                return false;

            // 解析文件名中的日期部分
            string datePart = args.Substring(0, 6);
            if (!DateTime.TryParseExact(datePart, "yyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime fileDate))
            {
                // 如解析失败，表示格式不正确
                return false;
            }

            // 计算时间差值
            var daysDifference = (DateTime.Now - fileDate).TotalDays;

            // 如果超过我们设定的天数阈值，则认为应被归档
            return daysDifference > 7;

        }

        static bool IsInternetAvailable()
        {
            try
            {
                using (var ping = new Ping())
                {
                    var reply = ping.Send("8.8.8.8", 3000); // 3秒的超时
                    return reply.Status == IPStatus.Success;
                }
            }
            catch
            {
                return false;
            }
        }


        static void Main(string[] args)
        {
            if (IsInternetAvailable())
            {
                ProcessDesktopFiles();
            }
            else
            {
                Console.WriteLine("No Internet connection available. Retrying in 30 seconds...");
                Thread.Sleep(30000);

                Console.WriteLine("Retrying now...");
                if (IsInternetAvailable())
                {
                    ProcessDesktopFiles();
                    Thread.Sleep(5000);

                }
                else
                {
                    Console.WriteLine("Still no Internet connection. Exiting.");
                }
            }
        }

        static void ProcessDesktopFiles()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var unclassifiedFiles = Directory.GetDirectories(desktopPath)
                .Where(folder => jdj_day(Path.GetFileName(folder)))
                .ToList();

            // 创建一个字典来存储每个人的未归档文件  
            var userFiles = new Dictionary<string,List<string>>();
            var userName = new Dictionary<string, string>();

            var adminFiles = new List<string>(); // 存储只有日期的文件  

            foreach (var folder in unclassifiedFiles)
            {
                string folderName = Path.GetFileName(folder);
                bool foundUser = false;

                // 检查文件名中是否包含任何用户的姓名  
                for (int i = 0; i < data.GetLength(0); i++)
                {
                    if (!string.IsNullOrEmpty(data[i, 0]) && folderName.Contains(data[i, 0]))
                    {

                        string user  = data[i, 0];
                        string email = $"{data[i, 1]}@bjtu.edu.cn";

                        if (!userFiles.ContainsKey(email))
                        {
                            userFiles[email] = new List<string>();
                        }

                        userName[email] = user; 

                        userFiles[email].Add(folderName);
                        foundUser = true;
                        break;
                    }
                }

                // 如果文件名中只包含日期，添加到管理员列表  
                if (!foundUser)
                {
                    adminFiles.Add(folderName);
                }
            }

            // 发送邮件给每个用户  
            foreach (var kvp in userFiles)
            {
                string thisUser = userName[kvp.Key];   

                Console.WriteLine("人员信息：" + thisUser);
                foreach (var item in kvp.Value) { Console.WriteLine(item); }

                SendEmailWithFolderContents(kvp.Key, kvp.Value, true, thisUser);
            }

            // 如果有只包含日期的文件，发送给管理员  
            if (adminFiles.Any())
            {
                Console.WriteLine("未正确命名文档，传送给管理员（admin）...");
                foreach (var item in adminFiles) { Console.WriteLine(item); }
                SendEmailWithFolderContents("23251328@bjtu.edu.cn", adminFiles, false, null);
            }
        }

        static void SendEmailWithFolderContents(string recipientEmail, List<string> folders, bool jdj, string name)
        {
            string smtpServer = "smtp.qq.com";
            int smtpPort = 465;
            string smtpUser = "3450462398@qq.com";
            string smtpPassword = "dgqkxcsjlxxrdcab";

            if (folders.Any())
            {
                // 获取exe所在目录  
                string exePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string templatePath = Path.Combine(exePath, "iMeg.html");
                string htmlTemplate = Htmlsupport.LoadTemplate(templatePath);

                string fileListHtml = string.Join("", folders.Select(folder => $"<li>{folder}</li>"));
                string htmlBody = Htmlsupport.FillTemplate(htmlTemplate, fileListHtml);

                var email = new MimeMessage();
                email.From.Add(MailboxAddress.Parse(smtpUser));
                email.To.Add(MailboxAddress.Parse(recipientEmail));
                email.Subject = "校党宣传部自动通知";

                string date = DateTime.Today.ToString("yy/MM/dd");

                if (jdj)
                {

                    var textPart = new TextPart(TextFormat.Html)
                    {
                        Text = name + "同学你好：<br><br>&nbsp;&nbsp;&nbsp;&nbsp;以下照片已超过7天没有归档,请记得及时归档。" +
                           htmlBody +
                           $@"<div style='text-align:right;'>  
                                from:315办公室计算机<br>{date}  
                            </div>",
                        ContentTransferEncoding = ContentEncoding.Base64
                    };
                    textPart.ContentType.Charset = "utf-8";
                    email.Body = textPart;
                }
                else
                {

                    var textPart = new TextPart(TextFormat.Html)
                    {
                        Text = "To admin：<br><br>&nbsp;&nbsp;&nbsp;&nbsp;以不符合命名规范照片已超过7天没有归档！" +
                           htmlBody +
                           $@"<div style='text-align:right;'>  
                                from:315办公室计算机<br>{date}  
                            </div>",
                        ContentTransferEncoding = ContentEncoding.Base64
                    };
                    textPart.ContentType.Charset = "utf-8";
                    email.Body = textPart;
                }


                using (var smtpClient = new SmtpClient())
                {
                    int retries = 3;
                    while (retries > 0)
                    {
                        try
                        {
                            smtpClient.Connect(smtpServer, smtpPort, SecureSocketOptions.SslOnConnect);
                            smtpClient.Authenticate(smtpUser, smtpPassword);
                            smtpClient.Send(email);
                            Console.WriteLine($"Email sent successfully to {recipientEmail}\n");
                            break;
                        }
                        catch (Exception ex)
                        {
                            retries--;
                            Console.WriteLine($"Failed to send email to {recipientEmail}: {ex.Message}. Retries left: {retries}");
                            if (retries == 0)
                            {
                                Console.WriteLine("All retries failed.");
                            }
                        }
                        finally
                        {
                            smtpClient.Disconnect(true);
                        }
                    }
                }
            }
        }
    }
}