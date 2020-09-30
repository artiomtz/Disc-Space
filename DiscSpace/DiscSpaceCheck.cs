using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using NLog;
using System.ServiceProcess;
using System.Text.RegularExpressions;
using System.Timers;
using System.Net;
using System.Net.Mail;

namespace DiscSpace
{
    public partial class DiscSpaceCheck : ServiceBase
    {
        Timer _timer = new Timer();
        int _timerIntervalHours = 1;
        string _configFileName = "DriveSpaceConfig.json";
        string _logFileName = "DriveSpaceConfig_log.txt";
        List<string> _toEmail = new List<string>() { "artiomtz@gmail.com" };
        string _fromEmail = "artiomtz@gmail.com";
        string _fromEmailPassword = "PasswordHere";
        string _defaultSpaceNotification = "10GB";
        int _defaultDelayBetweenTwoNotificationsInHours = 4;
        NLog.Logger _logger = NLog.LogManager.GetCurrentClassLogger();
        string _path = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);

        public DiscSpaceCheck()
        {
            InitializeComponent();
            var config = new NLog.Config.LoggingConfiguration();
            var logFile = new NLog.Targets.FileTarget("logfile") { FileName = _path + "\\" + _logFileName};
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, logFile);         
            NLog.LogManager.Configuration = config;
        }

        protected override void OnStart(string[] args)
        {
            _logger.Info("Starting drive space notification service...");
            _logger.Info("Monitoring every {0} hour(s).", _timerIntervalHours);
            _logger.Info("Running a space check on start:");
            CheckSpace(CheckConfig()); // run on start
            _timer.Elapsed += new ElapsedEventHandler(TimerElapsed);
            //_timer.Interval = 2000; // test
            _timer.Interval = _timerIntervalHours * 60 * 60 * 1000; // milliseconds
            _timer.Enabled = true;            
        }

        protected override void OnStop()
        {
            _logger.Info("Stopping drive space notification service..."); 
            NLog.LogManager.Shutdown();
        }

        public void TimerElapsed(object source, ElapsedEventArgs e)
        {
            _logger.Info(" --- Timer elapsed - Checking space --- "); 
            CheckSpace(CheckConfig());
        }

        public Config CheckConfig()
        {
            _logger.Info("Checking the JSON config file...");
            string filePath = _path + "\\" + _configFileName;
            Config jsonFile = new Config();

            if (!File.Exists(filePath)) // file doesn't exist
            {
                _logger.Info("Creating a new JSON config file.");
                jsonFile.NotifyTo = _toEmail;
                jsonFile.NotifyIfFreeSpaceLessThan = _defaultSpaceNotification;
                jsonFile.DelayBetweenTwoNotificationsInHours = _defaultDelayBetweenTwoNotificationsInHours;
                jsonFile.LastNotificationSentTime = DateTime.Today.AddDays(-1);
                File.WriteAllText(filePath, JsonConvert.SerializeObject(jsonFile));                
            }

            jsonFile = JsonConvert.DeserializeObject<Config>(File.ReadAllText(filePath));

            try // convert to bytes
            {
                if (jsonFile.NotifyIfFreeSpaceLessThan.ToUpper().Contains("TB"))
                {
                    jsonFile.NotifyIfFreeSpaceLessThan = Math.Round((Convert.ToDouble(Regex.Match(jsonFile.NotifyIfFreeSpaceLessThan, @"\d+").Value) * 1000), 2).ToString();
                }
                else if (jsonFile.NotifyIfFreeSpaceLessThan.ToUpper().Contains("GB"))
                {
                    jsonFile.NotifyIfFreeSpaceLessThan = Math.Round((Convert.ToDouble(Regex.Match(jsonFile.NotifyIfFreeSpaceLessThan, @"\d+").Value)), 2).ToString();
                }
                else if (jsonFile.NotifyIfFreeSpaceLessThan.ToUpper().Contains("MB"))
                {
                    jsonFile.NotifyIfFreeSpaceLessThan = Math.Round((Convert.ToDouble(Regex.Match(jsonFile.NotifyIfFreeSpaceLessThan, @"\d+").Value) / 1000), 2).ToString();
                }
                else if (jsonFile.NotifyIfFreeSpaceLessThan.ToUpper().Contains("KB"))
                {
                    jsonFile.NotifyIfFreeSpaceLessThan = Math.Round((Convert.ToDouble(Regex.Match(jsonFile.NotifyIfFreeSpaceLessThan, @"\d+").Value) / 1000000), 2).ToString();
                }
                else // in bytes
                {
                    jsonFile.NotifyIfFreeSpaceLessThan = Math.Round((Convert.ToDouble(Regex.Match(jsonFile.NotifyIfFreeSpaceLessThan, @"\d+").Value) / 1000000000), 2).ToString();
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Couldn't parse NotifyIfFreeSpaceLessThan from JSON. {0}", ex.Message);
            }

            try // verify hours is a proper int
            {
                if ((Convert.ToInt32(jsonFile.DelayBetweenTwoNotificationsInHours) < 0) || (Convert.ToInt32(jsonFile.DelayBetweenTwoNotificationsInHours) > 24))
                {
                    _logger.Warn("Setting default value ({0}) for DelayBetweenTwoNotificationsInHours.", _defaultDelayBetweenTwoNotificationsInHours);
                    jsonFile.DelayBetweenTwoNotificationsInHours = _defaultDelayBetweenTwoNotificationsInHours;
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Couldn't parse DelayBetweenTwoNotificationsInHours from JSON. {0}", ex.Message);
            }
            return jsonFile;
        }

        public void CheckSpace(Config configFile)
        {
            _logger.Info("Checking drive spaces..."); 
            List<DriveInfo> skippedDrives = new List<DriveInfo>();
            List<DriveInfo> fullDrives = new List<DriveInfo>();
            DriveInfo[] allDrives;

            try
            {
                allDrives = DriveInfo.GetDrives();
            }
            catch (Exception ex)
            {
                _logger.Error("Couldn't retrieve drives. {0}", ex.Message);
                return;
            }

            foreach (DriveInfo drive in allDrives)
            {
                if (drive.DriveType != DriveType.Fixed) // check only fixed drives
                {
                    _logger.Info("DRIVE {0} : skipping (not fixed).", drive.Name); 
                    continue;
                }
                else if (!drive.IsReady) // cannot access drive
                {
                    _logger.Warn("DRIVE {0} : Couldn't access.", drive.Name);
                    skippedDrives.Add(drive);
                }
                else // check drive space
                {
                    if ((Convert.ToDouble(drive.AvailableFreeSpace) / 1000000000) < Convert.ToDouble(configFile.NotifyIfFreeSpaceLessThan)) // comparison in GB
                    {
                        _logger.Warn("DRIVE {0} : RUNNING OUT OF FREE SPACE.", drive.Name); 
                        fullDrives.Add(drive);
                    }
                    else
                    {
                        _logger.Info("DRIVE {0} : enough free space.", drive.Name);
                    }
                }
            }

            if (fullDrives.Count > 0 || skippedDrives.Count > 0)
            {
                //TimeSpan notificationDelay = new TimeSpan(0, 0, configFile.DelayBetweenTwoNotificationsInHours); // test
                TimeSpan notificationDelay = new TimeSpan(configFile.DelayBetweenTwoNotificationsInHours, 0, -1);

                if ((DateTime.Now.Subtract(configFile.LastNotificationSentTime)) >= notificationDelay)
                {
                    SendNotification(fullDrives, skippedDrives, configFile);
                }
                else
                {
                    _logger.Info("Skipping a notification. {0} hours didn't pass since last notification.", configFile.DelayBetweenTwoNotificationsInHours);
                }
            }
            else
            {
                _logger.Info("No issues with drive spaces.");
            }
        }

        public void SendNotification(List<DriveInfo> fullDrives, List<DriveInfo> skippedDrives, Config configFile)
        {
            _logger.Info("Sending a notification.");
            configFile.LastNotificationSentTime = DateTime.Now;
            configFile.NotifyIfFreeSpaceLessThan += "GB";
            string filePath = _path + "\\" + _configFileName;
            File.WriteAllText(filePath, JsonConvert.SerializeObject(configFile));
            string sendHtml = BuildHtml(fullDrives, skippedDrives);
            try
            {
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                message.From = new MailAddress(_fromEmail);
                foreach (string email in configFile.NotifyTo)
                {
                    message.To.Add(new MailAddress(email));
                }
                message.Subject = "Drive space warning";
                message.IsBodyHtml = true;
                message.Body = sendHtml;
                smtp.Port = 587;
                smtp.Host = "smtp.gmail.com"; 
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential(_fromEmail, _fromEmailPassword);
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Send(message);
            }
            catch (Exception ex)
            {
                _logger.Error("Couldn't send a notification. {0}", ex.Message);
            }
        }

        public string BuildHtml(List<DriveInfo> fullDrives, List<DriveInfo> skippedDrives)
        {
            try
            {
                string messageBody = "<b>Current status:</b><br><br>";
                if (fullDrives.Count > 0)
                {
                    string htmlTableStart = "<table style=\"border-collapse:collapse; text-align:center;\" >";
                    string htmlTableEnd = "</table>";
                    string htmlHeaderRowStart = "<tr style=\"background-color:#6FA1D2; color:#ffffff;\">";
                    string htmlHeaderRowEnd = "</tr>";
                    string htmlTrStart = "<tr style=\"color:#555555;\">";
                    string htmlTrEnd = "</tr>";
                    string htmlTdStart = "<td style=\" border-color:#5c87b2; border-style:solid; border-width:thin; padding: 5px;\">";
                    string htmlTdEnd = "</td>";

                    messageBody += htmlTableStart;
                    messageBody += htmlHeaderRowStart;
                    messageBody += htmlTdStart + "Drive Name" + htmlTdEnd;
                    messageBody += htmlTdStart + "Free Space" + htmlTdEnd;
                    messageBody += htmlTdStart + "Total Size" + htmlTdEnd;
                    messageBody += htmlTdStart + "Space Used" + htmlTdEnd;
                    messageBody += htmlHeaderRowEnd;

                    foreach (DriveInfo drive in fullDrives)
                    {
                        messageBody += htmlTrStart;
                        messageBody += htmlTdStart + drive.Name + htmlTdEnd;
                        messageBody += htmlTdStart + Math.Round((Convert.ToDouble(drive.TotalFreeSpace) / 1000000000), 2).ToString()+ "GB" + htmlTdEnd;
                        messageBody += htmlTdStart + Math.Round((Convert.ToDouble(drive.TotalSize) / 1000000000), 2).ToString() + "GB" + htmlTdEnd;
                        messageBody += htmlTdStart + Math.Round((100-((Convert.ToDouble(drive.AvailableFreeSpace) * 100) / Convert.ToDouble(drive.TotalSize))), 2).ToString() + "%" + htmlTdEnd;
                        messageBody += htmlTrEnd;
                    }
                    messageBody += htmlTableEnd + "<br>";
                }

                foreach (DriveInfo drive in skippedDrives)
                    messageBody += "Couldn't access drive " + drive.Name + "<br>";

                return messageBody;
            }
            catch (Exception ex)
            {
                _logger.Error("Couldn't build HTML. {0}", ex.Message);
                return null;
            }
        }
    }
}