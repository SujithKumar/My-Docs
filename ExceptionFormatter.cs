using System;
using System.Data;
using System.Configuration;
using System.IO;
using System.Web;

namespace ExceptionHandler
{
	public class ExceptionHandler
	{
		public static void HandleException(string sException)
		{
			try
			{
                string sLogFileName = "[ E R R O R L O G ]ChannelGenie.txt";
                StreamWriter sw;
                long MaxFileSize = long.Parse("1048576");
                FileInfo FI;
                string BckDir = null;                
                string sLogDirectory = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "Log\\";
                sLogDirectory = sLogDirectory + @"\Log\";
                if (File.Exists(sLogDirectory + sLogFileName))
                {
                    FI = new FileInfo(sLogDirectory + sLogFileName);
                    if (FI.Length > MaxFileSize)
                    {
                        BckDir = sLogDirectory + "BCK " + DateTime.Now.Day.ToString() + " - "
                            + DateTime.Now.Month.ToString() + " - " + DateTime.Now.Year.ToString()
                            + " " + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString()
                            + DateTime.Now.Second.ToString() + DateTime.Now.Millisecond.ToString() + @"\";
                        if (!Directory.Exists(BckDir)) Directory.CreateDirectory(BckDir);
                        File.Move(sLogDirectory + sLogFileName, BckDir + sLogFileName);
                    }
                }
                if (!Directory.Exists(sLogDirectory)) Directory.CreateDirectory(sLogDirectory);

                if (!File.Exists(sLogDirectory + sLogFileName))
                {
                    sw = File.CreateText(sLogDirectory + sLogFileName);
                    sw.WriteLine("*".PadRight(100, '*'));
                    sw.WriteLine("File Created in {0} at {1}", Environment.MachineName, DateTime.Now.ToShortDateString());
                    sw.WriteLine("*".PadRight(100, '*'));
                    sw.Flush();
                    sw.Close();
                }
                sw = File.AppendText(sLogDirectory + sLogFileName);
                sw.WriteLine("Date :{0}", DateTime.Now.ToString() + " => " + sException);
                sw.WriteLine("*".PadRight(100, '*'));
                sw.Flush();
                sw.Close();
			}
			catch (Exception ex)
            {
                
            }
		}
	}
}
