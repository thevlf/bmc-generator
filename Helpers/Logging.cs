using System.IO;
using System.Reflection;
using System.Windows;

namespace BMCGen
{
    public class Logging
    {
        static string assembly = Assembly.GetEntryAssembly().FullName.Split(',')[0];
        //static string logFileName = $"{assembly}_{Guid.NewGuid()}.log";
        static string logFileName = $"{assembly}.log";

        public static void WriteLog(string message)
        {
            try
            {
                
                using (StreamWriter sw = File.AppendText(logFileName))
                {
                    if (string.IsNullOrEmpty(message))
                    {
                        sw.WriteLine(message);
                    }
                    else
                    {
                        sw.WriteLine(DateTime.Now + " " + message);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error trying to write to log:\r\n" + ex.Message);                
            }
        }

        public static void WriteException(string message, Exception ex)
        {
            try
            {                
                using (StreamWriter sw = File.AppendText(logFileName))
                {
                    sw.WriteLine(DateTime.Now + " " + message);
                    sw.WriteLine(ex.Message);
                    if (ex.InnerException != null)
                    {
                        sw.WriteLine(ex.InnerException);
                    }
                    if (!string.IsNullOrEmpty(ex.StackTrace))
                    {
                        sw.WriteLine(ex.StackTrace);
                    }

                    MessageBox.Show(message);                    
                }

            }
            catch (Exception logEx)
            {
                MessageBox.Show("Error trying to write to log:\r\n" + logEx.Message);                
            }
        }                

        public static void LogCleanup()
        {
            try
            {
                if (File.Exists(logFileName))
                {
                    File.Delete(logFileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error trying to delete log:\r\n" + ex.Message);
            }
        }
    }
}
