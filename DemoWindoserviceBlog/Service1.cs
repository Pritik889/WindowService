using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.ServiceProcess;
using System.Text;

namespace DemoWindoserviceBlog
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// OnStart Method write code if we want to execute code when service is started.
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStart(string[] args)
        {
            int hour = DateTime.Now.Hour;
            int minute = DateTime.Now.Minute;
            if (hour == 15 && minute == 01)
            {
                Copy_Excelfile_And_Paste_at_anotherloaction_OnServiceStart();
            }

        }
        /// <summary>
        /// OnStop Method write code if we want to execute the particular code when service is Stop.
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStop()
        {
            Create_ServiceStoptextfile();
        }


        #region User defined Method

        /// <summary>
        /// The purpose of this method is to copy Excel sheet from one location to another location at specified time period.
        /// </summary>
        public static void Copy_Excelfile_And_Paste_at_anotherloaction_OnServiceStart()
        {
            try
            {
                string source = "D:\\DemoWebservice";
                string Destination = "D:\\DemoWebservice\\ServicestartExcelSheetCollection";
                string filename = string.Empty;
                if (!(Directory.Exists(Destination) && Directory.Exists(source)))
                    return;
                string[] Templateexcelfile = Directory.GetFiles(source);
                foreach (string file in Templateexcelfile)
                {
                    if (Templateexcelfile[0].Contains("Template"))
                    {
                        filename = System.IO.Path.GetFileName(file);
                        Destination = System.IO.Path.Combine(Destination, filename.Replace(".xlsx", DateTime.Now.ToString("yyyyMMdd")) + ".xlsx");
                        System.IO.File.Copy(file, Destination, true);
                    }
                }

            }
            catch (Exception ex)
            {
                Create_ErrorFile(ex);
            }

        }
        /// <summary>
        /// purpose of this method is to maintain error log in text file.
        /// </summary>
        /// <param name="exx"></param>
        public static void Create_ErrorFile(Exception exx)
        {
            StreamWriter SW;
            if (!File.Exists(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "txt_" + DateTime.Now.ToString("yyyyMMdd") + ".txt")))
            {
                SW = File.CreateText(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "txt_" + DateTime.Now.ToString("yyyyMMdd") + ".txt"));
                SW.Close();
            }
            using (SW = File.AppendText(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "txt_" + DateTime.Now.ToString("yyyyMMdd") + ".txt")))
            {
                string[] str = new string[] { exx.Message==null?"":exx.Message.ToString(), exx.StackTrace==null?"":exx.StackTrace.ToString(),
                    exx.InnerException==null?"":exx.InnerException.ToString()};
                for (int i = 0; i < str.Length; i++)
                {
                    SW.Write("\r\n\n");
                    if (str[i] == str[0])
                        SW.WriteLine("Exception Message:" + str[i]);
                    else if (str[i] == str[1])
                        SW.WriteLine("StackTrace:" + str[i]);
                    else if (str[i] == str[2])
                        SW.WriteLine("InnerException:" + str[i]);
                }
                SW.Close();
            }
        }
        /// <summary>
        /// The purpose of this method is maintain service stop information in text file.
        /// </summary>
        public static void Create_ServiceStoptextfile()
        {
            string Destination = "D:\\DemoWebservice\\ServiceStopInforation";
            StreamWriter SW;
            if (Directory.Exists(Destination))
            {
                Destination = System.IO.Path.Combine(Destination, "txtServiceStop_" + DateTime.Now.ToString("yyyyMMdd") + ".txt");
                if (!File.Exists(Destination))
                {
                    SW = File.CreateText(Destination);
                    SW.Close();
                }
            }
            using (SW = File.AppendText(Destination))
            {
                SW.Write("\r\n\n");
                SW.WriteLine("Service Stopped at: " + DateTime.Now.ToString("dd-MM-yyyy H:mm:ss"));
                SW.Close();
            }
        }
        #endregion
    }
}
