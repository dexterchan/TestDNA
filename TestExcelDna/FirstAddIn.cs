using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using System.Windows.Forms;
using System.Threading;
using System.Configuration;

namespace TestExcelDna
{
    public class FirstAddIn : IExcelAddIn
    {
        public delegate void GetRowCol(ref int row, ref int col, object[,] obj);
        static bool RunOnce = false;
        static object lastRun = null;

        static string hostname = "";
        static int port = 0;

        public void AutoOpen()
        {
            // Do your initialization here...
            try
            {
                var appSettings = ConfigurationManager.AppSettings;

                if (appSettings["hostname"] != null)
                {
                    hostname = appSettings["hostname"];
                }

                if (appSettings["port"] != null)
                {
                    string strPort = appSettings["port"];
                    port = Int32.Parse(strPort);
                }

                
                ExcelIntegration.RegisterUnhandledExceptionHandler(
                    ex => "!!! EXCEPTION: " + ex.ToString());
                //if (appSettings.Count == 0)
                //{
                //    Console.WriteLine("AppSettings is empty.");
                //}
                //else
                //{
                //    foreach (var key in appSettings.AllKeys)
                //    {
                //        Console.WriteLine("Key: {0} Value: {1}", key, appSettings[key]);
                //    }
                //}
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Error reading app settings");
            }
        }

        public void AutoClose()
        {
        }

        [ExcelFunction(Description = "My first Excel-DNA function")]
        public static string MyGetHostname()
        {
            return hostname;
        }

        [ExcelFunction(Description = "My first Excel-DNA function")]
        public static int MyGetPort()
        {
            return port;
        }

        [ExcelFunction(Description = "My first Excel-DNA function")]
        public static string MyFirstFunction(string name)
        {
            return "Hello " + name;
        }
        
        [ExcelFunction(Description = "My testing Excel-DNA function")]
        public static object mySampleOutput()
        {
            if (RunOnce)
            {
                RunOnce = !RunOnce;
                return XlCall.Excel(XlCall.xlUDF, "Resize", lastRun); ;
            }
            else
            {
                lastRun = null;
                RunOnce = !RunOnce;
            }
            object[,] o= new object[3,1];
            int a=20;
            o[0,0] = a;
            o[1,0] = "ABCD";
            o[2,0] = DateTime.Now;

            object result = o;
            lastRun = result;


            return XlCall.Excel(XlCall.xlUDF, "Resize", lastRun); ;
        }

        //Async not working here with row resize
        //[ExcelFunction(Description = "My testing Excel-DNA function")]
        //public static object mySampleTblAsync()
        //{
        //    if (RunOnce)
        //    {
        //        RunOnce = !RunOnce;
        //        return XlCall.Excel(XlCall.xlUDF, "Resize", lastRun); ;
        //    }
        //    else
        //    {
        //        lastRun = null;
        //        RunOnce = !RunOnce;
        //    }

        //    TestWinForm.UserControl1.CtrlmRequest cr = new TestWinForm.UserControl1.CtrlmRequest();
        //    object result = null;
        //    bool waitSet = false;
        //    //object result = MakeArrayet(5, 5);
        //    TestWinForm.UserControl1 ctrl = new TestWinForm.UserControl1();

        //    var wait = new ManualResetEvent(false);

        //    var handler = new EventHandler((o, e) =>
        //    {
        //        cr = (TestWinForm.UserControl1.CtrlmRequest)o;
        //        result = MakeArrayet(cr.row, cr.col);
        //        waitSet = true;
        //        wait.Set();

        //    });

        //    ctrl.registerCallback(handler);
        //    ctrl.Show();
        //    //var tHandler = new Action<object>((o) =>
        //    //{
        //    //    ctrl.ShowDialog();
        //    //});
        //    //ThreadPool.QueueUserWorkItem(new WaitCallback(tHandler));

        //    return ExcelAsyncUtil.Run("TableCreate", null,
        //        delegate
        //        {
                    

        //            //MyAsyncMethod(data, handler); // so it started and will fire handler soon 
                    
        //                wait.WaitOne();
                    
        //            lastRun = result;
        //            // Call Resize via Excel - so if the Resize add-in is not part of this code, it should still work.
        //            return XlCall.Excel(XlCall.xlUDF, "Resize", lastRun);
        //        });
            
            
        //}

        [ExcelFunction(Description = "My testing Excel-DNA function")]
        public static object mySampleTbl()
        {
            if (RunOnce)
            {
                RunOnce = !RunOnce;
                return XlCall.Excel(XlCall.xlUDF, "Resize", lastRun); ;
            }
            else
            {
                lastRun = null;
                RunOnce = !RunOnce;
            }
            TestWinForm.UserControl1.CtrlmRequest cr = new TestWinForm.UserControl1.CtrlmRequest();
            object result = null;
            bool waitSet = false;
            //object result = MakeArrayet(5, 5);
            TestWinForm.UserControl1 ctrl = new TestWinForm.UserControl1();
            
            var wait = new ManualResetEvent(false);

            var handler = new EventHandler((o, e) =>
            {
                cr = (TestWinForm.UserControl1.CtrlmRequest)o;
                result = MakeArrayet(cr.row, cr.col);
                waitSet = true;
                wait.Set();

            });
            ctrl.registerCallback(handler);

            var tHandler = new Action<object> ( (o) => {
                ctrl.ShowDialog();
            });
            ThreadPool.QueueUserWorkItem(new WaitCallback(tHandler));

            //MyAsyncMethod(data, handler); // so it started and will fire handler soon 
            if (!waitSet)
            {
                wait.WaitOne();
            }
            lastRun = result;
            // Call Resize via Excel - so if the Resize add-in is not part of this code, it should still work.
            return XlCall.Excel(XlCall.xlUDF, "Resize", lastRun);
        }

        public static object MakeArrayet(int rows, int columns)
        {
            object[,] result = new string[rows, columns];
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    result[i, j] = string.Format("({0},{1})", i, j);
                }
            }

            return result;
        }

       
        

    }
}
