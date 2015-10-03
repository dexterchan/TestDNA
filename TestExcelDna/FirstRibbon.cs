using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration;
using System.Threading;

namespace TestExcelDna
{
    [ComVisible(true)]
    public class MyRibbon : ExcelRibbon
    {
        TestWinForm.UserControl1.CtrlmRequest cr = new TestWinForm.UserControl1.CtrlmRequest();
        TestWinForm.UserControl1 ctrl = new TestWinForm.UserControl1();
        object[,] result = null;
        bool waitSet = false;
        ExcelReference refCell = null;

         [ExcelFunction(IsMacroType = true)] 
        public void OnButtonPressed(IRibbonControl control)
        {
            //MessageBox.Show("Hello from control " + control.Id);


            string str=FirstAddIn.MyGetHostname();

            var UIHandler = new Action<object>((o) =>
            {
                ctrl.ShowDialog();
            });
             ExcelAsyncUtil.QueueAsMacro(() =>
            {
                refCell = (ExcelReference)XlCall.Excel(XlCall.xlfActiveCell);
            });
            
            if (waitSet) //avoid double submission
            {

                ThreadPool.QueueUserWorkItem(new WaitCallback(UIHandler));
                return;
            }
            else
            {
                waitSet = true;
            }
            
            var wait = new ManualResetEvent(false);

            var handler = new EventHandler((o, e) =>
            {
                cr = (TestWinForm.UserControl1.CtrlmRequest)o;
                result = MakeArrayetest(cr.row, cr.col);
                waitSet = false;
                wait.Set();

            });

            ctrl.registerCallback(handler);


            ThreadPool.QueueUserWorkItem(new WaitCallback(UIHandler));

            //For simplicity, we implement the wait here
            wait.WaitOne();
            

            //ExcelReference cell = ExcelAsyncUtil.QueueAsMacro(() =>XlCall.Excel(XlCall.xlfActiveCell);
            
            //ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            //MessageBox.Show("Active cell:" + cell.RowFirst+","+cell.ColumnFirst);
            //var activeCell = new ExcelReference(1, 1);
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                //ExcelReference cell = (ExcelReference)XlCall.Excel(XlCall.xlfActiveCell);
                ExcelReference cell = refCell;
                int testRowSize = cr.row;
                int testColSize = cr.col;

                var activeCell = new ExcelReference(cell.RowFirst,testRowSize+cell.RowFirst-1, cell.ColumnLast ,cell.ColumnLast + testColSize-1);
                //object[,] o = new object[testRowSize, testColSize];

                //for (int i = 0; i < testRowSize; i++)
                //{
                //    o[i, 0] = i;
                //    o[i, 1] = "test" + i;
                //    o[i, 2] = DateTime.Now;
                //    o[i, 3] = "" + i + ",3";
                //    o[i, 4] = "" + i + ",4";
                //}


                activeCell.SetValue(result);
                XlCall.Excel(XlCall.xlcSelect, activeCell);
                
            });
        }

        public static void ShowHelloMessage()
        {
            MessageBox.Show("Hello from 'ShowHelloMessage'.");
        }

        public static object[,] MakeArrayetest(int rows, int columns)
        {
            object[,] result = new object[rows, columns];
            for (int i = 0; i < rows; i++)
            {
                result[i, 0] = i;
                for (int j = 1; j < columns-1; j++)
                {
                    result[i, j] = string.Format("({0},{1})", i, j);
                }
                result[i, columns - 1] = DateTime.Now;
            }

            return result;
        }
    }


    public static class MyFunctions
    {
        public static string TestFunction()
        {
            return "Testing...OK";
        }

    }
}
