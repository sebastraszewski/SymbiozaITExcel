using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Security.Cryptography;

namespace SymbiozaITExcel
{
    class Program
    {
       
        static void Main(string[] args)
        {
            handler = new ConsoleEventDelegate(ConsoleEventCallback);
            SetConsoleCtrlHandler(handler, true);
            try
            {
                RunProgram.fullfilNewRowInExcel();
                while (RunProgram.Continue == true)
                {
                    RunProgram.continueOrExit();
                }

            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                closeApp();
            }
        }



        public static void closeApp()
        {
            do
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            while (Marshal.AreComObjectsAvailableForCleanup());
            Process id = GetExcelProcess(ExcelFile.Instance.Application);
            id.Kill();
        }
        static bool ConsoleEventCallback(int eventType)
        {
            Process id = GetExcelProcess(ExcelFile.Instance.Application);
            id.Kill();


            return false;

        }
        static ConsoleEventDelegate handler;

        private delegate bool ConsoleEventDelegate(int eventType);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetConsoleCtrlHandler(ConsoleEventDelegate callback, bool add);
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        static Process GetExcelProcess(Microsoft.Office.Interop.Excel.Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }
    }
}


