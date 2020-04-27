using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace SymbiozaITExcel
{
    public class ExcelFile
    {

        public Worksheet Sheet { get; set; }
        public  Workbook Book { get; set; }
        public  Microsoft.Office.Interop.Excel.Application Application { get; set; }
        public string PathFileName { get; set; }
        private static ExcelFile instance = null;
        private static readonly object padlock = new object();
        ExcelFile()
        {
        }
        public static ExcelFile Instance
        {
            get
            {
                lock (padlock)
                {
                    if (instance == null)
                    {
                        instance = new ExcelFile();
                        instance.PathFileName = Path.Combine(Environment.CurrentDirectory, "Skoroszyt.xlsx");
                        instance.Application = new Microsoft.Office.Interop.Excel.Application();
                        if (!System.IO.File.Exists(instance.PathFileName))
                        {
                            instance.Book = instance.Application.Workbooks.Add();
                        }
                        else
                        {
                            instance.Book = instance.Application.Workbooks.Open(instance.PathFileName);
                        }
                        instance.Sheet = (Worksheet)instance.Book.Worksheets.get_Item(1);

                    }
                    return instance;
                }
            }
        }
    }
}

