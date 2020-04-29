using System;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;


namespace SymbiozaITExcel
{
    public class ExcelFile
    {
        public ISheet Sheet { get; set; }
        public IWorkbook Book { get; set; }
        public FileStream FS { get; set; }
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
                    if (instance == null || instance.Book == null)
                    {
                        connectToFile();
                        
                        return instance;
                         
                    }
                    else
                    {
                        return instance;
                    }
                    
                }
            }
        }

        private static void connectToFile()
        {
            instance = new ExcelFile();
            instance.PathFileName = Path.Combine(Environment.CurrentDirectory, "Skoroszyt.xls");
            if (File.Exists(instance.PathFileName))
            {
                using (FileStream fs = new FileStream(instance.PathFileName, FileMode.Open, FileAccess.Read))
                {
                    instance.Book = new HSSFWorkbook(fs);
                    instance.Sheet = instance.Book.GetSheetAt(0);
                }
            }
            else
            {

                instance.Book = new NPOI.HSSF.UserModel.HSSFWorkbook();
                instance.Sheet = instance.Book.CreateSheet("Symbioza");
            }
            
            
        }
    }
}

