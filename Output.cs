using System;
using System.IO;
using NPOI.SS.UserModel;

namespace SymbiozaITExcel
{
    class Output 
    {

        public int Lp { get; set; }
        private int lastRow;

        public int LastRow
        {
            get 
            {
                int lastCell = ExcelFile.Instance.Sheet.LastRowNum;
                return lastCell; 
            }
            set { lastRow = ExcelFile.Instance.Sheet.LastRowNum; }
        }

        
        public Input Input { get; set; }
        public Output(Input input)
        {
            this.Input = input;
        }
        public  void SaveExcelFile()
        {
            try
            {
                int lastRow = LastRow + 1;
                Lp = lastRow;
                if (Lp == 1)
                {
                    IRow firstRow = ExcelFile.Instance.Sheet.CreateRow(0);
                    firstRow.CreateCell(0).SetCellValue("Lp.");
                    firstRow.CreateCell(1).SetCellValue("Data");
                    firstRow.CreateCell(2).SetCellValue("Numer zlecenia");
                    firstRow.CreateCell(3).SetCellValue("Opis");
                    firstRow.CreateCell(4).SetCellValue("Kwota");

                }
                IRow row = ExcelFile.Instance.Sheet.CreateRow(lastRow);
                row.CreateCell(0).SetCellValue(Lp);
                row.CreateCell(1).SetCellValue(Input.Date);
                row.CreateCell(2).SetCellValue(Input.Order);
                row.CreateCell(3).SetCellValue(Input.Description);
                row.CreateCell(4).SetCellValue(Input.Amount);
         

                if (!File.Exists(ExcelFile.Instance.PathFileName))
                {
                    using (FileStream fs = new FileStream(ExcelFile.Instance.PathFileName, FileMode.Create, FileAccess.Write))
                    {
                        ExcelFile.Instance.Book.Write(fs);
                    }
                }
                else
                {
                    using (FileStream fs = new FileStream(ExcelFile.Instance.PathFileName, FileMode.Open, FileAccess.Write))
                    {
                        ExcelFile.Instance.Book.Write(fs);
                    }
                }
              
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}


