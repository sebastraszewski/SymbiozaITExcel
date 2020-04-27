using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

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
                Range lastCell = ExcelFile.Instance.Sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                return lastCell.Row; 
            }
            set { lastRow = value; }
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
                Lp = lastRow - 1;
                if (Lp == 1)
                {
                    ExcelFile.Instance.Sheet.Cells[1, 1].Value = "Lp.";
                    ExcelFile.Instance.Sheet.Cells[1, 2].Value = "Data";
                    ExcelFile.Instance.Sheet.Cells[1, 3].Value = "Numer zlecenia";
                    ExcelFile.Instance.Sheet.Cells[1, 4].Value = "Opis";
                    ExcelFile.Instance.Sheet.Cells[1, 5].Value = "Kwota";
                }
                ExcelFile.Instance.Sheet.Cells[lastRow, 1].Value = Lp;
                ExcelFile.Instance.Sheet.Cells[lastRow, 2].Value = Input.Date;
                ExcelFile.Instance.Sheet.Cells[lastRow, 3].Value = Input.Order;
                ExcelFile.Instance.Sheet.Cells[lastRow, 4].Value = Input.Description;
                ExcelFile.Instance.Sheet.Cells[lastRow, 5].Value = Input.Amount;
                ExcelFile.Instance.Application.DisplayAlerts = false;
                ExcelFile.Instance.Book.SaveAs(ExcelFile.Instance.PathFileName);
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}


