using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace SymbiozaITExcel
{

    class Orders : Input
    {
       
        public List<String> AllOrders { get; set; }
        
      
        public bool findOrderInExcelFile(int order)
        {
            AllOrders = new List<String>();
            Range lastCell = ExcelFile.Instance.Sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            for (int row = 2; row < lastCell.Row+1; row++)
            {
                AllOrders.Add(ExcelFile.Instance.Sheet.Cells[row, 3].Value.ToString());
            }
            if (AllOrders != null)
            {
                string result = AllOrders.FirstOrDefault(s => s == order.ToString());
                if (!String.IsNullOrEmpty(result))
                {
                    return true;
                }
                else
                {
                    return false;
                    
                }
            }
            else
            {
                return false;
                
            }
            

        }
        
    }
}


