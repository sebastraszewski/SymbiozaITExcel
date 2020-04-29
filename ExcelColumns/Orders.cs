using System;
using System.Collections.Generic;
using System.Linq;

namespace SymbiozaITExcel
{

    class Orders : Input
    {
       
        public List<String> AllOrders { get; set; }
        
      
        public bool findOrderInExcelFile(int order)
        {
            AllOrders = new List<String>();
            int lastCell = ExcelFile.Instance.Sheet.LastRowNum;
            for (int row = 1; row < lastCell+1; row++)
            {
                AllOrders.Add(ExcelFile.Instance.Sheet.GetRow(row).GetCell(2).ToString());
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


