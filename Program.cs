using System;

namespace SymbiozaITExcel
{
    class Program
    {
       
        static void Main(string[] args)
        {
         
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
             
            }
        }



  
    }
}


