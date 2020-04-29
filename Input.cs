using System;

namespace SymbiozaITExcel
{
    public class Input
    {
        public String Date { get; set; }
        public int Order { get; set; }
        public String Description { get; set; }
        public String Amount { get; set; }

        public void CollectInfoAboutRow()
        {
            try
            {
                Console.WriteLine("Wprowadź datę: ");
                Date = Console.ReadLine();
                FullfillProductionNumber();
                Console.WriteLine("Wprowadź opis: ");
                Description = Console.ReadLine();
                Console.WriteLine("Wprowadź kwotę: ");
                Amount = Console.ReadLine();
                Console.WriteLine("Wprowadzone dane:");
                Console.WriteLine("Data: " + Date + ", Numer zlecenia: " + Order + ", Opis: "+ Description+ ", Kwota: "+Amount);
              
            }
            catch (Exception ex)
            {

                throw new Exception("Wystąpił błąd.", ex);
            }
        }

        private void FullfillProductionNumber()
        {
            bool tryAgain = true;
            Orders orders = new Orders();
            while (tryAgain)
            {
                try
                {
                    Console.WriteLine("Wprowadź numer zlecenia: ");
                    var filledOrder = Console.ReadLine();
                    Order = int.Parse(filledOrder);
                   

                    
                    if (orders.findOrderInExcelFile(Order))
                    {
                        Console.WriteLine("Numer zlecenia już występuje.");
                        
                        FullfillProductionNumber();
                    }
                    tryAgain = false;
                }

                catch (System.FormatException)
                {
                    tryAgain = true;
                    Console.WriteLine("Numer zlecenia musi być liczbą całkowitą. Spróbuj jeszcze raz.");

                    
                }
                catch (Exception ex)
                {
                    throw ex;
                }
             
             
            }
        }
    }
}

