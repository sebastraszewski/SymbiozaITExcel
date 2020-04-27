using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SymbiozaITExcel
{
    class RunProgram
    {
        public static bool Continue { get; set; }

        public static void continueOrExit()
        {

            Console.WriteLine("Wciśnij 'Enter', aby rozpocząć wpisywanie nowego wiersza do Excela.");
            if (Console.ReadKey().Key.ToString() == "Enter")
            {
                fullfilNewRowInExcel();
                Continue = true;
            }
            else
            {
                Console.WriteLine("Wciśnij 'Esc', jeśli chcesz wyjść.");
                if (Console.ReadKey().Key.ToString() == "Escape")
                {
                    Program.closeApp();
                    Environment.Exit(0);
                }
                else
                {

                    fullfilNewRowInExcel();
                    Continue = true;
                }
            }
        }

        public static void fullfilNewRowInExcel()
        {
            Input input = new Input();
            input.CollectInfoAboutRow();
            Output output = new Output(input);
            Console.WriteLine("Wciśnij 'Enter', aby zapisać dane do arkuszu excel - 'Skoroszyt.xslx'.");
            if (Console.ReadKey().Key.ToString() == "Enter")
            {
                output.SaveExcelFile();
                Console.WriteLine("Zapisano.");
                CreateNew();
            }
            else
            {
                CreateNew();
            }


        }

        public static void CreateNew()
        {
            Console.WriteLine("Wciśnij 'Enter', jeśli chcesz rozpocząć od nowa albo 'Escape', by wyjść z programu.");
            if (Console.ReadKey().Key.ToString() == "Enter")
            {
                fullfilNewRowInExcel();
            }
            else if (Console.ReadKey().Key.ToString() == "Escape")
            {
                Program.closeApp();
                Environment.Exit(0);
            }
            else
            {
                CreateNew();
            }
        }
    }
}

