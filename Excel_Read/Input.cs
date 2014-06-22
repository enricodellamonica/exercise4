using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Excel_Read {
    class Input {

        public Input() {

            Console.WriteLine("Please enter 1 for Read Excel");
            Console.WriteLine("Please enter 2 for writing Excel");
            var userinput = Console.ReadLine();
            switch(userinput) {
                case "1":
                     new CreateExcel();
                    break;
                case "2":
                    new ReadExcel();
                    break;

                default:
                    break;
                }
            }
        }
    }
