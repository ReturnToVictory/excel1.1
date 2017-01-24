using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;



namespace test_excel_1._0
{
    class Program
    {
        public static string test1(string buf, string ful_adress)
        {
            string pattern = buf;
            string answer = "";
            Regex regex1 = new Regex(pattern);
            Match match1 = regex1.Match(ful_adress);
            answer = match1.Groups[1].Value;
            return answer;
        }
        static void Main(string[] args)
        {

            Excel.Application excelapp;
            Excel.Range excelcells;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            Excel.Workbook excelappworkbook;



            string file_name = "D:\\Users\\ysibirkin\\Desktop\\Нормализация\\Копия ADDRHOUSE - возможна автоматическая обработка_end.xlsm";

            excelapp = new Excel.Application();
            excelapp.Visible = true;
            excelappworkbook = excelapp.Workbooks.Open(@file_name);
            excelsheets = excelappworkbook.Worksheets;

            /* excelworksheet = (Excel.Worksheet)excelsheets.get_Item(2);
             excelcells = excelworksheet.get_Range("AG10", "AG10");
             excelcells.Value2 = "Hello man";
             */
            string cell = "AO";
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            for (int i = 5; i < 38; i++)
            {
                excelcells = excelworksheet.get_Range((cell + i), (cell + i));
                string ful_adress = excelcells.Value2;

                Console.WriteLine(ful_adress);


                string pattern = @"([д.]+\s\d+)";

                pattern = @"^(\d+)";  // find index 
                Console.WriteLine("Index============" + test1(pattern, ful_adress));

                pattern = @"(Россия|Украина|РФ|РОССИЯ|Российская Федерация)";  // find country
                Console.WriteLine("Country==========" + test1(pattern, ful_adress));

                pattern = @"\b(г[.]\s\w+\D\w+|г[.]\s\w+|село\s\w+|пос[.]\s\w+|город\s\w+|п[.]\w+|г[.]\w+|Москва)";//город\s\w+|город\w+|г.\s\w+|г.\w+|село\s\w+|пос.\s\w+|Москва)";  // find_City
                Console.WriteLine("City=============" + test1(pattern,ful_adress));

                pattern = @"(д.\d+\w|д.\s\d+\w|д.\s\d+|д.\d+|д.\d+\s\d+|вл.\s+\d+|вл.\d+|влад.\s+\d+|дом\s\d+\w|дом\s\d+)";  // find_house
                Console.WriteLine("House=============" + test1(pattern, ful_adress));

                pattern = @"(корп\.\s*\d*)";  // find_Housing
                Console.WriteLine("Housing===========" + test1(pattern,ful_adress));

                pattern = @"(\w+\W\w+\sРеспублика|\w+\sРеспублика|Республика\s\w+|\w+\W+обл|\w+\W+край)";//@"([д]\W+\d+)"; // республика , область, край
                Console.WriteLine("Respublica=========" + test1(pattern, ful_adress));
                
                pattern = @"(\w+\W+р-н|\w+\W+район)";// район
                Console.WriteLine("Район=========" + test1(pattern, ful_adress));

                pattern = @"(\w+\sвал|\w+\sпроезд|ул.\w+|ул.\s\w+|\w+\sплощадь|\w+\sш.|\w+\sшоссе|\w+\sнаб.|\w+\sнабережная|бульвар\s\w+|\w+\sпер.|\w+\sпереулок|проспект\s\w+)";
                Console.WriteLine("Street===========" + test1(pattern, ful_adress));

                pattern = @"(к\s\d+|к.\d+|к.\s\d+|корпус\s+\d+|корп.\s+\d+|помещение\s+\d+)";//|строение\s+\d+|стр.\s+\d+|стр.\d+)";
                Console.WriteLine("Корпус===========" + test1(pattern, ful_adress));

                pattern = @"(строение\s+\d+|стр.\s+\d+|стр.\d+)";
                Console.WriteLine("Строение===========" + test1(pattern, ful_adress));

                pattern = @"(комн.\s\d+|комн.\d+|комн\s\d+|комн\d+|кв.\s\d+|кв.\d+|кв\s\d+|кв\d+|ком.\s\d+|ком.\d+|ком\s\d+|ком\d+|оф.\s\d+|оф.\d+|оф\s\d+|оф\d+)";  // find_Appartament
                Console.WriteLine("Appartament=======" +test1(pattern,ful_adress));
            }
            Console.ReadKey();

            excelappworkbook.Saved = true;
            excelapp.Quit();
        }
    }
}





   
