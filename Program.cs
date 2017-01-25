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
            Regex regex1 = new Regex(pattern , RegexOptions.IgnoreCase);
            Match match1 = regex1.Match(ful_adress);
            answer = match1.Groups[1].Value;
            return answer;
        }
        public static void read(string text ,string cell, int index , Excel.Workbook excelappworkbook)
        {
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            Excel.Range excelcells;

            cell = cell + index;
            excelsheets = excelappworkbook.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            excelcells = excelworksheet.get_Range(cell, cell);
            excelcells.Value2 = text;
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
            for (int i = 3; i < 6768; i++)
            {
                excelcells = excelworksheet.get_Range((cell + i), (cell + i));
                string ful_adress = excelcells.Value2;

                //Console.WriteLine(ful_adress);

                if (excelcells.Value2==null)
                {
                    continue;
                }
                else
                {
                    string pattern = "";

                    pattern = @"\b(\d+)";  // find index 
                    //read(test1(pattern, ful_adress), "AT", i, excelappworkbook);

                    pattern = @"(Россия|Украина|РФ|РОССИЯ|Российская Федерация)";  // find country
                    //read(test1(pattern, ful_adress), "AU", i, excelappworkbook);

                    pattern = @"(\w+\W\w+\sРеспублика|\w+\sРеспублика|Республика\s\w+|\w+\W+обл|\w+\W+край)";//@"([д]\W+\d+)"; // республика , область, край
                    //read(test1(pattern, ful_adress), "AV", i, excelappworkbook);

                    pattern = @"\b(г[.]\s\w+\D\w+|г[.]\s\w+|село\s\w+|пос[.]\s\w+|город\s\w+|п[.]\w+|г[.]\w+|Москва)";//город\s\w+|город\w+|г.\s\w+|г.\w+|село\s\w+|пос.\s\w+|Москва)";  // find_City
                    //read(test1(pattern, ful_adress), "AW", i, excelappworkbook);

                    pattern = @"(\w+\sпроезд|ул.\w+|ул.\s\w+|\w+\sплощадь|\w+\sшоссе|\w+\sш.|\w+\sнабережная|\w+\sнаб.|бульвар\s\w+|\w+\sпереулок|\w+\sпер.|проспект\s\w+|\w+\sвал)";
                    //read(test1(pattern, ful_adress), "AX", i, excelappworkbook);

                    pattern = @"(д.\d+\w|д.\s\d+\w|д.\s\d+|д.\d+|д.\d+\s\d+|вл.\s+\d+|вл.\d+|влад.\s+\d+|дом\s\d+\w|дом\s\d+)";  // find_house
                    //read(test1(pattern, ful_adress), "AY", i, excelappworkbook);

                    pattern = @"(\w+\W+р-н|\w+\W+район)";// район
                    //read(test1(pattern, ful_adress), "BC", i, excelappworkbook);

                    pattern = @"(к\s\d+|к[.]\d+|к[.]\s\d+|корпус\s+\d+|корп[.]\s*\d+|помещение\s+\d+)"; // find_Housing
                    //read(test1(pattern, ful_adress), "AZ", i, excelappworkbook);

                    pattern = @"(строение\s+\d+|стр[.]\s+\d+|стр[.]\d+)"; // STROENIE
                    //read(test1(pattern, ful_adress), "BA", i, excelappworkbook);

                    pattern = @"(квартира\s*\d*|комн[.]\s\d+|комн[.]\d+|комн\s\d+|комн\d+|помещение\s*\d*|пом[.]\s*\d+|кв[.]\s\d+|кв[.]\d+|кв\s\d+|кв\d+|ком[.]\s\d+|ком[.]\d+|ком\s\d+|ком\d+|оф[.]\s\d+|оф[.]\d+|оф\s\d+|оф\d+)";  // find_Appartament
                    read(test1(pattern, ful_adress), "BB", i, excelappworkbook);
                }
               
            }
            Console.WriteLine("End");
            Console.ReadKey();

            excelappworkbook.Saved = true;
            excelapp.Quit();
        }
    }
}





   
