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


            List<int[]> array_list = new List<int[]>();
            List<string[]> array_list_answer = new List<string[]>();
            int[] array = new int[10];

            string cell = "AO";
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            for (int i = 3; i < 6768; i++)
            {
                excelcells = excelworksheet.get_Range((cell + i), (cell + i));
                string ful_adress = excelcells.Value2;
                string[] array_string = new string[10];
                Array.Clear(array, 0, array.Length);
                //Console.WriteLine(ful_adress);
                Console.ReadKey();
                if (excelcells.Value2==null)
                {
                    continue;
                }
                else
                {
                    string pattern = "";

                    pattern = @"\b(\d{5,7})";  // find index 
                    if (test1(pattern, ful_adress).Length != 0)
                    {
                        array[0] = 1;
                         array_string[0]= test1(pattern, ful_adress);
                        read(test1(pattern, ful_adress), "AT", i, excelappworkbook);
                    }

                    pattern = @"(Россия|Украина|РФ|РОССИЯ|Российская Федерация)";  // find country
                    if (test1(pattern, ful_adress).Length != 0)
                    {
                        array[1] = 1;
                        array_string[1] = test1(pattern, ful_adress);
                        read(test1(pattern, ful_adress), "AU", i, excelappworkbook);
                    }

                    pattern = @"(\w+\W\w+\sРеспублика|\w+\sРеспублика|Республика\s\w+|\w+\W+обл|\w+\W+край)";//@"([д]\W+\d+)"; // республика , область, край
                    if (test1(pattern, ful_adress).Length != 0)
                    {
                        array[2] = 1;
                        array_string[2] = test1(pattern, ful_adress);
                        read(test1(pattern, ful_adress), "AV", i, excelappworkbook);
                    }

                    pattern = @"\b(г[.]\w+|г[.]\s+\w+\D\w+|г[.]\s+\w+|село\s\w+|пос[.]\s\w+|город\s\w+|п[.]\w+|г[.]\w+
                                |Москва|Санкт-Петербург|Ярославль|Якутск|Курумкан|Красноярск|Волгодонск|Казань|Иркутск|Димитровград|Жуковский|Кострома|
                                |Выборг|Воронеж|Всеволожск|Брянск|Воркута|Волхов|Волжский|Вологда|Волгоград|Верхний Тагил|Великий Устюг|Великие Луки|Псков|Бобров|
                                |Благовещенск|Камень-на-Оби|Белгород|Тюмень|Архангельск|Балаково|Бабаево|Апатиты|Тула|Старый Оскол|Ставрополь|Северодвинск|Тверь|Самара|
                                |Сыктывкар|Торжок|Ростов-на-Дону|Химки|Улан-Удэ|Уфа|Хабаровск|Людиново|Екатеринбург|Шилка|Чита|Шарья|Таганрог|Череповец|Иваново|Кемерово|
                                |Новосибирск|Гатчина|Мытищи|Альметьевск|Обнинск|Челябинск|Нижний Новгород|Алексин|Балашиха|Балей|Березовский|Владивосток|Великий Новгород|
                                |Владивосток|Каменск-Уральский|Вятские Поляны|Горно-Алтайск|Гуково|Курск|Заречный|Зерноград|Йошкар-Ола|Калининград|Калуга|Каменка|Канаш|
                                |Касимов|Кемерово|Кингисепп|Киров|Кировск|Климовск|Костомукша|Королев|Краснодар|Томск|Красный Кут|Крымск|Кузнецк|Ливны|Липецк|ЛЫСКОВО|Лыткарино|
                                |Магадан|Майкоп|Мамадыш|Маркс|Меленки|Мончегорск|Мурманск|Набережные Челны|Нерюнгри|Нижнекамск|Новомосковск|Оленегорск|Омск|Орёл|Оренбург|Орехово-Зуево|Пенза|
                                |Пермь|Петрозаводск|Петропавловск-Камчатский|Подольск|Одинцово|Рыбинск|Рязань|Сальск|Саранск|Саратов|Саяногорск|Симферополь|Смоленск|Сосенский|Стерлитамак|
                                |Сургут|Сызрань|Орел|Новокузнецк|Приозерск|Северобайкальск|Солнечногорск|Тольятти|Топки|Тосно|Углич|Новочеркасск|Мирный|Ухта|Ханты-Мансийск|Чебоксары)";//город\s\w+|город\w+|г.\s\w+|г.\w+|село\s\w+|пос.\s\w+|Москва)";//город\s\w+|город\w+|г.\s\w+|г.\w+|село\s\w+|пос.\s\w+|Москва)";  // find_City
                    if (test1(pattern, ful_adress).Length != 0)
                    {
                        array[3] = 1;
                        array_string[3] = test1(pattern, ful_adress);
                        read(test1(pattern, ful_adress), "AW", i, excelappworkbook);
                    }

                    pattern = @"\b(ул[.]\s+\w+[.]\s+\w+|Новая Басманная|\w+\s+улица|\w+\s+переулок|\w+\s+ул[.]|ул[.]\w+[.]\w+|ул\s+\w+|ул[.]\s+\w+[.]\w+|
                                |ул[.]\s+\w+\s+\w+|улица\s+\w+\s+\w+|улица\s+\w+|\w+\sпроезд|ул[.]\w+|ул[.]\s+\w+|\w+\sплощадь|\w+\sшоссе|
                                |\w+\sш[.]|\w+\sш|\w+\s+бульвар|\w+\s+пр-д|проезд\s+\w+|б-р[.]\s+\w+|просп[.]\s+\w+|\w+\s+проспект|наб[.]\s+\w+|
                                |\w+\s+пер|\w+\sнабережная|\w+\sнаб[.]|\w+\sнаб|бульвар\s\w+|\w+\s+ул|пр-кт\s+\w+|пр\s+\w+|пр[.]\w+|пр[.]\s+\w+|
                                |пр-т[.]\s+\w+\s+\w+|\w+\s+пр[.]|\w+\s+пр-кт|\w+\s+пр-т|\w+\s+пр-т[.,]|\w+\s+пл[.,]|пр-т\s+\w+|шос[.]\s+\w+|мкр-н\s+\w+\W\d+|
                                |мкр-н\s+\w+|мкр[.,]\s+\w+\W\d+|мкр[.,]\s+\w+|м-н\s+\w+\W\d+|переулок\s+\w+|м-н\s+\w+|б-р\s+\w+|\w+\s+б-р|\w+\s+тракт|\w+\sпереулок|
                                |пл[.]\s+\w+|ш[.]\s+\w+|шоссе\s+\w+|\w+\s+бул[.,]|пер[.]\w+|пер[.]\s+\w+|\w+\sпер[.]|проспект\s\w+|\w+\sвал)";
                    if (test1(pattern, ful_adress).Length != 0)
                    {
                        array[4] = 1;
                        array_string[4] = test1(pattern, ful_adress);
                        read(test1(pattern, ful_adress), "AX", i, excelappworkbook);
                    }

                    pattern = @"(д.\d+\w|д.\s\d+\w|д.\s\d+|д.\d+|д.\d+\s\d+|вл.\s+\d+|вл.\d+|влад.\s+\d+|дом\s\d+\w|дом\s\d+)";  // find_house
                    if (test1(pattern, ful_adress).Length != 0)
                    {
                        array[5] = 1;
                        array_string[5] = test1(pattern, ful_adress);
                        read(test1(pattern, ful_adress), "AY", i, excelappworkbook);
                    }

                    pattern = @"(\w+\W+р-н|\w+\W+район)";// район
                    if (test1(pattern, ful_adress).Length != 0)
                    {
                        array[9] = 1;
                        array_string[9] = test1(pattern, ful_adress);
                        read(test1(pattern, ful_adress), "BC", i, excelappworkbook);
                    }

                    pattern = @"(к\s\d+|к[.]\d+|к[.]\s\d+|корпус\s+\d+|корп[.]\s*\d+|помещение\s+\d+)"; // find_Housing
                    if (test1(pattern, ful_adress).Length != 0)
                    {
                        array[6] = 1;
                        array_string[6] = test1(pattern, ful_adress);
                        read(test1(pattern, ful_adress), "AZ", i, excelappworkbook);
                    }

                    pattern = @"(строение\s+\d+|стр[.]\s+\d+|стр[.]\d+)"; // STROENIE
                    if (test1(pattern, ful_adress).Length != 0)
                    {
                        array[7] = 1;
                        array_string[7] = test1(pattern, ful_adress);
                        read(test1(pattern, ful_adress), "BA", i, excelappworkbook);
                    }

                    pattern = @"(квартира\s*\d*|комн[.]\s\d+|комн[.]\d+|комн\s\d+|комн\d+|помещение\s*\d*|пом[.]\s*\d+|кв[.]\s\d+|кв[.]\d+|кв\s\d+|кв\d+|ком[.]\s\d+|ком[.]\d+|ком\s\d+|ком\d+|оф[.]\s\d+|оф[.]\d+|оф\s\d+|оф\d+)";  // find_Appartament
                   /* if (test1(pattern, ful_adress).Length !=0)
                    {
                        array_string[8] = test1(pattern, ful_adress);
                        read(test1(pattern, ful_adress), "BB", i, excelappworkbook);
                    }
                    */
                    array_string[8] = test1(pattern, ful_adress);
                    read(test1(pattern, ful_adress), "BB", i, excelappworkbook);

                    array_list_answer.Add(array_string);
                    for (int k = 0; k < array_list_answer.LongCount(); k++)
                        for (int j = 0; j < array_list[k].Length; j++)
                        {
                            Console.WriteLine("array_list_answer[{0}]= {1}", j, array_list_answer[k][j]);
                        }
                }
            }
            /*вывод всего листа.
                    for (int k = 0; k < array_list_answer.LongCount(); k++)
                        for (int j = 0; j < array_list[k].Length; j++)
                        {
                            Console.WriteLine("array_list_answer[{0}]= {1}", j, array_list_answer[k][j]);
                        }
                 */
            Console.WriteLine("End");
            Console.ReadKey();

            excelappworkbook.Saved = true;
            excelapp.Quit();
        }
    }
}





   
