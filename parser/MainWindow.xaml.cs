using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Web;
using System.Net;
using Excel =Microsoft.Office.Interop.Excel;
using System.Threading;
using System.IO;
using ExcelDataReader;
using System.Data;
using GemBox.Spreadsheet.Tables;
using ClosedXML.Excel;
using System.Collections;


namespace parser
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
       
        public MainWindow()
        { 
            InitializeComponent();
            Opened();

            ExcelGrid.Cursor = Cursors.IBeam;
            NewOpen2.Cursor = Cursors.Hand;
            F5_2.Cursor = Cursors.ArrowCD;
            Interop2.Cursor = Cursors.ArrowCD;
            Button_Click_Down2.Cursor = Cursors.Hand;
            Button_Click_Next2.Cursor = Cursors.Hand;
            Button_Click_MiniInfo2.Cursor = Cursors.Hand;
            Button_Click_Save2.Cursor = Cursors.ArrowCD;
            Button_Click_Open2.Cursor = Cursors.Hand;
            Pars2.Cursor = Cursors.Help;
            Str.Cursor = Cursors.SizeWE;
        }

        static int page = 0;
        //static int vsego = 0;
       // static int i = 0;
        public static List<Tabl> buffer = new List<Tabl>();
        public static List<String> bufferSTR = new List<String>();
        public static List<List<String>> BUFFER = new List<List<String>>();

        public void Opened()
        {
            try
            {
                var vivod = Tabl.EnumerateTabl("thrlist.xlsx").ToList();
                foreach (var item in vivod)
                {
                    bufferSTR.Add(item.ToString());
                }
                BUFFER.Add(bufferSTR);
                bufferSTR = new List<string>();

                var vivod2 = new List<Tabl>();
                for (int i = 0; i < 15; i++)
                {
                    vivod2.Add(vivod[i]);
                }
                ExcelGrid.ItemsSource = vivod2;
                Str.Content = $"Cтраница 1 из {(int)((vivod.Count - 1) / 15 + 1)}";

                try
                {
                    if (BUFFER[BUFFER.Count - 1].Count > BUFFER[BUFFER.Count - 2].Count)
                    {
                        MessageBox.Show($"Открыто!\nВсего записей: {vivod.Count}\nИз них новых: {BUFFER[BUFFER.Count - 1].Count - BUFFER[BUFFER.Count - 2].Count}\nВ прошлый раз было: {BUFFER[BUFFER.Count - 2].Count} записей\nПроверить различия можно через БЫЛО-СТАЛО.");
                    }
                    else if (BUFFER[BUFFER.Count - 1].Count < BUFFER[BUFFER.Count - 2].Count)
                    {
                        MessageBox.Show($"Открыто!\nВсего записей: {vivod.Count}\nИз базы удалили: {BUFFER[BUFFER.Count - 2].Count - BUFFER[BUFFER.Count - 1].Count}\nВ прошлый раз было: {BUFFER[BUFFER.Count - 2].Count} записей\nПроверить различия можно через БЫЛО-СТАЛО.");
                    }
                    else
                    {
                        MessageBox.Show($"Открыто!\nВсего записей: {vivod.Count}\nВ прошлый раз было: {BUFFER[BUFFER.Count - 2].Count} записей.\nКолличество не изменилось, однако содержание могло поменяться.\nПроверить различия в содержании можно через БЫЛО-СТАЛО.");
                    }
                    MessageBox.Show($"Выведены первые 15 угроз из {vivod.Count}.\nОстальные находятся на следующих страницах.\nВсего страниц {(int)((vivod.Count - 1) / 15 + 1)}");
                }
                catch
                {
                    MessageBox.Show("Эта ваша первая выгрузка локальной базы!\nИстория изменений не знает\nБЫЛО-СТАЛО ничего не ответит\nКак только локальная база измениться, БЫЛО-СТАЛО даст знать!");
                    MessageBox.Show($"Выведены первые 15 угроз из {vivod.Count}.\nОстальные находятся на следующих страницах.\nВсего страниц {(int)((vivod.Count - 1) / 15 + 1)}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Локальной базы данных не существует! Неоходимо нажать кнопку Обновить Базу");
            }
        }

        public void Stranica ()
        {
            var vivod = Tabl.EnumerateTabl("thrlist.xlsx").ToList();
            List<Tabl> vivod2 = new List<Tabl>();

            for (int i = (page) * 15; i < (page + 1) * 15; i++)
            {
                try
                {
                    vivod2.Add(vivod[i]);
                }
                catch (Exception ex)
                {
                    if (i < 1)
                    {
                        MessageBox.Show($"Это была первая страница");
                        break;
                    }
                    else
                    {
                        MessageBox.Show($"Это последняя страница. Выведены последние {i - 15 * page} угроз");
                        break;
                    }
                }
            }

            ExcelGrid.ItemsSource = vivod2;

            if ((page+1<= (int)((vivod.Count - 1) / 15 + 1))&& (page + 1 >= 1))
            {
                Str.Content = $"Cтраница {page + 1} из {(int)((vivod.Count - 1) / 15 + 1)}";
            }
            else
            {
                Str.Content = $"Cтраница 0 из {(int)((vivod.Count - 1) / 15 + 1)}";
            }
           

        }
        public void Parser()
        {
            List<Bylo_Stalo> sravn = new List<Bylo_Stalo>();
            try
            {
                int i = 0;
                int z = 0;
                string ii = "";

                if (BUFFER[BUFFER.Count - 1].Count<= BUFFER[BUFFER.Count - 2].Count) {
                    foreach (var item in BUFFER[BUFFER.Count - 1])
                    { 
                        if (item != BUFFER[BUFFER.Count - 2][i])
                        {
                            z++;
                            Bylo_Stalo b = new Bylo_Stalo(BUFFER[BUFFER.Count - 2][i], item);
                            sravn.Add(b);
                            if (z != 1)
                            {
                                ii += "," + Convert.ToString(i + 1);
                            }
                            else
                            {
                                ii += Convert.ToString(i + 1);
                            }
                        }
                        else
                        {
                            
                        }
                        i++;
                    }
                    if (sravn.Count == 0)
                    {
                        MessageBox.Show("Изменений не было");
                    }
                    else
                    {
                        MessageBox.Show($"Было изменено угроз: {sravn.Count}.\nНомера данных угроз: {ii}");
                        ExcelGrid.ItemsSource = sravn;
                    }
                }
                else
                {
                    MessageBox.Show($"В базу было добавлено {BUFFER[BUFFER.Count - 1].Count - BUFFER[BUFFER.Count - 2].Count} новых угроз!!!\nСейчас я их выведу");

                    for (int u = 0; u < BUFFER[BUFFER.Count - 1].Count - BUFFER[BUFFER.Count - 2].Count; u++)
                    {
                        Bylo_Stalo b = new Bylo_Stalo("-", BUFFER[BUFFER.Count - 1][BUFFER[BUFFER.Count - 1].Count - u - 1]);
                        sravn.Add(b);
                        ExcelGrid.ItemsSource = sravn;
                    }
                }
            }      
           catch 
            {
                MessageBox.Show("Изменений не было");
            }
        }

        private void F5(object sender, RoutedEventArgs e)
        {
            WebClient Down = new WebClient();
            string url = "https://bdu.fstec.ru/files/documents/thrlist.xlsx";
            try
            {
                Down.DownloadFile(url, "thrlist.xlsx");
                MessageBox.Show("Локальная база успешно обновлена!\n(файл успешно сохранён на компьютер)");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Локальная база не обновлена!");

            }
        }

        private void Button_Click_Open(object sender, RoutedEventArgs e)
        {
            Opened();
        }

        private void Button_Click_MiniInfo(object sender, RoutedEventArgs e)
        {
            try
            {
                var vivod = MiniTabl.MiniEnumerateTabl("thrlist.xlsx").ToList();
                ExcelGrid.ItemsSource = vivod;
                MessageBox.Show($"Выведена основная информация по {vivod.Count} угрозам (УИБ)");
                // i = 2;
                Str.Content = $"Cтраница 1 из 1";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Что-то не так!");
            }
        }

        private void Button_Click_Save(object sender, RoutedEventArgs e)
        {

            try 
            {
                MessageBox.Show($"Будет сохранено {ExcelGrid.Items.Count-1} записей");
                FileStream f1 = new FileStream("thrlist.txt", FileMode.Create);
                StreamWriter wrF1 = new StreamWriter(f1);
   
                foreach (var item in ExcelGrid.Items)
                {
                    wrF1.Write(item);
                }

                f1.Close();
                MessageBox.Show("Файл *txt успешно сохранён!\nЕго имя - thrlist.txt.\n(находится там же, где и локальная база)");
            }

                catch (Exception ex)
              {
                    MessageBox.Show(ex.Message);
             }
           
        }

        private void Interop(object sender, RoutedEventArgs e)
        {
            try
            {
                MessageBox.Show($"Будет сохранено {BUFFER[BUFFER.Count-1].Count} записей в Excel");
    
                using (ExcelHelper helper = new ExcelHelper())
                {
                    if (helper.Open(filePath: System.IO.Path.Combine(Environment.CurrentDirectory, "NEWthrlist.xlsx")))
                    {
                        int i = 1;
                   
                        helper.Set(column: "A", row: 1, data: "Общая информация");
                        helper.Set(column: "F", row: 1, data: "Последствия");

                        helper.Set(column: "A", row: 2, data: "Идентификатор УБИ");
                        helper.Set(column: "B", row: 2, data: "Наименование УБИ");
                        helper.Set(column: "C", row: 2, data: "Описание");
                        helper.Set(column: "D", row: 2, data: "Источник угрозы (характеристика и потенциал нарушителя)");
                        helper.Set(column: "E", row: 2, data: "Объект воздействия");
                        helper.Set(column: "F", row: 2, data: "Нарушение конфиденциальности");
                        helper.Set(column: "G", row: 2, data: "Нарушение целостности");
                        helper.Set(column: "H", row: 2, data: "Нарушение доступности");

                        foreach (var item in buffer)
                        {
                            helper.Set(column: "A", row: i, data: item.Id.Trim());
                            helper.Set(column: "B", row: i, data: item.Name.Trim());
                            helper.Set(column: "C", row: i, data: item.Info.Trim());
                            helper.Set(column: "D", row: i, data: item.Sourse.Trim());
                            helper.Set(column: "E", row: i, data: item.Target.Trim());
                            helper.Set(column: "F", row: i, data: item.Conf.Replace("0", "нет").Replace("1", "да").Trim());
                            helper.Set(column: "G", row: i, data: item.Integ.Replace("0", "нет").Replace("1", "да").Trim());
                            helper.Set(column: "H", row: i, data: item.Avail.Replace("0", "нет").Replace("1", "да").Trim());
                            
                            i++;
                            // File.W   WriteLine("NEWthrlist.txt", item.ToString());
                        }
                        helper.Save();
                        MessageBox.Show("Файл Excel успешно сохранён\nНовая база имеет имя NEWthrlist.xlsx и будет находиться там же, где Локальная база\n\nP.S. Размеры колонок и вид придётся редактировать самостоятельно (:");
                    }
                    else
                    {
                        MessageBox.Show("Нет файла");
                    }
                }             
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Pars(object sender, RoutedEventArgs e)
        {
            Parser();
 
        }
    
        private void Button_Click_Next(object sender, RoutedEventArgs e)
        {
            page++;
            Stranica();
        }

        private void Button_Click_Down(object sender, RoutedEventArgs e)
        {
            page--;
            Stranica();
           
        }

        private void NewOpen(object sender, RoutedEventArgs e)
        {
            try
            {
                var vivod = Tabl.EnumerateTabl("thrlist.xlsx").ToList();
                foreach (var item in vivod)
                {
                    //buffer.Add(item);
                    bufferSTR.Add(item.ToString());
                    // MessageBox.Show(item.ToString());
                }
                BUFFER.Add(bufferSTR);
                bufferSTR = new List<string>();

                MessageBox.Show($"Все {vivod.Count} записей выведены в окно"); 

                ExcelGrid.ItemsSource = vivod;
                Str.Content = $"Cтраница 1 из 1";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Не получилось вывести!");
            }
         }
    }
    }
