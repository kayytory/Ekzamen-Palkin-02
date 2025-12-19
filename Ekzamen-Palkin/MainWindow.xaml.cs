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
using word = Microsoft.Office.Interop.Word;

namespace Ekzamen_Palkin
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public double cost = 0.0;
        public double amountMinutes = 0;
        public const double tariff1 = 0.7;
        public const double tariff2 = 0.3;
        public const double tariffOver = 1.6;
        public int count = 0;
        public MainWindow()
        {
            InitializeComponent();
            List<string> list = new List<string>();
            list.Add("Первый");
            list.Add("Второй");
            cmbTariff.ItemsSource = list.ToList();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try { 
                //проверка на заполнение
            if (txtMinutes.Text == "" || cmbTariff.SelectedIndex == -1 || Convert.ToDouble(txtMinutes.Text) > 100000)
            {
                MessageBox.Show("Введите корректные значения");
                return;
            }
            //в зависимости от тарифа меняет цену
            switch (cmbTariff.SelectedIndex)
            {
                case 0:
                    if (Convert.ToDouble(txtMinutes.Text) < 200)
                    {
                        cost = Convert.ToDouble(txtMinutes.Text) * tariff1;
                    }
                    else
                    {
                        var amountOver = (Convert.ToDouble(txtMinutes.Text) - 200);
                        cost = 200 * tariff1 + amountOver * tariffOver;
                        amountMinutes = amountOver;
                    }
                    break;
                case 1:
                    if (Convert.ToDouble(txtMinutes.Text) < 100)
                    {
                        cost = Convert.ToDouble(txtMinutes.Text) * tariff2;
                    }
                    else
                    {
                        var amountOver = (Convert.ToDouble(txtMinutes.Text) - 100);
                        cost = 100 * tariff2 + amountOver * tariffOver;
                        amountMinutes = amountOver;
                    }
                    break;
            }
            lbMinutes.Content = "Лишние минуты: " + amountMinutes;
            lbPay.Content = "К оплате: " + Math.Round(cost, 2);
            btnWord.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtMinutes_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            //Проверка заполнения на число
            if (!char.IsDigit(e.Text, 0)) e.Handled = true;
        }

        private void btnWord_Click(object sender, RoutedEventArgs e)
        {
            //Сохранение ворд
            try
            {
                count++;
                word.Document document = null;

                word.Application app = new word.Application();

                string putword = Environment.CurrentDirectory.ToString() + @"\Шаблон.docx";

                document = app.Documents.Add(putword);

                document.Activate();

                word.Bookmarks bookm = document.Bookmarks;

                word.Range range;

                string[] data = new string[5] { count.ToString(), DateTime.Now.ToString("dd.MM.yyyy HH:mm"), cmbTariff.Text, amountMinutes.ToString(), Math.Round(cost, 2).ToString() };

                int i = 0;

                foreach (word.Bookmark mark in bookm)

                {

                    range = mark.Range;

                    range.Text = data[i];

                    i++;

                }

                document.SaveAs2(Environment.CurrentDirectory.ToString() + @"\Документ " + count + ".docx");

                document.Close();

                document = null;
                btnWord.IsEnabled = false;
                MessageBox.Show("Распечатан, рассчитайте снова для последующей печати");
            } catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}

