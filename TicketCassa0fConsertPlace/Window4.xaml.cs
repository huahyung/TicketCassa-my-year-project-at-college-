using DotLiquid;
using InteractiveDataDisplay.WPF;
using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using Mysqlx.Crud;
using PdfSharp.Xps;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reactive.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Xps.Serialization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace TicketCassa0fConsertPlace
{
    /// <summary>
    /// Логика взаимодействия для Window4.xaml
    /// </summary>
    public partial class Window4 : System.Windows.Window
    {
        static string serverName = "localhost";
        static string userName = "root";
        static string dbname = "consert_cassa";
        static string port = "3306";
        static string password = "vika_17092007";
        public string ConnectionString = $"Server={serverName};Port={port};Database={dbname};User ID={userName};Password={password};SslMode=None;AllowPublicKeyRetrieval=True;";
        public int IdSotrudnik { get; set; }
        double[] all;
       


        double[] Consecutive(int count)
        {
            return Enumerable.Range(0, count).Select(x => (double)x).ToArray();
        }
        public Window4(int id)
        {
            InitializeComponent();
            IdSotrudnik = id;
        }


        int click = 0;
        public int num;
        public string artist;
        public DateTime data;
        public TimeSpan time;
        public string place;
        public string adress;
        public string type;
        public decimal cost;
        private void numTic_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void return_Click(object sender, RoutedEventArgs e)
        {
            canvas3.Visibility = Visibility.Visible;
            canvas2.Visibility = Visibility.Hidden;
            canvas.Visibility = Visibility.Hidden;
        }

        private void sold_Click(object sender, RoutedEventArgs e)
        {
            canvas3.Visibility = Visibility.Hidden;
            canvas2.Visibility = Visibility.Hidden;
            canvas.Visibility = Visibility.Visible;
        }

        private void main_Click(object sender, RoutedEventArgs e)
        {
            canvas3.Visibility = Visibility.Hidden;
            canvas2.Visibility = Visibility.Visible;
            canvas.Visibility = Visibility.Hidden;
        }

        private void exit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            load();
            ConboboxConc();

        }
        private void ConboboxConc()
        {

            string query = "SELECT IdConsert, Artist FROM Consert";

            System.Data.DataTable dataTable = new System.Data.DataTable();

            using (MySqlConnection connection = new MySqlConnection(ConnectionString))
            {
                MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                adapter.Fill(dataTable);
            }

            conc.ItemsSource = dataTable.DefaultView;
            conc.DisplayMemberPath = "Artist";
            conc.SelectedValuePath = "IdConsert";

        }
        private void TypeticCon()
        {
            int idconc = conc.SelectedIndex + 1;
            string query = $"select tic.Type, ct.Cost, ct.Quanity, tic.idTypeOfTic from consertticket ct, typeoftic tic where idConsert  = {idconc} and ct.IdType = tic.idTypeOfTic;";

            System.Data.DataTable dataTable = new System.Data.DataTable();

            using (MySqlConnection connection = new MySqlConnection(ConnectionString))
            {
                MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                adapter.Fill(dataTable);
            }

            typetic.ItemsSource = dataTable.DefaultView;
            typetic.DisplayMemberPath = "Type";
            typetic.SelectedValuePath = "idTypeOfTic";

        }

        private void conc_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TypeticCon();
        }
        private void ticinfo()
        {
            int idconc = conc.SelectedIndex + 1;
            int idtic = typetic.SelectedIndex + 1;
            string query = $"select DetailsOfTic from ticket tice, consertticket ct, typeoftic tic where idConsert  = {idconc} and {idtic} = tice.idTypeOfTic;";

            System.Data.DataTable dataTable = new System.Data.DataTable();

            using (MySqlConnection connection = new MySqlConnection(ConnectionString))
            {
                MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                adapter.Fill(dataTable);
            }

            info.ItemsSource = dataTable.DefaultView;
            info.DisplayMemberPath = "DetailsOfTic";

        }

        private void typetic_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ticinfo();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (info.Text != "" && conc.SelectedIndex != -1 && typetic.Text != "")
            {
                string details = this.info.Text.ToString();
                int idconc = this.conc.SelectedIndex + 1;
                string idtypetic = typetic.Text.ToString();
                int idsotn = IdSotrudnik;

                string query2 = $"select idTypeOfTic from TypeOfTic where Type = '{idtypetic}';";

                try
                {

                    conn.Open();
                    MySqlCommand command3 = new MySqlCommand();

                    command3.Connection = conn;
                    command3.CommandText = query2;
                    int index = Convert.ToInt32(command3.ExecuteScalar());
                    System.Data.DataTable data = new System.Data.DataTable();
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command3);
                    adapter.Fill(data);
                    int id = (int)data.Rows[0]["idTypeOfTic"];
                    try
                    {
                        string query = $"INSERT INTO Ticket(DetailsOfTic,Consert_idConsert,idTypeOfTic,id_Sotrudniki) VALUES ('{details}',{idconc},{id},{idsotn});";
                        string query1 = $"UPDATE consertticket  SET Quanity = Quanity - 1 WHERE IdConsert = {idconc} AND IdType = {id};";
                        MySqlCommand command1 = new MySqlCommand();

                        command1.Connection = conn;
                        command1.CommandText = query1;

                        try
                        {
                            int updateResult = command1.ExecuteNonQuery();

                            if (updateResult > 0)
                            {
                                MessageBox.Show($"Обновлено записей: {updateResult}");


                                MySqlCommand command2 = new MySqlCommand();
                                command2.Connection = conn;
                                command2.CommandText = query;

                                try
                                {
                                    int insertResult = command2.ExecuteNonQuery();

                                    MessageBox.Show($"В таблицу добавлена {insertResult} запись успешно!");


                                }
                                catch
                                {
                                    MessageBox.Show("Ошибка в добавлении билета");
                                }
                            }
                            else
                            {
                                MessageBox.Show("Не удалось обновить количество билетов. Возможно, билеты закончились.");
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка в обновлении количества билетов");
                        }


                    }
                    catch
                    {
                        MessageBox.Show("Ошибка в добавлении");
                    }
                }
                finally
                {
                    conn.Close();
                }
            }
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void Ticketsret()
        {

            string query = $"select NumTicket from ticket;";

            System.Data.DataTable dataTable = new System.Data.DataTable();

            using (MySqlConnection connection = new MySqlConnection(ConnectionString))
            {
                MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                adapter.Fill(dataTable);
            }

            numruc.ItemsSource = dataTable.DefaultView;
            numruc.DisplayMemberPath = "NumTicket";

        }

        private void canvas3_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (canvas3.Visibility == Visibility.Visible)
            {
                Ticketsret();
            }
            if (canvas3.Visibility == Visibility.Collapsed) { }
            else { }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            
            if (this.numruc.SelectedItem is DataRowView rowView)
            {
                
                int numrucc = Convert.ToInt32(rowView["NumTicket"]);



                string query1 = $"UPDATE consertticket SET Quanity = Quanity + 1 WHERE (IdConsert, IdType) IN (SELECT Consert_idConsert, idTypeOfTic FROM ticket WHERE NumTicket = {numrucc});";
                string query2 = $"DELETE FROM Ticket WHERE NumTicket = {numrucc};";

                try
                {
                    conn.Open();

                   
                    MySqlCommand command1 = new MySqlCommand(query1, conn);
                    int updateResult = command1.ExecuteNonQuery();

                    if (updateResult > 0)
                    {
                        MessageBox.Show($"Обновлено записей: {updateResult}");

                        MySqlCommand command2 = new MySqlCommand(query2, conn);

                        try
                        {
                            int deleteResult = command2.ExecuteNonQuery();

                            if (deleteResult > 0)
                            {
                                MessageBox.Show($"Запись {deleteResult} удалена успешно!");
                            }
                            else
                            {
                                MessageBox.Show("Билет с указанным номером не найден.");
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Ошибка в удалении билета: {ex.Message}");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Не удалось обновить количество билетов. Возможно, билет с указанными параметрами не найден.");
                    }

                    Ticketsret();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}");
                }
                finally
                {
                    if (conn.State == System.Data.ConnectionState.Open)
                    {
                        conn.Close();
                    }
                }
            }
            else { MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error); }
        }
        private void concer()
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);


            string query = $"select tc.NumTicket, c.Artist, c.Data, c.Time, c.Place, c.Adress , t.Type, ct.cost from  consert  c , typeoftic t, consertticket ct, ticket tc where tc.NumTicket = (SELECT MAX(NumTicket) FROM ticket) and t.idtypeoftic = tc.idtypeoftic and tc.consert_idconsert = c.idConsert  and ct.idConsert = c.Idconsert and ct.idType = t.idTypeoftic ;";


            conn.Open();

            MySqlDataAdapter dataAdapter = new MySqlDataAdapter(query, conn);
            System.Data.DataTable dt2 = new System.Data.DataTable();
            dataAdapter.Fill(dt2);


            if (dt2.Rows.Count > 0)
            {
                num = (int)dt2.Rows[0]["Numticket"];
                artist = (string)dt2.Rows[0]["Artist"];
                data = (DateTime)dt2.Rows[0]["Data"];
                time = (TimeSpan)(dt2.Rows[0]["Time"]); ;
                place = (string)dt2.Rows[0]["Place"];
                adress = (string)dt2.Rows[0]["Adress"];
                type = (string)dt2.Rows[0]["Type"];
                cost = (decimal)dt2.Rows[0]["cost"];


            }
        }
        private DotLiquid.Hash CreateTicketContext()
        {

            var context = new
            {
                EventName = artist,
                EventDate = data + " : " + time,
                Venue = place,
                Address = adress,

                TicketType = "Входной билет",
                TicketPrice = cost,
                TicketNumber = num,

                Barcode = "1234567890128",


                AdditionalInfo = new List<dynamic>
        {
            new { Label = "Тип билета", Value = type },
            new { Label = "Место", Value = "-" },

        },

                Rules = new List<string>
        {
            "Вход строго по билетам",
            "Запрещена профессиональная фото/видеосъемка",
            "Администрация вправе отказать во входе без объяснения причин",
            "Билет действителен только на указанную дату"
        }
            };

            return DotLiquid.Hash.FromAnonymousObject(context);
        }
        private void GenerateTicketButton_OnClick(object sender, RoutedEventArgs e)
        {
            concer();
            try
            {
               
                using (var stream = new FileStream("Templates\\ticket.lqd", FileMode.Open))
                using (var reader = new StreamReader(stream))
                {
                    var templateString = reader.ReadToEnd();
                    var template = DotLiquid.Template.Parse(templateString);
                    var ticketContext = CreateTicketContext();
                    var docString = template.Render(ticketContext);

             
                    DocViewer.Document = (FlowDocument)XamlReader.Parse(docString);
                }

           
                GeneratePdf($"concert_ticket{num.ToString()}.pdf");

                MessageBox.Show("Билет успешно создан и сохранен в PDF!", "Успех",
                               MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании билета: {ex.Message}", "Ошибка",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void GeneratePdf(string fileName)
        {
            using (var stream = new FileStream(fileName, FileMode.Create))
            {
                using (var package = Package.Open(stream, FileMode.Create, FileAccess.ReadWrite))
                {
                    using (var xpsDoc = new System.Windows.Xps.Packaging.XpsDocument(package, CompressionOption.Maximum))
                    {
                        var rsm = new XpsSerializationManager(new XpsPackagingPolicy(xpsDoc), false);
                        var paginator = ((IDocumentPaginatorSource)DocViewer.Document).DocumentPaginator;
                        rsm.SaveAsXaml(paginator);
                        rsm.Commit();
                    }
                }

                stream.Position = 0;
                var pdfXpsDoc = PdfSharp.Xps.XpsModel.XpsDocument.Open(stream);
                PdfSharp.Xps.XpsConverter.Convert(pdfXpsDoc, fileName.Replace(".pdf", "_1.pdf"), 0);
            }
        }

        private void load()
        {

            using (MySqlConnection connection = new MySqlConnection(ConnectionString))
            {
                string commandText = $"SELECT   MONTH(c.Data) as Month,   SUM(t.NumTicket) as TotalTickets  FROM consert c   JOIN ticket t ON c.idConsert = t.Consert_idConsert   WHERE YEAR(c.Data) = 2025 or YEAR(c.Data) = 2026  GROUP BY YEAR(c.Data), MONTH(c.Data) ORDER BY YEAR(c.Data), MONTH(c.Data);";
                connection.Open();

                MySqlDataAdapter dataAdapter = new MySqlDataAdapter(commandText, connection);
                System.Data.DataTable dt2 = new System.Data.DataTable();
                dataAdapter.Fill(dt2);





                all = new double[dt2.Rows.Count];

                for (int i = 0; i < dt2.Rows.Count; i++)
                {
                    all[i] = Convert.ToDouble(dt2.Rows[i]["TotalTickets"]);
                }

                var line1 = new InteractiveDataDisplay.WPF.LineGraph
                {
                    Stroke = new SolidColorBrush(Colors.Gray),
                    Description = "Проданные билеты",
                    StrokeThickness = 2
                };
                connection.Close();
                int pointCount = all.Length;
                double[] xs = Consecutive(pointCount);
                double[] ys1 = all;
                line1.Plot(xs, ys1);

                linegraph.Children.Clear();
                linegraph.Children.Add(line1);

            }

        }
        public void showDocument()
        {
            try
            {

                string filePath = ".\\Руководство кассира.docx";


                if (!File.Exists(filePath))
                {
                    MessageBox.Show($"Файл не найден: {filePath}");
                    return;
                }


                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось открыть файл: {ex.Message}");
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            showDocument();
        }


    }
}