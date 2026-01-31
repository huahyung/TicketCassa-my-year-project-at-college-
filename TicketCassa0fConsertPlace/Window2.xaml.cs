using InteractiveDataDisplay.WPF;
using MySql.Data.MySqlClient;
using Mysqlx.Crud;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Reactive.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
namespace TicketCassa0fConsertPlace
{
    /// <summary>
    /// Логика взаимодействия для Window2.xaml
    /// </summary>
    public partial class Window2 : Window
    {
        static string serverName = "localhost";
        static string userName = "root";
        static string dbname = "consert_cassa";
        static string port = "3306";
        static string password = "vika_17092007";
        public string ConnectionString = $"Server={serverName};Port={port};Database={dbname};User ID={userName};Password={password};SslMode=None;AllowPublicKeyRetrieval=True;";
        
        string change = "";
        
        public Window2()
        {
            InitializeComponent();
            cl.SelectedDate = DateTime.Today;
            cl.DisplayDate = DateTime.Today;
            /*
            int pointCount = 10_000;
            double[] xs = Consecutive(pointCount);
            double[] ys1 = RandomWalk(pointCount);
            double[] ys2 = RandomWalk(pointCount);
            
            // create the lines and describe their styling
            var line1 = new InteractiveDataDisplay.WPF.LineGraph
            {
                Stroke = new SolidColorBrush(Colors.Blue),
                Description = "Line A",
                StrokeThickness = 2
            };

            var line2 = new InteractiveDataDisplay.WPF.LineGraph
            {
                Stroke = new SolidColorBrush(Colors.Red),
                Description = "Line B",
                StrokeThickness = 2
            };

            // load data into the lines
            line1.Plot(xs, ys1);
            line2.Plot(xs, ys2);

            // add lines into the grid
            linegraph.Children.Clear();
            linegraph.Children.Add(line1);
            linegraph.Children.Add(line2);
            
            */  

        }
        
        int click = 0;
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            click++;
            if (click == 1)
            {
               
            } else if (click == 2) {
                menu.Visibility = Visibility.Hidden; click = 0;
            }
        }

        private void concerts_Click(object sender, RoutedEventArgs e)
        {
            change = "consert_clicked";
            try
            {
                consertsb.Visibility = Visibility.Visible;
                ticketsb.Visibility = Visibility.Hidden;
                sotrudB.Visibility = Visibility.Hidden;
                doljd.Visibility = Visibility.Hidden;
                typeoftic.Visibility = Visibility.Hidden;
                canvas4.Visibility = Visibility.Hidden;

                canvas.Visibility = Visibility.Visible;
                canvas2.Visibility = Visibility.Hidden;
                canvas3.Visibility = Visibility.Hidden;
                //funccc.Visibility = Visibility.Visible;

                using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                {
                    string commandText = "SELECT * FROM consert;";
                    connection.Open();

                    using (MySqlCommand mySqlCommand = new MySqlCommand(commandText, connection))
                    {
                        DataSet data = new DataSet();
                        MySqlDataAdapter adapter = new MySqlDataAdapter(mySqlCommand);
                        
                        adapter.Fill(data, "consert");
                        datagrid.ItemsSource = data.Tables["consert"].DefaultView;
                        datagrid.Columns[0].Header = "#";
                        datagrid.Columns[1].Header = "Дата проведения";
                        datagrid.Columns[2].Header = "Время проведения";
                        datagrid.Columns[3].Header = "Артист";
                        datagrid.Columns[4].Header = "Адрес";
                        datagrid.Columns[5].Header = "Вид";
                        datagrid.Columns[6].Header = "Место проведения";
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading concerts: {ex.Message}", "Error",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
           
            
        }

        private void tickets_Click(object sender, RoutedEventArgs e)
        {
            change = "ticket_clicked";
            try {
                consertsb.Visibility = Visibility.Hidden;
                ticketsb.Visibility = Visibility.Visible;
                sotrudB.Visibility = Visibility.Hidden;
                doljd.Visibility = Visibility.Hidden;
                typeoftic.Visibility = Visibility.Hidden;

                canvas4.Visibility = Visibility.Hidden;

                canvas.Visibility = Visibility.Visible;
                canvas2.Visibility = Visibility.Hidden;
                canvas3.Visibility = Visibility.Hidden;
                //funccc.Visibility = Visibility.Visible;

                using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                {
                    string commandText = "SELECT * FROM ticket;";
                    connection.Open();

                    using (MySqlCommand mySqlCommand = new MySqlCommand(commandText, connection))
                    {
                        DataSet data = new DataSet();
                        MySqlDataAdapter adapter = new MySqlDataAdapter(mySqlCommand);
                        adapter.Fill(data,"ticket");
                        datagrid.ItemsSource = data.Tables["ticket"].DefaultView;
                        datagrid.Columns[0].Header = "#";
                        datagrid.Columns[1].Header = "Информация";
                        datagrid.Columns[2].Header = "# концерта";
                        datagrid.Columns[3].Header = "# типа билета";
                        datagrid.Columns[4].Header = "# сотрудника";
                        
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading concerts: {ex.Message}", "Error",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void typetic_Click(object sender, RoutedEventArgs e)
        {
            change = "typeoftic_clicked";
            
            try
            {
                consertsb.Visibility = Visibility.Hidden;
                ticketsb.Visibility = Visibility.Hidden;
                sotrudB.Visibility = Visibility.Hidden;
                doljd.Visibility = Visibility.Hidden;
                typeoftic.Visibility = Visibility.Visible;

                canvas4.Visibility = Visibility.Hidden;
                canvas.Visibility = Visibility.Visible;
                canvas2.Visibility = Visibility.Hidden;
                canvas3.Visibility = Visibility.Hidden;
                //funccc.Visibility = Visibility.Visible;

                using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                {
                    string commandText = "SELECT * FROM typeoftic;";
                    connection.Open();

                    using (MySqlCommand mySqlCommand = new MySqlCommand(commandText, connection))
                    {
                        DataSet data = new DataSet();
                        MySqlDataAdapter adapter = new MySqlDataAdapter(mySqlCommand);
                        adapter.Fill(data, "typeoftic");
                        datagrid.ItemsSource = data.Tables["typeoftic"].DefaultView;
                        datagrid.Columns[0].Header = "#";
                        datagrid.Columns[1].Header = "Тип";
                        datagrid.Columns[2].Header = "Стоимость";
                        datagrid.Columns[3].Header = "Количество";
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading concerts: {ex.Message}", "Error",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void sotr_Click(object sender, RoutedEventArgs e)
        {
            change = "sotrudniki_clicked";
            try
            {
                consertsb.Visibility = Visibility.Hidden;
                ticketsb.Visibility = Visibility.Hidden;
                sotrudB.Visibility = Visibility.Visible;
                doljd.Visibility = Visibility.Hidden;
                typeoftic.Visibility = Visibility.Hidden;
                canvas4.Visibility = Visibility.Hidden;

                canvas.Visibility = Visibility.Visible;
                canvas2.Visibility = Visibility.Hidden;
                //funccc.Visibility = Visibility.Visible;
                canvas3.Visibility = Visibility.Hidden;

                using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                {
                    string commandText = "SELECT id_Sotrudniki, login, cast(aes_decrypt(password,'mysecretkey')as Char), Familia,Name, Otchestvo, IdDoljnost FROM sotrudniki;";
                    connection.Open();

                    using (MySqlCommand mySqlCommand = new MySqlCommand(commandText, connection))
                    {
                        DataSet data = new DataSet();
                        MySqlDataAdapter adapter = new MySqlDataAdapter(mySqlCommand);
                        adapter.Fill(data, "sotrudniki");
                        datagrid.ItemsSource = data.Tables["sotrudniki"].DefaultView;
                        datagrid.Columns[0].Header = "#";
                        datagrid.Columns[1].Header = "Логин";
                        datagrid.Columns[2].Header = "Пароль";
                        datagrid.Columns[3].Header = "Фамилия";
                        datagrid.Columns[4].Header = "Имя";
                        datagrid.Columns[5].Header = "Отчество";
                        datagrid.Columns[6].Header = "# должности";
                        
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading concerts: {ex.Message}", "Error",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void dolj_Click(object sender, RoutedEventArgs e)
        {
            change =  "dolj_clicked";
            try
            {
                consertsb.Visibility = Visibility.Hidden;
                ticketsb.Visibility = Visibility.Hidden;
                sotrudB.Visibility = Visibility.Hidden;
                doljd.Visibility = Visibility.Visible;
                typeoftic.Visibility = Visibility.Hidden;
                canvas4.Visibility = Visibility.Hidden;
                canvas.Visibility = Visibility.Visible;
                canvas2.Visibility = Visibility.Hidden;
                canvas3.Visibility = Visibility.Hidden;
                //funccc.Visibility = Visibility.Visible;

                using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                {
                    string commandText = "SELECT * FROM doljnost;";
                    connection.Open();

                    using (MySqlCommand mySqlCommand = new MySqlCommand(commandText, connection))
                    {
                        DataSet data = new DataSet();
                        MySqlDataAdapter adapter = new MySqlDataAdapter(mySqlCommand);
                        adapter.Fill(data, "doljnost");
                        datagrid.ItemsSource = data.Tables["doljnost"].DefaultView;
                        DataGridTextColumn textcol = new DataGridTextColumn();
                        datagrid.Columns[0].Header = "#";
                        datagrid.Columns[1].Header = "Должность";



                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading concerts: {ex.Message}", "Error",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

       

        private void mainw_Click(object sender, RoutedEventArgs e)
        {
            consertsb.Visibility = Visibility.Hidden;
            ticketsb.Visibility = Visibility.Hidden;
            sotrudB.Visibility = Visibility.Hidden;
            doljd.Visibility = Visibility.Hidden;
            typeoftic.Visibility = Visibility.Hidden;
           
            canvas.Visibility = Visibility.Hidden;
            canvas2.Visibility = Visibility.Visible;
            canvas3.Visibility = Visibility.Hidden;
            canvas4.Visibility = Visibility.Hidden;
            //funccc.Visibility = Visibility.Hidden;
        }

        
        
        private void MenuItem_MouseEnter(object sender, MouseEventArgs e)
        {
            var menuItem = sender as MenuItem;
            menuItem?.SetCurrentValue(MenuItem.IsSubmenuOpenProperty, true);
        }

        private void Button_Click_exit(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ticketbut_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int numticc = int.Parse(this.numTic.Text);
            string details = this.detailtic.Text;
            int idconc = int.Parse(this.idconserttic.Text);
            int idtypetic = int.Parse(this.idtypetictick.Text);
            int idsotrr = int.Parse(this.idsotrtick.Text);
            string query = $"INSERT INTO Ticket VALUES ({numticc},'{details}',{idconc},{idtypetic},{idsotrr});";
            try
            {
               
                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"В таблицу добавлена {number} запись успешно!");
                

            }
            catch
            {
                MessageBox.Show("Ошибка в добавлении");
            }
            finally
            {
                conn.Close();
            }
        }

        private void ticketupdate_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int numticc = int.Parse(this.numTic.Text);
            string details = this.detailtic.Text;
            int idconc = int.Parse(this.idconserttic.Text);
            int idtypetic = int.Parse(this.idtypetictick.Text);
            int idsotrr = int.Parse(this.idsotrtick.Text);
            string query = $"update  Ticket  set  DetailsOfTic ='{details}',Consert_idConsert ={idconc},idTypeOfTic={idtypetic},id_Sotrudniki ={idsotrr} where  NumTicket = {numticc};";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"Запись {number} обновлена запись успешно!");
                DeleteTextBox();

            }
            catch
            {
                MessageBox.Show("Ошибка в изменении");
            }
            finally
            {
                conn.Close();
            }
        }

        private void ticketdelete_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int numticc = int.Parse(this.numTic.Text);
           
            string query = $"delete  from  Ticket  where  NumTicket = {numticc};";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"Запись {number} удалена запись успешно!");
                DeleteTextBox();

            }
            catch
            {
                MessageBox.Show("Ошибка в удалении");
            }
            finally
            {
                conn.Close();
            }
        }

        private void consertadd_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int idcons = int.Parse(this.idconser.Text);
            DateTime datacons = Convert.ToDateTime(this.timeconsert.Text);
            string timee = this.time.Text;
            string artistt = this.artist.Text;
            string adresss = this.adress.Text;
            string typeconc = this.typecon.Text;
            string placen = this.place.Text;
            string query = $"INSERT INTO Consert VALUES ({idcons},'{datacons.ToString("yyyy-MM-dd")}','{timee}','{artistt}','{adresss}', '{typeconc}','{placen}');";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"В таблицу добавлена {number} запись успешно!");
                DeleteTextBox();

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString(),"Ошибка в добавлении");
            }
            finally
            {
                conn.Close();
            }
        }

        private void updateconc_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int idcons = int.Parse(this.idconser.Text);
            DateTime datacons = Convert.ToDateTime(this.timeconsert.Text);
            string timee = this.time.Text;
            string artistt = this.artist.Text;
            string adresss = this.adress.Text;
            string typeconc = this.typecon.Text;
            string placen = this.place.Text;
            string query = $"update Consert set  Data='{datacons.ToString("yyyy-MM-dd")}',Time='{timee}',Artist ='{artistt}', Adress='{adresss}',Type= '{typeconc}', Place='{placen}' where idConsert={idcons} ;";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"Запись  {number} изменена успешно!");
                DeleteTextBox();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка в изменении");
            }
            finally
            {
                conn.Close();
            }
        }

        private void deleteconc_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int idcons = int.Parse(this.idconser.Text);
           
            string query = $"delete from Consert where idConsert={idcons} ;";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"Запись  {number} удалена успешно!");
                DeleteTextBox();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка в удалении");
            }
            finally
            {
                conn.Close();
            }
        }

        private void addsotr_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int idsotrud = int.Parse(this.idsotr.Text);
            string loginn = this.login.Text;
            string passwordik = this.passwordsotr.Text;
            string fam = this.fdmilia.Text;
            string namee = this.name.Text;
            string otch = this.otchest.Text;
            int iddolg = int.Parse(this.iddoljsotr.Text);
            string query = $"INSERT INTO sotrudniki VALUES ({idsotrud},'{loginn}','{passwordik}','{fam}','{namee}', '{otch}',{iddolg});";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"В таблицу добавлена {number} запись успешно!");
                DeleteTextBox();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка в добавлении");
            }
            finally
            {
                conn.Close();
            }
        }

        private void updatesotr_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int idsotrud = int.Parse(this.idsotr.Text);
            string loginn = this.login.Text;
            string passwordik = this.passwordsotr.Text;
            string fam = this.fdmilia.Text;
            string namee = this.name.Text;
            string otch = this.otchest.Text;
            int iddolg = int.Parse(this.iddoljsotr.Text);
            string query = $"update  sotrudniki set  Login ='{loginn}', Password ='{passwordik}', Familia ='{fam}',Name='{namee}', Otchestvo ='{otch}',idDoljnost={iddolg} where id_Sotrudniki = {idsotrud};";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"Запись  {number} изменена успешно!");
                DeleteTextBox();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка в изменении");
            }
            finally
            {
                conn.Close();
            }
        }

        private void deletesotr_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int idsotrud = int.Parse(this.idsotr.Text);
            
            string query = $"delete from  sotrudniki  where id_Sotrudniki = {idsotrud};";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"Запись  {number} удалена успешно!");
                DeleteTextBox();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка в удалении");
            }
            finally
            {
                conn.Close();
            }
        }

        private void dolgadd_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int iddolgd = int.Parse(this.iddolg.Text);
            string naimen = this.namedolj.Text;
           
            string query = $"INSERT INTO doljnost VALUES ({iddolgd},'{naimen}');";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"В таблицу добавлена {number} запись успешно!");
                DeleteTextBox();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка в добавлении");
            }
            finally
            {
                conn.Close();
            }
        }

        private void dolgupd_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int iddolgd = int.Parse(this.iddolg.Text);
            string naimen = this.namedolj.Text;

            string query = $"update  doljnost set DoljnostName ='{naimen}' where idDoljnost = {iddolgd};";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"Запись  {number} изменена успешно!");
                DeleteTextBox();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка в изменении");
            }
            finally
            {
                conn.Close();
            }
        }

        private void dolgdel_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int iddolgd = int.Parse(this.iddolg.Text);
            string naimen = this.namedolj.Text;

            string query = $"delete from  doljnost  where idDoljnost = {iddolgd};";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"Запись  {number} удалена успешно!");
                DeleteTextBox();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка в удалена");
            }
            finally
            {
                conn.Close();
            }
        }

        private void addtypetic_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int idtt = int.Parse(this.idtypeoftic.Text);
            string tytic = this.nametypetic.Text;
            int costt = int.Parse(this.cost.Text);
            int col = int.Parse(this.quanity.Text);
            string query = $"INSERT INTO typeoftic VALUES ({idtt},'{tytic}',{costt},{col});";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"В таблицу добавлена {number} запись успешно!");
                DeleteTextBox();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка в добавлении");
            }
            finally
            {
                conn.Close();
            }
        }

        private void updatetictype_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int idtt = int.Parse(this.idtypeoftic.Text);
            string tytic = this.nametypetic.Text;
            int costt = int.Parse(this.cost.Text);
            int col = int.Parse(this.quanity.Text);
            string query = $"update  typeoftic set Type ='{tytic}',Cost = {costt},Quantity={col} where idTypeOfTic = {idtt};";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"Запись  {number} изменена успешно!");
                DeleteTextBox();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка в изменении");
            }
            finally
            {
                conn.Close();
            }
        }

        private void deletetypetic_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            int idtt = int.Parse(this.idtypeoftic.Text);
            
            string query = $"delete from  typeoftic  where idTypeOfTic = {idtt};";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"Запись  {number} удалена успешно!");
                DeleteTextBox();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка в удалении");
            }
            finally
            {
                conn.Close();
            }
        }

       
        private void DeleteTextBox()
        {
           
            numTic.Text = "";
            detailtic.Text = "";
            idconserttic.Text = "";
            idtypetictick.Text = "";
            idsotrtick.Text = "";
            idconser.Text = "";
            time.Text = "";
            artist.Text = "";
            adress.Text = "";
            typecon.Text = "";
            place.Text = "";
            idsotr.Text = "";
            login.Text = "";
            passwordsotr.Text = "";
            fdmilia.Text = "";
            name.Text = "";
            otchest.Text = "";
            iddoljsotr.Text = "";
            iddolg.Text = "";
            namedolj.Text = "";
            idtypeoftic.Text = "";
            nametypetic.Text = "";
            cost.Text = "";
            quanity.Text = "";
        }

        private void datagrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                DataRowView dataRowView = datagrid.SelectedItem as DataRowView;
                switch (change)
                {
                    case "consert_clicked":
                        if (datagrid.SelectedIndex != -1)
                        {
                            idconser.Text = dataRowView[0].ToString();
                            timeconsert.Text = dataRowView[1].ToString();
                            time.Text = dataRowView[2].ToString();
                            artist.Text = dataRowView[3].ToString();
                            adress.Text = dataRowView[4].ToString();
                            typecon.Text = dataRowView[5].ToString();
                            place.Text = dataRowView[6].ToString();
                        }

                        break;
                    case "ticket_clicked":
                        if (datagrid.SelectedIndex != -1)
                        {
                            numTic.Text = dataRowView[0].ToString();
                            detailtic.Text = dataRowView[1].ToString();
                            idconserttic.Text = dataRowView[2].ToString();
                            idtypetictick.Text = dataRowView[3].ToString();
                            idsotrtick.Text = dataRowView[4].ToString();
                        }
                        break;
                    case "typeoftic_clicked":
                        if (datagrid.SelectedIndex != -1)
                        {
                            idtypeoftic.Text = dataRowView[0].ToString();
                            nametypetic.Text = dataRowView[1].ToString();
                            cost.Text = dataRowView[2].ToString();
                            quanity.Text = dataRowView[3].ToString();

                        }
                        break;
                    case "sotrudniki_clicked":
                        if (datagrid.SelectedIndex != -1)
                        {

                            idsotr.Text = dataRowView[0].ToString();
                            login.Text = dataRowView[1].ToString();
                            passwordsotr.Text = dataRowView[2].ToString();
                            fdmilia.Text = dataRowView[3].ToString();
                            name.Text = dataRowView[4].ToString();
                            otchest.Text = dataRowView[5].ToString();
                            iddoljsotr.Text = dataRowView[6].ToString();
                        }
                        break;

                    case "dolj_clicked":
                        if(datagrid.SelectedIndex != -1)
                        {
                            iddolg.Text = dataRowView[0].ToString();
                            namedolj.Text = dataRowView[1].ToString();
                        }
                        
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void poseh_Click(object sender, RoutedEventArgs e)
        {
            canvas3.Visibility = Visibility.Visible;
            canvas2.Visibility = Visibility.Hidden;
            canvas.Visibility = Visibility.Hidden;
            canvas4.Visibility = Visibility.Hidden;
        }

        private void prib_Click(object sender, RoutedEventArgs e)
        {
            canvas4.Visibility = Visibility.Visible;
            canvas2.Visibility = Visibility.Hidden;
            canvas.Visibility = Visibility.Hidden;
            canvas3.Visibility = Visibility.Hidden;
        }

        

        private void otchetdatapost_Click(object sender, RoutedEventArgs e)
        {
            datefirstpos
        }
    }
}
