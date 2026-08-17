using DotLiquid;
using InteractiveDataDisplay.WPF;
using Microsoft.Win32;
using MySql.Data.MySqlClient;
using Mysqlx.Crud;
using ScottPlot.TickGenerators.TimeUnits;
using System;
using System.Collections;
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
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Xps.Serialization;
using static System.Windows.Forms.LinkLabel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using Excel = Microsoft.Office.Interop.Excel;
using MessageBox = System.Windows.MessageBox;

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
        double[] all;
        double[] all3;
        string change = "";
        string change2 = "";
        public int num1;
        public string artist1;
        public DateTime data1;
        public TimeSpan time1;
        public string place1;
        public string adress1;
        public string type1;
        public decimal cost1;
        public int IdSotrudnik { get; set; }
        double[] Consecutive(int count)
        {
            return Enumerable.Range(0, count).Select(x => (double)x).ToArray();
        }

        
        public Window2(int id)
        {
            InitializeComponent();
            IdSotrudnik = id;
            cl.SelectedDate = DateTime.Today;
            cl.DisplayDate = DateTime.Today;
            

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
        public void ConssertsDT()
        {
            try
            {
                ChangeCanva();
                
                
                

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
                    DeleteTextBox();
                }
}
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading concerts: {ex.Message}", "Error",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void concerts_Click(object sender, RoutedEventArgs e)
        {
            change = "consert_clicked";
            change2 = "canvas";
            ConssertsDT();



        }

       
        private void TicketsDT()
        {
            try
            {
                ChangeCanva();

                
               

                using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                {
                    string commandText = "SELECT * FROM ticket;";
                    connection.Open();

                    using (MySqlCommand mySqlCommand = new MySqlCommand(commandText, connection))
                    {
                        DataSet data = new DataSet();
                        MySqlDataAdapter adapter = new MySqlDataAdapter(mySqlCommand);
                        adapter.Fill(data, "ticket");
                        datagrid.ItemsSource = data.Tables["ticket"].DefaultView;
                        datagrid.Columns[0].Header = "#";
                        datagrid.Columns[1].Header = "Информация";
                        datagrid.Columns[2].Header = "# концерта";
                        datagrid.Columns[3].Header = "# типа билета";
                        datagrid.Columns[4].Header = "# сотрудника";

                    }
                    DeleteTextBox();
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading concerts: {ex.Message}", "Error",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void tickets_Click(object sender, RoutedEventArgs e)
        {
            change = "ticket_clicked";
            change2 = "canvas";
            TicketsDT();
            
        }
        private void TypeTicDT()
        {
            try
            {
                ChangeCanva();
               

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
                        
                    }
                    connection.Close();
                    DeleteTextBox();
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
            change2 = "canvas";
            TypeTicDT();
            
        }
        private void SotrDT()
        {
            try
            {
                ChangeCanva();

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
                    DeleteTextBox();
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
            change2 = "canvas";
            SotrDT();
        }
        private void DoljDt()
        {
            try
            {
                ChangeCanva();

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
                    DeleteTextBox();
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
            change = "dolj_clicked";
            change2 = "canvas";
            DoljDt();
        }

       

        private void mainw_Click(object sender, RoutedEventArgs e)
        {
            
            change2 = "main";
            ChangeCanva();
            DeleteTextBox();
            
        }

        
        
        private void MenuItem_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var menuItem = sender as System.Windows.Controls.MenuItem;
            menuItem?.SetCurrentValue(System.Windows.Controls.MenuItem.IsSubmenuOpenProperty, true);
        }

        private void Button_Click_exit(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ticketbut_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.numTic.Text != "" && detailtic.Text != "" && this.idconserttic.Text != "" && this.idtypetictick.Text != "" && this.idsotrtick.Text != "") { 
            
            int numticc = int.Parse(this.numTic.Text);
            string details = this.detailtic.Text;
            int idconc = int.Parse(this.idconserttic.Text);
            int idtypetic = int.Parse(this.idtypetictick.Text);
            int idsotrr = int.Parse(this.idsotrtick.Text);
            string query = $"INSERT INTO Ticket VALUES ({numticc},'{details}',{idconc},{idtypetic},{idsotrr});";
            string query1 = $"UPDATE consertticket  SET Quanity = Quanity - 1 WHERE IdConsert = {idconc} AND IdType = {idtypetic};";
            
                try
                {

                    conn.Open();
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

                                TicketsDT();
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
                    TicketsDT();

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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ticketupdate_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.numTic.Text != "" && detailtic.Text != "" && this.idconserttic.Text != "" && this.idtypetictick.Text != "" && this.idsotrtick.Text != "")
            {
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
                    TicketsDT();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }

        private void ticketdelete_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.numTic.Text != ""  && this.idconserttic.Text != "" && this.idtypetictick.Text != "" )
            {
                int numticc = int.Parse(this.numTic.Text);
                int idconc = int.Parse(this.idconserttic.Text);
                int idtypetic = int.Parse(this.idtypetictick.Text);
                string query1 = $"UPDATE consertticket  SET Quanity = Quanity + 1 WHERE IdConsert = {idconc} AND IdType = {idtypetic};";
                string query = $"delete  from  Ticket  where  NumTicket = {numticc};";
           
                try
                {

                    conn.Open();
                    MySqlCommand command1 = new MySqlCommand();
                    command1.Connection = conn;
                    command1.CommandText = query1;
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
                            MessageBox.Show($"Из таблицы запись {insertResult} удалена успешно!");

                            TicketsDT();
                        }
                        catch
                        {
                            MessageBox.Show("Ошибка в удалении билета");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Не удалось обновить количество билетов. ");
                    }

                    DeleteTextBox();
                    TicketsDT();

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
            else
            {
                 MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void consertadd_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.idconser.Text != "" && this.artist.Text != "" && this.timeconsert.Text != "" && this.time.Text != "" && this.adress.Text != "" && this.typecon.Text != "" && this.place.Text != "") 
            {
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
                    ConssertsDT();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            
        }
        }

        private void updateconc_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.idconser.Text != "" && this.artist.Text != "" && this.timeconsert.Text != "" && this.time.Text != "" && this.adress.Text != "" && this.typecon.Text != "" && this.place.Text != "")
            {
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
                    ConssertsDT();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void deleteconc_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            
            if (this.idconser.Text != "")
            {
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
                    ConssertsDT();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
             
        }

        private void addsotr_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.idsotr.Text != "" && this.login.Text != "" && this.passwordsotr.Text != "" && this.fdmilia.Text != "" && this.name.Text != "" && this.otchest.Text != "" && this.iddoljsotr.Text != "")
            {
                int idsotrud = int.Parse(this.idsotr.Text);
                string loginn = this.login.Text;
                string passwordik = this.passwordsotr.Text;
                string fam = this.fdmilia.Text;
                string namee = this.name.Text;
                string otch = this.otchest.Text;
                int iddolg = int.Parse(this.iddoljsotr.Text);
                string query = $"INSERT INTO sotrudniki VALUES ({idsotrud},'{loginn}',AES_ENCRYPT('{passwordik}','mysecretkey'),'{fam}','{namee}', '{otch}',{iddolg});";
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void updatesotr_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.idsotr.Text != "" && this.login.Text != "" && this.passwordsotr.Text != "" && this.fdmilia.Text != "" && this.name.Text != "" && this.otchest.Text != "" && this.iddoljsotr.Text != "")
            {
                int idsotrud = int.Parse(this.idsotr.Text);
                string loginn = this.login.Text;
                string passwordik = this.passwordsotr.Text;
                string fam = this.fdmilia.Text;
                string namee = this.name.Text;
                string otch = this.otchest.Text;
                int iddolg = int.Parse(this.iddoljsotr.Text);
                string query = $"update  sotrudniki set  Login ='{loginn}', Password =AES_ENCRYPT('{passwordik}','mysecretkey'), Familia ='{fam}',Name='{namee}', Otchestvo ='{otch}',idDoljnost={iddolg} where id_Sotrudniki = {idsotrud};";
                try
                {

                    conn.Open();
                    MySqlCommand command = new MySqlCommand();
                    command.Connection = conn;
                    command.CommandText = query;
                    int number = command.ExecuteNonQuery();

                    MessageBox.Show($"Запись  {number} изменена успешно!");
                    DeleteTextBox();
                    SotrDT();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void deletesotr_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.idsotr.Text != "" )
            {
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
                    SotrDT();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void dolgadd_Click(object sender, RoutedEventArgs e)
        {

            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.iddolg.Text != "" && this.namedolj.Text != "")
            {
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
                    SotrDT();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
         }

        private void dolgupd_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.iddolg.Text != "" && this.namedolj.Text != "")
            {
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
                    DoljDt();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void dolgdel_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.iddolg.Text != "")
            {
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
                    DoljDt();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void addtypetic_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.idtypeoftic.Text != "" && this.nametypetic.Text != "")
            {
                int idtt = int.Parse(this.idtypeoftic.Text);
                string tytic = this.nametypetic.Text;

                string query = $"INSERT INTO typeoftic VALUES ({idtt},'{tytic}');";
                try
                {

                    conn.Open();
                    MySqlCommand command = new MySqlCommand();
                    command.Connection = conn;
                    command.CommandText = query;
                    int number = command.ExecuteNonQuery();

                    MessageBox.Show($"В таблицу добавлена {number} запись успешно!");
                    DeleteTextBox();
                    TypeTicDT();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void updatetictype_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.idtypeoftic.Text != "" && this.nametypetic.Text != "")
            {
                int idtt = int.Parse(this.idtypeoftic.Text);
                string tytic = this.nametypetic.Text;
            
            string query = $"update  typeoftic set Type ='{tytic}' where idTypeOfTic = {idtt};";
            try
            {

                conn.Open();
                MySqlCommand command = new MySqlCommand();
                command.Connection = conn;
                command.CommandText = query;
                int number = command.ExecuteNonQuery();

                MessageBox.Show($"Запись  {number} изменена успешно!");
                DeleteTextBox();
                TypeTicDT();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void deletetypetic_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.idtypeoftic.Text != "")
            {
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
                    TypeTicDT();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

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
            lol2.Text = "";
            lol3.Text = "";
            lol4.Text = "";
            lol5.Text = "";
            lol6.Text = "";

          
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
                    case "conserttic_clicked":
                        if (datagrid.SelectedIndex != -1)
                        {
                            lol2.Text = dataRowView[0].ToString();
                            lol3.Text = dataRowView[1].ToString();
                            lol4.Text = dataRowView[2].ToString();
                            lol5.Text = dataRowView[3].ToString();
                            lol6.Text = dataRowView[4].ToString();
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
            
            change2 = "poseh";
            ChangeCanva();
        }

        private void prib_Click(object sender, RoutedEventArgs e)
        {
            
            change2 = "prib";
            ChangeCanva();
        }

        

        private void otchetdatapost_Click(object sender, RoutedEventArgs e)
        {
            if (this.yeaars.Text != "" && monthcomn.Text != "")
            {
                int year = int.Parse(comyear.Text);
                int month = int.Parse(monthcomn.Text);

                try
                {


                    using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                    {
                        string commandText = $"SELECT c.Artist AS 'Артист',    c.Data AS 'Дата концерта',   c.Place AS 'Место проведения',  COUNT(t.NumTicket) AS 'Количество проданных билетов' FROM consert c JOIN ticket t ON c.idConsert = t.Consert_idConsert  WHERE YEAR(c.Data) = {year}    AND MONTH(c.Data) ={month}  GROUP BY YEAR(c.Data), MONTH(c.Data), c.idConsert ORDER BY c.Data;";
                        connection.Open();

                        using (MySqlCommand mySqlCommand = new MySqlCommand(commandText, connection))
                        {
                            DataSet data = new DataSet();
                            MySqlDataAdapter adapter = new MySqlDataAdapter(mySqlCommand);
                            adapter.Fill(data, "otchet");
                            datagrid2.ItemsSource = data.Tables["otchet"].DefaultView;


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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
            }

        private void otchetyearpos_Click(object sender, RoutedEventArgs e)
        {
            if (this.yeaars.Text != "")
            {

                int year = int.Parse(yeaars.Text);

                try
                {


                    using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                    {
                        string commandText = $"SELECT c.Artist AS 'Артист',    c.Data AS 'Дата концерта',   c.Place AS 'Место проведения',  COUNT(t.NumTicket) AS 'Количество проданных билетов' FROM consert c JOIN ticket t ON c.idConsert = t.Consert_idConsert  WHERE YEAR(c.Data) = {year}      GROUP BY YEAR(c.Data), MONTH(c.Data), c.idConsert ORDER BY c.Data;";
                        connection.Open();

                        using (MySqlCommand mySqlCommand = new MySqlCommand(commandText, connection))
                        {
                            DataSet data = new DataSet();
                            MySqlDataAdapter adapter = new MySqlDataAdapter(mySqlCommand);
                            adapter.Fill(data, "otchet");
                            datagrid2.ItemsSource = data.Tables["otchet"].DefaultView;


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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void posofmonth_Click(object sender, RoutedEventArgs e)
        {
            monthes.Visibility = Visibility.Visible;
            years.Visibility = Visibility.Hidden;
            quarter.Visibility = Visibility.Hidden;
            
        }

        private void posofkvart_Click(object sender, RoutedEventArgs e)
        {

            quarter.Visibility = Visibility.Visible;
            years.Visibility = Visibility.Hidden;
            monthes.Visibility = Visibility.Hidden;
            
        }

        private void posofyear_Click(object sender, RoutedEventArgs e)
        {
            years.Visibility = Visibility.Visible;
            monthes.Visibility = Visibility.Hidden;
            quarter.Visibility = Visibility.Hidden;
            
        }

        private void otchetquarter_Click(object sender, RoutedEventArgs e)
        {
            if (this.quarters.Text != "" && this.yeaarss.Text != "")
            {
                int quarter = int.Parse(quarters.Text);
            int year = int.Parse(yeaarss.Text);
            try
            {


                using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                {
                    string commandText = $"SELECT c.Artist AS 'Артист',    c.Data AS 'Дата концерта',   c.Place AS 'Место проведения',  COUNT(t.NumTicket) AS 'Количество проданных билетов' FROM consert c JOIN ticket t ON c.idConsert = t.Consert_idConsert  WHERE YEAR(c.Data) = {year} and QUARTER(c.Data) = {quarter}    GROUP BY YEAR(c.Data), MONTH(c.Data), c.idConsert ORDER BY c.Data;";
                    connection.Open();

                    using (MySqlCommand mySqlCommand = new MySqlCommand(commandText, connection))
                    {
                        DataSet data = new DataSet();
                        MySqlDataAdapter adapter = new MySqlDataAdapter(mySqlCommand);
                        adapter.Fill(data, "otchet");
                        datagrid2.ItemsSource = data.Tables["otchet"].DefaultView;


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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void otchetdatapostC_Click(object sender, RoutedEventArgs e)
        {
            if (this.comyearC.Text != "" && this.monthcomnC.Text != "")
            {
                int year = int.Parse(comyearC.Text);
                 int month = int.Parse(monthcomnC.Text);

            try
            {


                using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                {
                    string commandText = $"SELECT    YEAR(c.Data) AS Год,   MONTH(c.Data) AS Месяц,   COUNT(t.NumTicket) AS Количество_билетов, SUM(toca.Cost) AS Общая_прибыль FROM ticket t JOIN consert c ON t.Consert_idConsert = c.idConsert JOIN consertticket toca ON t.idTypeOfTic = toca.idType  and  toca.IdConsert = c.idConsert WHERE YEAR(c.Data) = {year}    AND MONTH(c.Data) = {month}  GROUP BY YEAR(c.Data), MONTH(c.Data);";
                    connection.Open();

                    using (MySqlCommand mySqlCommand = new MySqlCommand(commandText, connection))
                    {
                        DataSet data = new DataSet();
                        MySqlDataAdapter adapter = new MySqlDataAdapter(mySqlCommand);
                        adapter.Fill(data, "otchet");
                        datagrid3.ItemsSource = data.Tables["otchet"].DefaultView;


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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void otchetyearposC_Click(object sender, RoutedEventArgs e)
        {
            if (this.yeaarsC.Text != "")
            {
                int year = int.Parse(yeaarsC.Text);

                try
                {


                    using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                    {
                        string commandText = $"SELECT    YEAR(c.Data) AS Год,    COUNT(t.NumTicket) AS Количество_билетов,    SUM(toca.Cost) AS Общая_прибыль FROM ticket t JOIN consert c ON t.Consert_idConsert = c.idConsert JOIN consertticket toca ON t.idTypeOfTic = toca.idType and  toca.IdConsert = c.idConsert  WHERE YEAR(c.Data) = {year} GROUP BY YEAR(c.Data);";
                        connection.Open();

                        using (MySqlCommand mySqlCommand = new MySqlCommand(commandText, connection))
                        {
                            DataSet data = new DataSet();
                            MySqlDataAdapter adapter = new MySqlDataAdapter(mySqlCommand);
                            adapter.Fill(data, "otchet");
                            datagrid3.ItemsSource = data.Tables["otchet"].DefaultView;


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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
            }

        private void otchetquarterC_Click(object sender, RoutedEventArgs e)
        {
            if (this.quartersC.Text != "" && this.yeaarssC.Text != "")
            {
                int quarter = int.Parse(quartersC.Text);
                int year = int.Parse(yeaarssC.Text);
                try
                {


                    using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                    {
                        string commandText = $"SELECT    YEAR(c.Data) AS Год,   QUARTER(c.Data) AS Квартал,   COUNT(t.NumTicket) AS Количество_билетов,  SUM(toca.Cost) AS Общая_прибыль FROM ticket t JOIN consert c ON t.Consert_idConsert = c.idConsert JOIN consertticket toca ON t.idTypeOfTic = toca.idType and  toca.IdConsert = c.idConsert WHERE YEAR(c.Data) = {year}    AND QUARTER(c.Data) = {quarter}  GROUP BY YEAR(c.Data), QUARTER(c.Data);";
                        connection.Open();

                        using (MySqlCommand mySqlCommand = new MySqlCommand(commandText, connection))
                        {
                            DataSet data = new DataSet();
                            MySqlDataAdapter adapter = new MySqlDataAdapter(mySqlCommand);
                            adapter.Fill(data, "otchet");
                            datagrid3.ItemsSource = data.Tables["otchet"].DefaultView;


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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
            }

        

        private void monthCost_Click(object sender, RoutedEventArgs e)
        {
            yearsC.Visibility = Visibility.Hidden;
            monthesC.Visibility = Visibility.Visible;
            quarterC.Visibility = Visibility.Hidden;
        }

        private void quarterCost_Click(object sender, RoutedEventArgs e)
        {
            yearsC.Visibility = Visibility.Hidden;
            monthesC.Visibility = Visibility.Hidden;
            quarterC.Visibility = Visibility.Visible;
        }

        private void yearCost_Click(object sender, RoutedEventArgs e)
        {
            yearsC.Visibility = Visibility.Visible;
            monthesC.Visibility = Visibility.Hidden;
            quarterC.Visibility = Visibility.Hidden;
        }

        private void OtchetExcel1(object sender, RoutedEventArgs e)
        {
            Excel.Application excelApp = null;
   
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx|Excel 97-2003|*.xls",
                    DefaultExt = "xlsx",
                    FileName = $"Отчет по прибыли_{DateTime.Now:yyyy-MM-dd_HH-mm}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    excelApp = new Excel.Application();
                    excelApp.Visible = false; 
                    excelApp.DisplayAlerts = false;

                    workbook = excelApp.Workbooks.Add();
                    worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                    excelApp = new Excel.Application();

                  
                    for (int i = 0; i < datagrid3.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1] = datagrid3.Columns[i].Header?.ToString() ?? "";
                        worksheet.Cells[1, i + 1].Font.Bold = true;
                    }

                   
                    int row = 2;
                    foreach (var item in datagrid3.Items)
                    {
                        if (item != null)
                        {
                            for (int col = 0; col < datagrid3.Columns.Count; col++)
                            {
                                var column = datagrid3.Columns[col];
                                object value = null;

                                
                                var cellContent = column.GetCellContent(item);
                                if (cellContent is TextBlock textBlock)
                                {
                                    value = textBlock.Text;
                                }
                                else
                                {
                                }

                                worksheet.Cells[row, col + 1] = value?.ToString() ?? "";
                            }
                            row++;
                        }
                    }
                }

                worksheet.Columns.AutoFit();

               
                workbook.SaveAs(saveFileDialog.FileName);
                workbook.Close();
                excelApp.Quit();

               
                if (MessageBox.Show("Экспорт завершен. Открыть файл?", "Успех",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start(saveFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
            finally
            {
                
               
                if (excelApp != null) excelApp.Quit();
            }
        }

        private void OtchetExcel2(object sender, RoutedEventArgs e)
        {
            Excel.Application excelApp = null;

            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx|Excel 97-2003|*.xls",
                    DefaultExt = "xlsx",
                    FileName = $"Отчет по посещениям_{DateTime.Now:yyyy-MM-dd_HH-mm}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    excelApp = new Excel.Application();
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;

                    workbook = excelApp.Workbooks.Add();
                    worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                    excelApp = new Excel.Application();


                    for (int i = 0; i < datagrid2.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1] = datagrid2.Columns[i].Header?.ToString() ?? "";
                        worksheet.Cells[1, i + 1].Font.Bold = true;
                    }


                    int row = 2;
                    foreach (var item in datagrid2.Items)
                    {
                        if (item != null)
                        {
                            for (int col = 0; col < datagrid2.Columns.Count; col++)
                            {
                                var column = datagrid2.Columns[col];
                                object value = null;


                                var cellContent = column.GetCellContent(item);
                                if (cellContent is TextBlock textBlock)
                                {
                                    value = textBlock.Text;
                                }
                                else
                                {
                                }

                                worksheet.Cells[row, col + 1] = value?.ToString() ?? "";
                            }
                            row++;
                        }
                    }
                }

                worksheet.Columns.AutoFit();


                workbook.SaveAs(saveFileDialog.FileName);
                workbook.Close();
                excelApp.Quit();


                if (MessageBox.Show("Экспорт завершен. Открыть файл?", "Успех",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start(saveFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
            finally
            {


                if (excelApp != null) excelApp.Quit();
            }
        }
        public void ConserTic()
        {
            try
            {
               
               

                using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                {
                    string commandText = "SELECT * FROM consertticket;";
                    connection.Open();

                    using (MySqlCommand mySqlCommand = new MySqlCommand(commandText, connection))
                    {
                        DataSet data = new DataSet();
                        MySqlDataAdapter adapter = new MySqlDataAdapter(mySqlCommand);

                        adapter.Fill(data, "consertticket");
                        datagrid.ItemsSource = data.Tables["consertticket"].DefaultView;
                        datagrid.Columns[0].Header = "#";
                        datagrid.Columns[1].Header = "#Концерта";
                        datagrid.Columns[2].Header = "#ТипБилета";
                        datagrid.Columns[3].Header = "Количество";
                        datagrid.Columns[4].Header = "Цена";
                        
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading consertticket: {ex.Message}", "Error",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        

        private void lolvo_Click(object sender, RoutedEventArgs e)
        {
            change = "conserttic_clicked";
            ChangeCanva();
            ConserTic();
        }

        private void addticcobc_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.lol3.Text != "" && this.lol4.Text != "" && this.lol5.Text != "" && this.lol6.Text != "")
            {
                int idCons = int.Parse(this.lol3.Text);
                int idTypT = int.Parse(this.lol4.Text);
                int kolvo = int.Parse(this.lol5.Text);
                int cost = int.Parse(this.lol6.Text);

                string query = $"insert into  consertticket  (IdConsert,IdType,Cost, Quanity) Values ('{idCons}','{idTypT}','{kolvo}','{cost}');";
                try
                {

                    conn.Open();
                    MySqlCommand command = new MySqlCommand();
                    command.Connection = conn;
                    command.CommandText = query;
                    int number = command.ExecuteNonQuery();

                    MessageBox.Show($"Запись  {number} удалена  успешно!");
                    DeleteTextBox();
                    ConserTic();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void upticconc_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.lol3.Text != "" && this.lol4.Text != "" && this.lol5.Text != "" && this.lol6.Text != "" && this.lol2.Text != "")
            {
                int idct = int.Parse(this.lol2.Text);
                int idCons = int.Parse(this.lol3.Text);
                int idTypT = int.Parse(this.lol4.Text);
                int kolvo = int.Parse(this.lol5.Text);
                int cost = int.Parse(this.lol6.Text);

                string query = $"update  consertticket set IdConsert ='{idCons}', IdType ='{idTypT}', Quanity ='{kolvo}', Cost ='{cost}' where idConsertTicket = {idct};";
                try
                {

                    conn.Open();
                    MySqlCommand command = new MySqlCommand();
                    command.Connection = conn;
                    command.CommandText = query;
                    int number = command.ExecuteNonQuery();

                    MessageBox.Show($"Запись  {number} изменена успешно!");
                    DeleteTextBox();
                    ConserTic();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void deleteticcocn_Click(object sender, RoutedEventArgs e)
        {
            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (this.lol2.Text != "")
            {
                int idct = int.Parse(this.lol2.Text);

                string query = $"delete from  consertticket  where idConsertTicket = {idct};";
                try
                {

                    conn.Open();
                    MySqlCommand command = new MySqlCommand();
                    command.Connection = conn;
                    command.CommandText = query;
                    int number = command.ExecuteNonQuery();

                    MessageBox.Show($"Запись  {number} удалена успешно!");
                    DeleteTextBox();
                    ConserTic();
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
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            } 
            }

        private void Window_Loaded_1(object sender, RoutedEventArgs e)
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
                string commandText2 = $"SELECT      MONTH(c.Data) AS month, SUM(toca.Cost) AS totalcost FROM ticket t JOIN consert c ON t.Consert_idConsert = c.idConsert JOIN consertticket toca ON t.idTypeOfTic = toca.idType  WHERE YEAR(c.Data) = 2025   or YEAR(c.Data) = 2026 GROUP BY YEAR(c.Data), MONTH(c.Data);";
                connection.Open();

                MySqlDataAdapter dataAdapter2 = new MySqlDataAdapter(commandText2, connection);
                System.Data.DataTable dt3 = new System.Data.DataTable();
                dataAdapter.Fill(dt3);





                all3 = new double[dt3.Rows.Count];

                for (int i = 0; i < dt3.Rows.Count; i++)
                {
                    all3[i] = Convert.ToDouble(dt3.Rows[i]["TotalTickets"]);
                }
                var line2 = new InteractiveDataDisplay.WPF.LineGraph
                {
                    Stroke = new SolidColorBrush(Colors.Gray),
                    
                    StrokeThickness = 2
                };
                connection.Close();
                int pointCount2 = all3.Length; // количество точек равно количеству месяцев
                double[] xs2 = Consecutive(pointCount2); // создает массив 
                double[] ys12 = all3; // массив с проданными билетами
                line2.Plot(xs2, ys12);

                graph2.Children.Clear();
                graph2.Children.Add(line2);

            }
           
        }

        private void biletpdf_Click(object sender, RoutedEventArgs e)
        {
            Ticketsret();
            change2 = "tic";
            ChangeCanva();

        }

        private void tic_Click(object sender, RoutedEventArgs e)
        {
            concer();
            try
            {
                // Загрузка и рендеринг шаблона
                using (var stream = new FileStream("Templates\\ticket.lqd", FileMode.Open))
                using (var reader = new StreamReader(stream))
                {
                    var templateString = reader.ReadToEnd();
                    var template = DotLiquid.Template.Parse(templateString);
                    var ticketContext = CreateTicketContext();
                    var docString = template.Render(ticketContext);

                    // Отображение в Preview
                    DocViewer.Document = (FlowDocument)XamlReader.Parse(docString);
                }

                // Сохранение в PDF
                GeneratePdf($"concert_ticket{num1.ToString()}.pdf");

                MessageBox.Show("Билет успешно создан и сохранен в PDF!", "Успех",
                               MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании билета: {ex.Message}", "Ошибка",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void concer()
        {


            MySqlConnection conn = new MySqlConnection(ConnectionString);
            if (ticclick.Text != "") { 
            int nnn = int.Parse(ticclick.Text);
            string query = $"select tc.NumTicket, c.Artist, c.Data, c.Time, c.Place, c.Adress , t.Type, ct.cost from  consert  c , typeoftic t, consertticket ct, ticket tc where tc.NumTicket = {nnn} and t.idtypeoftic = tc.idtypeoftic and tc.consert_idconsert = c.idConsert  and ct.idConsert = c.Idconsert and ct.idType = t.idTypeoftic ;";


            conn.Open();

            MySqlDataAdapter dataAdapter = new MySqlDataAdapter(query, conn);
            System.Data.DataTable dt2 = new System.Data.DataTable();
            dataAdapter.Fill(dt2);


            if (dt2.Rows.Count > 0)
            {
                num1 = (int)dt2.Rows[0]["Numticket"];
                artist1 = (string)dt2.Rows[0]["Artist"];
                data1 = (DateTime)dt2.Rows[0]["Data"];
                time1 = (TimeSpan)(dt2.Rows[0]["Time"]); ;
                place1 = (string)dt2.Rows[0]["Place"];
                adress1 = (string)dt2.Rows[0]["Adress"];
                type1 = (string)dt2.Rows[0]["Type"];
                cost1 = (decimal)dt2.Rows[0]["cost"];


            } }
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }
        private DotLiquid.Hash CreateTicketContext()
        {

            var context = new
            {
                EventName = artist1,
                EventDate = data1 + " : " + time1,
                Venue = place1,
                Address = adress1,

                TicketType = "Входной билет",
                TicketPrice = cost1,
                TicketNumber = num1,

                Barcode = "1234567890128",


                AdditionalInfo = new List<dynamic>
        {
            new { Label = "Тип билета", Value = type1 },
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
        private void Ticketsret()
        {

            string query = $"select NumTicket from ticket;";

            System.Data.DataTable dataTable = new System.Data.DataTable();

            using (MySqlConnection connection = new MySqlConnection(ConnectionString))
            {
                MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                adapter.Fill(dataTable);
            }

            ticclick.ItemsSource = dataTable.DefaultView;
            ticclick.DisplayMemberPath = "NumTicket";

        } 
        public void ChangeCanva()
        {
            switch (change)
            {
                case "consert_clicked":
                    
                    
                        consertsb.Visibility = Visibility.Visible;
                        ticketsb.Visibility = Visibility.Hidden;
                        sotrudB.Visibility = Visibility.Hidden;
                        doljd.Visibility = Visibility.Hidden;
                        typeoftic.Visibility = Visibility.Hidden;
                    
                    lol.Visibility = Visibility.Hidden;

                    break;
                case "ticket_clicked":
                    
                        consertsb.Visibility = Visibility.Hidden;
                        ticketsb.Visibility = Visibility.Visible;
                        sotrudB.Visibility = Visibility.Hidden;
                        doljd.Visibility = Visibility.Hidden;
                        typeoftic.Visibility = Visibility.Hidden;
                    lol.Visibility = Visibility.Hidden;

                    break;
                case "typeoftic_clicked":
                    
                        consertsb.Visibility = Visibility.Hidden;
                        ticketsb.Visibility = Visibility.Hidden;
                        sotrudB.Visibility = Visibility.Hidden;
                        doljd.Visibility = Visibility.Hidden;
                        typeoftic.Visibility = Visibility.Visible;
                    lol.Visibility = Visibility.Hidden;


                    break;
                case "sotrudniki_clicked":
                    
                        consertsb.Visibility = Visibility.Hidden;
                        ticketsb.Visibility = Visibility.Hidden;
                        sotrudB.Visibility = Visibility.Visible;
                        doljd.Visibility = Visibility.Hidden;
                        typeoftic.Visibility = Visibility.Hidden;
                    lol.Visibility = Visibility.Hidden;
                    break;

                case "dolj_clicked":
                    consertsb.Visibility = Visibility.Hidden;
                    ticketsb.Visibility = Visibility.Hidden;
                    sotrudB.Visibility = Visibility.Hidden;
                    doljd.Visibility = Visibility.Visible;
                    typeoftic.Visibility = Visibility.Hidden;
                    lol.Visibility = Visibility.Hidden;
                    break;
                case "conserttic_clicked":
                    consertsb.Visibility = Visibility.Hidden;
                    ticketsb.Visibility = Visibility.Hidden;
                    sotrudB.Visibility = Visibility.Hidden;
                    doljd.Visibility = Visibility.Hidden;
                    typeoftic.Visibility = Visibility.Hidden;
                    lol.Visibility = Visibility.Visible;
                    break;
            }
            switch (change2)
            {
                case "canvas":
                    canvas5.Visibility = Visibility.Hidden;
                    canvas.Visibility = Visibility.Visible;
                    canvas2.Visibility = Visibility.Hidden;
                    canvas3.Visibility = Visibility.Hidden;
                    canvas4.Visibility = Visibility.Hidden;
                    break;
                case "main":
                    consertsb.Visibility = Visibility.Hidden;
                    ticketsb.Visibility = Visibility.Hidden;
                    sotrudB.Visibility = Visibility.Hidden;
                    doljd.Visibility = Visibility.Hidden;
                    typeoftic.Visibility = Visibility.Hidden;
                    canvas5.Visibility = Visibility.Hidden;
                    canvas.Visibility = Visibility.Hidden;
                    canvas2.Visibility = Visibility.Visible;
                    canvas3.Visibility = Visibility.Hidden;
                    canvas4.Visibility = Visibility.Hidden;
                    break ;
                case "poseh":
                    canvas3.Visibility = Visibility.Visible;
                    canvas2.Visibility = Visibility.Hidden;
                    canvas.Visibility = Visibility.Hidden;
                    canvas4.Visibility = Visibility.Hidden;
                    canvas5.Visibility = Visibility.Hidden;
                    break;
                case "prib":
                    canvas3.Visibility = Visibility.Hidden;
                    canvas2.Visibility = Visibility.Hidden;
                    canvas.Visibility = Visibility.Hidden;
                    canvas4.Visibility = Visibility.Visible;
                    canvas5.Visibility = Visibility.Hidden;
                    break;
                case "tic":
                    canvas3.Visibility = Visibility.Hidden;
                    canvas2.Visibility = Visibility.Hidden;
                    canvas.Visibility = Visibility.Hidden;
                    canvas4.Visibility = Visibility.Hidden;
                    canvas5.Visibility = Visibility.Visible;
                    break;

            }

            
        }
       
        public void showDocument()
        {
            try
            {

                string filePath = ".\\Руководство администратора.docx"; 


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

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            showDocument();
        }
    }
}
