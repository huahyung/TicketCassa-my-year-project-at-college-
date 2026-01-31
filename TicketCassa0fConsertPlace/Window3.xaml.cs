using MySql.Data.MySqlClient;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
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
using System.Windows.Shapes;
using WpfMessageBoxLibrary;

namespace TicketCassa0fConsertPlace
{
    /// <summary>
    /// Логика взаимодействия для Window3.xaml
    /// </summary>
    public partial class Window3 : Window
    {
        public Window3()
        {
            InitializeComponent();
        }

        static string serverName = "localhost";
        static string userName = "root";
        static string dbname = "consert_cassa";
        static string port = "3306";
        static string password = "vika_17092007";
        public string ConnectionString = $"Server={serverName};Port={port};Database={dbname};User ID={userName};Password={password};SslMode=None;AllowPublicKeyRetrieval=True;";


        private void vhod_Click(object sender, RoutedEventArgs e)
        {
            Window1 window = new Window1();
            this.Close();
            window.ShowDialog();
            
        }

        private void vhod_MouseEnter(object sender, MouseEventArgs e)
        {
            vhod.Foreground = Brushes.Blue;
        }

        private void vhod_MouseLeave(object sender, MouseEventArgs e)
        {
            vhod.Foreground = Brushes.Black;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string log = loginnn.Text;
            string pw = passworddd.Password;
            string famw = fam.Text;
            string namee = name.Text;
            string otchc = otch.Text;
            int dolj = 0;
            if (cas.IsChecked == true)
            {
                dolj = 2;
                var msgProperties = new WpfMessageBoxProperties()
                {
                    Button = MessageBoxButton.OKCancel,
                    ButtonOkText = "Подтвердить",
                    Image = MessageBoxImage.Exclamation,
                    Header = "Требуется подтверждение ",
                    IsTextBoxVisible = true,
                    Text = "Пожалуйста подождите пока  админ введёт пароль",
                    Title = "Требуется подтверждение",
                };

                string query = $"insert into sotrudniki(login,password,familia,name,otchestvo,idDoljnost) values ('{log}',AES_ENCRYPT('{pw}','mysecretkey'),'{famw}','{namee}','{otchc}',{dolj});";

                
                
                MessageBoxResult result = WpfMessageBox.Show(this, ref msgProperties);
                string passwordd  = msgProperties.TextBoxText;

                string queryAccept = $"select CAST(AES_DECRYPT(Password,'mysecretkey')as char) from sotrudniki where idDoljnost = 1;";
                MySqlConnection conn = new MySqlConnection(ConnectionString);
                conn.Open();
                var command = new MySqlCommand(queryAccept, conn);
                var res = command.ExecuteReader();
                bool passwordCorrect = false;

                // Проверяем все пароли администраторов
                while (res.Read())
                {
                    string storedPassword = res["CAST(AES_DECRYPT(Password,'mysecretkey')as char)"].ToString();
                    if (storedPassword == passwordd)
                    {
                        passwordCorrect = true;
                        break;
                    }
                }
                conn.Close();
                if (passwordCorrect == true)
                {
                    try
                    {
                        conn.Open();
                        MySqlCommand commands = new MySqlCommand(query, conn);
                        int number = commands.ExecuteNonQuery();

                        MessageBox.Show($"Успешно зарегистровались!");
                        conn.Close();
                        conn.Close();
                        loginnn.Clear();
                        passworddd.Clear();
                        fam.Clear();
                        name.Clear();
                        otch.Clear();
                        cas.IsChecked = false;
                        Window1 window = new Window1();
                        window.Show();
                        this.Close();
                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка в регистрации: {ex.Message}");
                        conn.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Неправильный пароль");
                    conn.Close();
                }
            }
            else if (admin.IsChecked == true)
            {
                dolj = 1;
                var msgProperties = new WpfMessageBoxProperties()
                {
                    Button = MessageBoxButton.OKCancel,
                    ButtonOkText = "Подтвердить",
                    Image = MessageBoxImage.Exclamation,
                    Header = "Требуется подтверждение ",
                    IsTextBoxVisible = true,
                    Text = "Пожалуйста подождите пока главный админ введёт пароль",
                    Title = "Требуется подтверждение",
                };

                string query = $"insert into sotrudniki(login,password,familia,name,otchestvo,idDoljnost) values ('{log}',AES_ENCRYPT('{pw}','mysecretkey'),'{famw}','{namee}','{otchc}',{dolj});";



                MessageBoxResult result = WpfMessageBox.Show(this, ref msgProperties);
                string passwordd = msgProperties.TextBoxText;

                string queryAccept = $"select CAST(AES_DECRYPT(Password,'mysecretkey')as char) from sotrudniki where idDoljnost = 1 and id_Sotrudniki = 1;";
                MySqlConnection conn = new MySqlConnection(ConnectionString);
                conn.Open();
                var command = new MySqlCommand(queryAccept, conn);
                var res = command.ExecuteScalar();

                if (res.ToString() == passwordd)
                {
                    try
                    {

                        MySqlCommand commands = new MySqlCommand(query, conn);
                        int number = commands.ExecuteNonQuery();

                        MessageBox.Show($"Успешно зарегистровались!");
                        conn.Close();
                        loginnn.Clear();
                        passworddd.Clear();
                        fam.Clear();
                        name.Clear();
                        otch.Clear();
                        admin.IsChecked = false;
                        Window1 window = new Window1();
                        window.Show();
                        this.Close();
                        

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка в регистрации: {ex.Message}");
                        conn.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Неправильный пароль");
                    conn.Close();
                }

            }




            }

        }
}
