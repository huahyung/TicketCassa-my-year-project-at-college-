using Microsoft.SqlServer.Server;
using MySql.Data;
using MySql.Data.MySqlClient;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
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
using System.Xml.Linq;
namespace TicketCassa0fConsertPlace
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        static string serverName = "localhost";
        static string userName = "root";
        static string dbname = "consert_cassa";
        static string port = "3306";
        static string password = "vika_17092007";
        public string ConnectionString = $"Server={serverName};Port={port};Database={dbname};User ID={userName};Password={password};SslMode=None;AllowPublicKeyRetrieval=True;";
        public Window1()
        {
            InitializeComponent();
        }

        

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string log = loginn.Text;
            string pw = passwordd.Password;

          
           
           

            string commandText2 = $"SELECT idDoljnost , login, Familia, Name, Otchestvo FROM sotrudniki WHERE login='{log}' AND CAST(AES_DECRYPT(password, 'mysecretkey')AS CHAR) = '{pw}';";

            MySqlConnection connection = new MySqlConnection(ConnectionString);
                
                    connection.Open();

                    MySqlDataAdapter dataAdapter = new MySqlDataAdapter(commandText2, connection);
                    DataTable dt = new DataTable();
                    dataAdapter.Fill(dt);
                    

                    if (dt.Rows.Count > 0 )
                    {
                        int id = (int)dt.Rows[0]["idDoljnost"];
                        switch (id)
                        {
                            case 1:
                            // Успешный вход
                            Window2 window = new Window2(id);
                            window.nameofus.Content = dt.Rows[0]["Familia"].ToString() + " " + dt.Rows[0]["Name"].ToString() + " " + dt.Rows[0]["Otchestvo"].ToString();
                            window.welcome.Content = "Приветствую Вас, \n" + dt.Rows[0]["Name"].ToString() + " !\n\n" + "Чудесно выглядите !";
                            
                            window.ShowDialog();
                            this.Hide();
                            break;
                            case 2:
                        
                                // Успешный вход
                                Window4 cassier = new Window4(id);
                                cassier.nameofus.Content = dt.Rows[0]["Familia"].ToString() + " " + dt.Rows[0]["Name"].ToString() + " " + dt.Rows[0]["Otchestvo"].ToString();
                                cassier.welcome.Content = "Приветствую Вас, \n" + dt.Rows[0]["Name"].ToString() + " !\n\n" + "Чудесно выглядите !";
                                cassier.ShowDialog();
                                this.Hide();
                            break;
                        }
                           
                    }
                    else
                    {
                        MessageBox.Show("Неверный логин или пароль.");
                    }
                
            
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
           Window3 window = new Window3();
            this.Close();
            window.ShowDialog();
           
        }

        private void reg_MouseEnter(object sender, MouseEventArgs e)
        {
            reg.Foreground = Brushes.Blue;
        }

        private void reg_MouseLeave(object sender, MouseEventArgs e)
        {
            reg.Foreground = Brushes.Black;
        }
    }
}
