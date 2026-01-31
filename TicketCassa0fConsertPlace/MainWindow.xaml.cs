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
using System.Windows.Threading;

namespace TicketCassa0fConsertPlace
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
   

    public partial class MainWindow : Window
    {


        private DispatcherTimer timer;
        int startpoint = 0;
        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += FormLoad;
             
        }
        private void InitializeTimer()
        {
            timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = TimeSpan.FromMilliseconds(10);
            timer.Tick += Timer_Tick;
        }

        private void FormLoad(object sender, RoutedEventArgs e)
        {
            if (timer == null)
            {
                InitializeTimer();
            }

            timer.IsEnabled = false;
            timer.Start();

        }
        private void Timer_Tick(object sender, EventArgs e)
        {
            startpoint += 2;
            barbar.Value = startpoint;
            procentssss.Content = barbar.Value.ToString() + "%";
            
            if (barbar.Value == 100)
            {
               timer.Stop();
                Window1 avtoriz = new Window1();
               
                this.Close();
                avtoriz.Show();

            }
           
        }
    }
}
