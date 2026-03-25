using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace FridgeLabReport
{
    /// <summary>
    /// Логика взаимодействия для ProgressWindow.xaml
    /// </summary>
    public partial class BusyWindow : Window
    {
        public BusyWindow(string message)
        {
            InitializeComponent();
            TbMessage.Text = message;
        }

        public void SetMessage(string message)
        {
            TbMessage.Text = message;
        }
    }
}
