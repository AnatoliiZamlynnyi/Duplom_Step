using AICO_CL.Entity;
using AICO_CL.Models;
using AICO_Desktop.View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace AICO_Desktop
{
    /// <summary>
    /// Логика взаимодействия для AdminLogin.xaml
    /// </summary>
    public partial class AdminLogin : Window
    {
        EfContext context;
        Employe user;
        public AdminLogin()
        {
            InitializeComponent();
            context = new EfContext();
            login.Focus();
        }

        public static string CodingGetHash(string password)
        {
            using (var hash = SHA1.Create())
                return string.Concat(hash.ComputeHash(Encoding.UTF8.GetBytes(password)).Select(x => x.ToString("X2")));
        }

        private void Click_EnterAdmin(object sender, RoutedEventArgs e)
        {
            string log = login.Text.ToLower();
            string password = CodingGetHash(pass.Password);
            user = context.Employes.FirstOrDefault(x => x.Name.ToLower() == log && x.Password == password);
            if (user == null)
            {
                stan.Content = "Ім'я або пароль невірнi.";
                login.Clear();
                pass.Clear();
                login.Focus();
            }
            else
            {
                var config = new ManagementWindow(user);
                config.Show();
                this.Close();
            }
        }
    }
}