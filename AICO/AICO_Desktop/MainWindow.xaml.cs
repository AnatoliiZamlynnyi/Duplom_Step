using AICO_CL.Entity;
using AICO_CL.Models;
using AICO_Desktop.View;
using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Media;

namespace AICO_Desktop
{
    public partial class MainWindow : Window
    {
        Computer myComp;
        EfContext context;
        public MainWindow()
        {
            var startLogo = new StartLogo();
            startLogo.Show();
            InitializeComponent();
            context = new EfContext();
            context.Database.CreateIfNotExists();
            Employe user = new Employe();
            user = context.Employes.FirstOrDefault(x => x.Name == "admin");
            if (user == null)
            {
                Department dep = new Department
                {
                    Name = "IT"
                };
                context.Departments.Add(dep);
                context.SaveChanges();
                Employe newUser = new Employe
                {
                    Name = "Admin",
                    Password = CodingGetHash("123456"),
                    Work = "IT-Inginer",
                    Phone = "+380671234567",
                    DepartmentID = dep.ID
                };
                context.Employes.Add(newUser);
                context.SaveChanges();
            }
            myComp = new Computer();
            myComp.InfoSystem();
            userPC.Content = "Поточне ПК. Користувач: " + Environment.UserName;
            lb0.Content = myComp.UserNamePC;
            lb1.Content = myComp.NamePC;
            lb2.Content = myComp.OSVersion;
            lb3.Content = myComp.BitOperating;
            lb4.Content = myComp.Motherboard;
            lb5.Content = myComp.CPUpc;
            lb6.Content = myComp.RAMpc;
            lb7.Content = myComp.HDDpc;
            lb8.Content = myComp.Video;
            userPCDB.Content = "Обладнання закріплене за ПК.";
            Computer newComp = new Computer();
            newComp = context.Computers.FirstOrDefault(x => x.NamePC == myComp.NamePC);
            if (newComp != null)
            {
                _lb0.Content = newComp.UserNamePC;
                _lb0.Foreground = newComp.UserNamePC == myComp.UserNamePC ? Brushes.Black : Brushes.Red;
                _lb1.Content = newComp.NamePC;
                _lb2.Content = newComp.OSVersion;
                _lb2.Foreground = newComp.OSVersion == myComp.OSVersion ? Brushes.Black : Brushes.Red;
                _lb3.Content = newComp.BitOperating;
                _lb3.Foreground = newComp.BitOperating == myComp.BitOperating ? Brushes.Black : Brushes.Red;
                _lb4.Content = newComp.Motherboard;
                _lb4.Foreground = newComp.Motherboard == myComp.Motherboard ? Brushes.Black : Brushes.Red;
                _lb5.Content = newComp.CPUpc;
                _lb5.Foreground = newComp.CPUpc == myComp.CPUpc ? Brushes.Black : Brushes.Red;
                _lb6.Content = newComp.RAMpc;
                _lb6.Foreground = newComp.RAMpc == myComp.RAMpc ? Brushes.Black : Brushes.Red;
                _lb7.Content = newComp.HDDpc;
                _lb7.Foreground = newComp.HDDpc == myComp.HDDpc ? Brushes.Black : Brushes.Red;
                _lb8.Content = newComp.Video;
                _lb8.Foreground = newComp.Video == myComp.Video ? Brushes.Black : Brushes.Red;
            }
            startLogo.Close();
        }
        public static string CodingGetHash(string password)
        {
            using (var hash = SHA1.Create())
                return string.Concat(hash.ComputeHash(Encoding.UTF8.GetBytes(password)).Select(x => x.ToString("X2")));
        }

        private void Click_AdminLogin(object sender, RoutedEventArgs e)
        {
            var adminLogin = new AdminLogin();
            adminLogin.Show();
            this.Close();
        }
    }
}
