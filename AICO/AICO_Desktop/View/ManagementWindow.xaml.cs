using AICO_CL.Entity;
using AICO_CL.Models;
using AICO_Desktop.View;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;
using System.Security.Cryptography;
using System.Diagnostics;
using System.Threading;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace AICO_Desktop
{
    public partial class ManagementWindow : Window
    {
        Computer myComp;
        EfContext context;
        public Employe root { get; set; }
        public ManagementWindow(Employe user)
        {
            InitializeComponent();
            context = new EfContext();
            root = user;
            this.Title = root.Name + ", " + root.Work + ", " + root.Phone;
        }

        public static string CodingGetHash(string password)
        {
            using (var hash = SHA1.Create())
                return string.Concat(hash.ComputeHash(Encoding.UTF8.GetBytes(password)).Select(x => x.ToString("X2")));
        }

        //Вибір закладок головного вікна = Selector_OnSelect
        private void Selector_OnSelect(object sender, SelectionChangedEventArgs e)
        {
            if ((tcSample.SelectedItem as TabItem).Name == "one")
            {
                log.Content = "";
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
                Computer newComp = context.Computers.FirstOrDefault(x => x.NamePC == myComp.NamePC);
                Accounting userTmp = new Accounting();
                if (newComp != null)
                {
                    userTmp = context.Accountings.FirstOrDefault(x => x.ComputerID == newComp.ID);
                    if (userTmp != null)
                    {
                        userPCDB.Content = "Обладнання закріплене за ПК " + newComp.NamePC;
                    }
                    else
                    {
                        userPCDB.Content = "Обладнання закріплене за ПК. ";
                    }
                    if (newComp != null)
                    {
                        AddComp.IsEnabled = false;
                        EditComp.IsEnabled = false;
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
                        if (_lb2.Foreground == Brushes.Red || _lb3.Foreground == Brushes.Red
                            || _lb4.Foreground == Brushes.Red || _lb5.Foreground == Brushes.Red
                            || _lb6.Foreground == Brushes.Red || _lb7.Foreground == Brushes.Red || _lb8.Foreground == Brushes.Red)
                        {
                            EditComp.IsEnabled = true;
                        }
                    }
                }
                else
                {
                    AddComp.IsEnabled = true;
                    userPCDB.Content = "База даних ПК порожня.";
                }
            }
            if ((tcSample.SelectedItem as TabItem).Name == "two")
            {
                department.ItemsSource = context.Departments.ToList();
                if (department.SelectedItem == null)
                    employe.ItemsSource = context.Employes.ToList();
                else
                {
                    Department editdep = department.SelectedItem as Department;
                    employe.ItemsSource = context.Employes.Where(x => x.Departments.ID == editdep.ID).ToList();
                }
                if (context.Departments.Count() != 0)
                    departmentsName.ItemsSource = context.Departments.Select(x => x.Name).ToList();
                else
                    departmentsName.ItemsSource = null;
                device_ENUM.ItemsSource = context.Device_ENUMs.ToList();
                if (device_ENUM.SelectedItem == null)
                    device.ItemsSource = context.Devices.ToList();
                else
                {
                    Device_ENUM editDevEnum = device_ENUM.SelectedItem as Device_ENUM;
                    device.ItemsSource = context.Devices.Where(x => x.Devices_ENUM.ID == editDevEnum.ID).ToList();
                }
                if (context.Device_ENUMs.Count() != 0)
                    deviceENUM.ItemsSource = context.Device_ENUMs.Select(x => x.Name).ToList();
                else
                    deviceENUM.ItemsSource = null;
                ScrollViewer.SetVerticalScrollBarVisibility(this.device, ScrollBarVisibility.Visible);
            }
            if ((tcSample.SelectedItem as TabItem).Name == "three")
            {
                departmentA.ItemsSource = context.Departments.ToList();
                if (departmentA.SelectedItem == null)
                    employeA.ItemsSource = context.Employes.ToList();
                else
                {
                    Department editdep = departmentA.SelectedItem as Department;
                    employeA.ItemsSource = context.Employes.Where(x => x.Departments.ID == editdep.ID).ToList();
                }
                computerA.ItemsSource = context.Computers.ToList();
                device_ENUMA.ItemsSource = context.Device_ENUMs.ToList();
                if (device_ENUMA.SelectedItem == null)
                    deviceA.ItemsSource = context.Devices.ToList();
                else
                {
                    Device_ENUM editDevEnum = device_ENUMA.SelectedItem as Device_ENUM;
                    deviceA.ItemsSource = context.Devices.Where(x => x.Devices_ENUM.ID == editDevEnum.ID).ToList();
                }
                accounting.ItemsSource = context.Accountings.ToList();
            }
            if ((tcSample.SelectedItem as TabItem).Name == "four")
            {
                employeA.ItemsSource = context.Employes.ToList();
                computerA.ItemsSource = context.Computers.ToList();
                deviceA.ItemsSource = context.Devices.ToList();
                reportA.ItemsSource = context.Accountings.ToList();
                if (context.Departments.Count() != 0)
                    depSelect.ItemsSource = context.Departments.Select(x => x.Name).ToList();
                else
                    depSelect.ItemsSource = "";
                if (context.Employes.Count() != 0)
                    empSelect.ItemsSource = context.Employes.Select(x => x.Name).ToList();
                else
                    empSelect.ItemsSource = "";
                if (context.Device_ENUMs.Count() != 0)
                    devSelect.ItemsSource = context.Device_ENUMs.Select(x => x.Name).ToList();
                else
                    devSelect.ItemsSource = "";
            }
        }

        //Додавання та редагування ПК========Selector_OnSelect = one
        private void Click_NewPC(object sender, RoutedEventArgs e)
        {
            Computer compAdd = new Computer();
            if (context.Computers.Count() == 0)
            {
                compAdd.UserNamePC = lb0.Content.ToString();
                compAdd.NamePC = lb1.Content.ToString();
                compAdd.OSVersion = lb2.Content.ToString();
                compAdd.BitOperating = lb3.Content.ToString();
                compAdd.Motherboard = lb4.Content.ToString();
                compAdd.CPUpc = lb5.Content.ToString();
                compAdd.RAMpc = lb6.Content.ToString();
                compAdd.HDDpc = lb7.Content.ToString();
                compAdd.Video = lb8.Content.ToString();
                context.Computers.Add(compAdd);
                context.SaveChanges();
                log.Foreground = Brushes.Green;
                log.Content = "ПК додано!";
            }
            else if (context.Computers.Any(x => x.NamePC != lb1.Content.ToString()))
            {
                compAdd.UserNamePC = lb0.Content.ToString();
                compAdd.NamePC = lb1.Content.ToString();
                compAdd.OSVersion = lb2.Content.ToString();
                compAdd.BitOperating = lb3.Content.ToString();
                compAdd.Motherboard = lb4.Content.ToString();
                compAdd.CPUpc = lb5.Content.ToString();
                compAdd.RAMpc = lb6.Content.ToString();
                compAdd.HDDpc = lb7.Content.ToString();
                compAdd.Video = lb8.Content.ToString();
                context.Computers.Add(compAdd);
                context.SaveChanges();
                log.Foreground = Brushes.Green;
                log.Content = "ПК додано!";
            }
            else
            {
                log.Foreground = Brushes.Red;
                log.Content = "Такий ПК вже існує в БД!";
            }
        }

        private void Click_EditPC(object sender, RoutedEventArgs e)
        {
            Computer compEdit = context.Computers.FirstOrDefault(x => x.NamePC == _lb1.Content.ToString());
            using (var dbContext = new EfContext())
            {
                compEdit.OSVersion = lb2.Content.ToString();
                compEdit.BitOperating = lb3.Content.ToString();
                compEdit.Motherboard = lb4.Content.ToString();
                compEdit.CPUpc = lb5.Content.ToString();
                compEdit.RAMpc = lb6.Content.ToString();
                compEdit.HDDpc = lb7.Content.ToString();
                compEdit.Video = lb8.Content.ToString();
                dbContext.Computers.Attach(compEdit);
                dbContext.Entry(compEdit).State = System.Data.Entity.EntityState.Modified;
                dbContext.SaveChanges();
            }
            log.Foreground = Brushes.Green;
            log.Content = "Зміни внесено успішно!";
            EditComp.IsEnabled = false;
        }

        //Організація довідників========Selector_OnSelect = two
        private void MouseDuble_Dep(object sender, MouseButtonEventArgs e)
        {
            logDirectory.Content = "";
            try
            {
                Department editdep = department.SelectedItem as Department;
                depText.Text = editdep.Name;
            }
            catch
            {
                department.ItemsSource = context.Departments.ToList();
                depText.Clear();
            }
        }

        private void MouseUp_Dep(object sender, MouseButtonEventArgs e)
        {
            logDirectory.Content = "";
            try
            {
                if (department.SelectedItem != null)
                {
                    Department editdep = department.SelectedItem as Department;
                    employe.ItemsSource = context.Employes.Where(x => x.Departments.ID == editdep.ID).ToList();
                }
            }
            catch
            {
                department.ItemsSource = context.Departments.ToList();
                depText.Clear();
            }
        }

        private void MouseDuble_Employe(object sender, MouseButtonEventArgs e)
        {
            logDirectory.Content = "";
            if (employe.SelectedItem != null)
            {
                Employe editUser = employe.SelectedItem as Employe;
                name.Text = editUser.Name;
                work.Text = editUser.Work;
                phone.Text = editUser.Phone;
                departmentsName.SelectedItem = editUser.Departments.Name;
                if (editUser.Password == null)
                {
                    oK.IsEnabled = true;
                    fine.IsEnabled = false;
                }
                else
                {
                    oK.IsEnabled = false;
                    fine.IsEnabled = true;
                }
            }
            else
            {
                department.ItemsSource = null;
                department.ItemsSource = context.Departments.ToList();
                employe.ItemsSource = context.Employes.ToList();
                name.Clear();
                work.Clear();
                phone.Clear();
                departmentsName.SelectedItem = null;
            }
        }

        private void Click_AddDep(object sender, RoutedEventArgs e)
        {
            if (depText.Text != "")
            {
                var tmpDep = context.Departments.FirstOrDefault(x => x.Name == depText.Text);
                if (tmpDep == null)
                {
                    Department newDep = new Department();
                    newDep.Name = depText.Text;
                    context.Departments.Add(newDep);
                    context.SaveChanges();
                    department.ItemsSource = context.Departments.ToList();
                    depText.Clear();
                    logDirectory.Foreground = Brushes.Green;
                    logDirectory.Content = "Відділ додано";
                }
                else
                {
                    logDirectory.Foreground = Brushes.Red;
                    logDirectory.Content = "Такий відділ вже існує";
                }
            }
        }

        private void Click_EditDep(object sender, RoutedEventArgs e)
        {
            if (department.SelectedItem as Department != null)
            {
                var tmp = department.SelectedItem as Department;
                Department editDep = context.Departments.First(x => x.ID == tmp.ID);
                using (var dbContext = new EfContext())
                {
                    editDep.Name = depText.Text;
                    dbContext.Departments.Attach(editDep);
                    dbContext.Entry(editDep).State = System.Data.Entity.EntityState.Modified;
                    dbContext.SaveChanges();
                }
                department.ItemsSource = context.Departments.ToList();
                depText.Clear();
                logDirectory.Foreground = Brushes.Green;
                logDirectory.Content = "Зміни внесено успішно";
            }
        }

        private void Click_DeleteDep(object sender, RoutedEventArgs e)
        {
            if (department.SelectedItem as Department != null)
            {
                var tmp = department.SelectedItem as Department;
                Department delDep = context.Departments.First(x => x.ID == tmp.ID);
                context.Departments.Remove(delDep);
                context.SaveChanges();
                department.ItemsSource = context.Departments.ToList();
                depText.Clear();
                logDirectory.Foreground = Brushes.BlueViolet;
                logDirectory.Content = "Відділ видалено";
            }
        }

        private void Click_AddEmloye(object sender, RoutedEventArgs e)
        {
            if (name.Text != "")
            {
                var tmpUser = context.Employes.FirstOrDefault(x => x.Name == name.Text);
                if (tmpUser == null)
                {
                    Employe newUser = new Employe();
                    newUser.Name = name.Text;
                    newUser.Work = work.Text;
                    newUser.Phone = phone.Text;
                    var list = context.Departments.ToList();
                    foreach (var item in list)
                        if (item.Name == departmentsName.SelectedItem.ToString())
                            newUser.DepartmentID = item.ID;
                    context.Employes.Add(newUser);
                    context.SaveChanges();
                    employe.ItemsSource = context.Employes.ToList();
                    name.Clear();
                    work.Clear();
                    phone.Clear();
                    logDirectory.Foreground = Brushes.Green;
                    logDirectory.Content = "Працівника додано";
                }
                else
                {
                    logDirectory.Foreground = Brushes.Red;
                    logDirectory.Content = "Такий працівник вже існує";
                }
            }
        }

        private void Click_EditEmloye(object sender, RoutedEventArgs e)
        {
            if (employe.SelectedItem as Employe != null)
            {
                var tmp = employe.SelectedItem as Employe;
                Employe editUser = context.Employes.First(x => x.ID == tmp.ID);
                using (EfContext dbContext = new EfContext())
                {
                    dbContext.Employes.Attach(editUser);
                    editUser.Name = name.Text;
                    editUser.Work = work.Text;
                    editUser.Phone = phone.Text;
                    var list = dbContext.Departments.ToList();
                    foreach (var item in list)
                        if (item.Name == departmentsName.SelectedItem.ToString())
                            editUser.DepartmentID = item.ID;
                    dbContext.Entry(editUser).State = System.Data.Entity.EntityState.Modified;
                    dbContext.SaveChanges();
                }
                employe.ItemsSource = context.Employes.ToList();
                name.Clear();
                work.Clear();
                phone.Clear();
                logDirectory.Foreground = Brushes.Green;
                logDirectory.Content = "Зміни внесено успішно";
            }
        }

        private void Click_DeleteEmloye(object sender, RoutedEventArgs e)
        {
            if (employe.SelectedItem as Employe != null)
            {
                try
                {
                    var temp = employe.SelectedItem as Employe;
                    Employe delUser = context.Employes.First(x => x.ID == temp.ID);
                    context.Employes.Remove(delUser);
                    context.SaveChanges();
                    employe.ItemsSource = context.Employes.ToList();
                    name.Clear();
                    work.Clear();
                    phone.Clear();
                    departmentsName.SelectedItem = "";
                    logDirectory.Foreground = Brushes.BlueViolet;
                    logDirectory.Content = "Працівника видалено";
                }
                catch
                {
                    logDirectory.Foreground = Brushes.Red;
                    logDirectory.Content = "Видалити неможливо! Треба від'єднати в БД ";
                }
            }
        }

        private void AddAdmin_CheckedChanged(object sender, RoutedEventArgs e)
        {
            passL.IsEnabled = true;
            pass.IsEnabled = true;
            passCopyL.IsEnabled = true;
            passCopy.IsEnabled = true;
            oK.IsEnabled = true;
            fine.IsEnabled = true;
        }

        private void AddAdmin_UncheckedChanged(object sender, RoutedEventArgs e)
        {
            passL.IsEnabled = false;
            pass.IsEnabled = false;
            passCopyL.IsEnabled = false;
            passCopy.IsEnabled = false;
            oK.IsEnabled = false;
            fine.IsEnabled = false;
        }

        private void Click_AddAdministrator(object sender, RoutedEventArgs e)
        {
            if (addAdmin.IsChecked == true)
            {
                Employe newUser = employe.SelectedItem as Employe;
                bool end = false;
                do
                {
                    if (pass.Password == passCopy.Password)
                    {
                        using (EfContext dbContext = new EfContext())
                        {
                            dbContext.Employes.Attach(newUser);
                            newUser.Password = CodingGetHash(pass.Password);
                            dbContext.Entry(newUser).State = System.Data.Entity.EntityState.Modified;
                            dbContext.SaveChanges();
                            pass.Password = "";
                            passCopy.Password = "";
                            logDirectory.Foreground = Brushes.Green;
                            logDirectory.Content = "Права адміністратора надано";
                        }
                        break;
                    }
                    else
                    {
                        pass.Password = "";
                        passCopy.Password = "";
                        logDirectory.Foreground = Brushes.Red;
                        logDirectory.Content = "Пароль та підтвердження не співпадають";
                    }
                } while (end == true);
                employe.ItemsSource = context.Employes.ToList();
            }
        }

        private void Click_DelAdministrator(object sender, RoutedEventArgs e)
        {
            if (addAdmin.IsChecked == true)
            {
                Employe newUser = employe.SelectedItem as Employe;
                using (EfContext dbContext = new EfContext())
                {
                    dbContext.Employes.Attach(newUser);
                    newUser.Password = null;
                    dbContext.Entry(newUser).State = System.Data.Entity.EntityState.Modified;
                    dbContext.SaveChanges();
                }
                logDirectory.Foreground = Brushes.Green;
                logDirectory.Content = "Права Адміністратора скасовано!";
                employe.ItemsSource = context.Employes.ToList();
            }
        }

        private void MouseDuble_DeviceENUM(object sender, MouseButtonEventArgs e)
        {
            logDirectory.Content = "";
            try
            {
                Device_ENUM editDevEnum = device_ENUM.SelectedItem as Device_ENUM;
                devENUM_Text.Text = editDevEnum.Name;
            }
            catch
            {
                device_ENUM.ItemsSource = context.Device_ENUMs.ToList();
                devENUM_Text.Clear();
            }
        }

        private void MouseUp_DeviceENUM(object sender, MouseButtonEventArgs e)
        {
            logDirectory.Content = "";
            try
            {
                Device_ENUM editDevEnum = device_ENUM.SelectedItem as Device_ENUM;
                device.ItemsSource = context.Devices.Where(x => x.Devices_ENUM.ID == editDevEnum.ID).ToList();
            }
            catch
            {
                device_ENUM.ItemsSource = context.Device_ENUMs.ToList();
                devENUM_Text.Clear();
            }
        }

        private void MouseDuble_Device(object sender, MouseButtonEventArgs e)
        {
            logDirectory.Content = "";
            if (device.SelectedItem != null)
            {
                Device editDevice = device.SelectedItem as Device;
                model.Text = editDevice.Model;
                description1.Text = editDevice.Description_1;
                description2.Text = editDevice.Description_2;
                description3.Text = editDevice.Description_3;
                description4.Text = editDevice.Description_4;
                description5.Text = editDevice.Description_5;
                deviceENUM.SelectedItem = editDevice.Devices_ENUM.Name;
            }
            else
            {
                device_ENUM.ItemsSource = null;
                device_ENUM.ItemsSource = context.Device_ENUMs.ToList();
                device.ItemsSource = context.Devices.ToList();
                model.Clear();
                description1.Clear();
                description2.Clear();
                description3.Clear();
                description4.Clear();
                description5.Clear();
                deviceENUM.ItemsSource = null;
            }
        }

        private void Click_AddDevENUM(object sender, RoutedEventArgs e)
        {
            if (devENUM_Text.Text != "")
            {
                var tmpDevENUM = context.Device_ENUMs.FirstOrDefault(x => x.Name == devENUM_Text.Text);
                if (tmpDevENUM == null)
                {
                    Device_ENUM newDevEnum = new Device_ENUM();
                    newDevEnum.Name = devENUM_Text.Text;
                    context.Device_ENUMs.Add(newDevEnum);
                    context.SaveChanges();
                    device_ENUM.ItemsSource = context.Device_ENUMs.ToList();
                    devENUM_Text.Clear();
                    logDirectory.Foreground = Brushes.Green;
                    logDirectory.Content = "Тип пристрою додано";
                }
                else
                {
                    logDirectory.Foreground = Brushes.Red;
                    logDirectory.Content = "Такий тип вже існує";
                }
            }
        }

        private void Click_EditDevENUM(object sender, RoutedEventArgs e)
        {
            if (device_ENUM.SelectedItem as Device_ENUM != null)
            {
                var tmp = device_ENUM.SelectedItem as Device_ENUM;
                Device_ENUM editDevEnum = context.Device_ENUMs.First(x => x.ID == tmp.ID);
                using (var dbContext = new EfContext())
                {
                    editDevEnum.Name = devENUM_Text.Text;
                    dbContext.Device_ENUMs.Attach(editDevEnum);
                    dbContext.Entry(editDevEnum).State = System.Data.Entity.EntityState.Modified;
                    dbContext.SaveChanges();
                }
                device_ENUM.ItemsSource = context.Device_ENUMs.ToList();
                devENUM_Text.Clear();
                logDirectory.Foreground = Brushes.Green;
                logDirectory.Content = "Зміни внесено успішно";
            }
        }

        private void Click_DeleteDevENUM(object sender, RoutedEventArgs e)
        {
            if (device_ENUM.SelectedItem as Device_ENUM != null)
            {
                var tmp = device_ENUM.SelectedItem as Device_ENUM;
                Device_ENUM delDevEnum = context.Device_ENUMs.First(x => x.ID == tmp.ID);
                context.Device_ENUMs.Remove(delDevEnum);
                context.SaveChanges();
                device_ENUM.ItemsSource = context.Device_ENUMs.ToList();
                devENUM_Text.Clear();
                logDirectory.Foreground = Brushes.BlueViolet;
                logDirectory.Content = "Тип видалено";
            }
        }

        private void Click_AddDevice(object sender, RoutedEventArgs e)
        {
            if (model.Text != "")
            {
                var tmp = context.Devices.FirstOrDefault(x => x.Model == model.Text);
                if (tmp == null)
                {
                    Device newDevice = new Device();
                    newDevice.Model = model.Text;
                    newDevice.Description_1 = description1.Text;
                    newDevice.Description_2 = description2.Text;
                    newDevice.Description_3 = description3.Text;
                    newDevice.Description_4 = description4.Text;
                    newDevice.Description_5 = description5.Text;
                    var list = context.Device_ENUMs.ToList();
                    foreach (var item in list)
                        if (item.Name == deviceENUM.SelectedItem.ToString())
                            newDevice.Device_ENUM_ID = item.ID;
                    context.Devices.Add(newDevice);
                    context.SaveChanges();
                    device.ItemsSource = context.Devices.ToList();
                    model.Clear();
                    description1.Clear();
                    description2.Clear();
                    description3.Clear();
                    description4.Clear();
                    description5.Clear();
                    logDirectory.Foreground = Brushes.Green;
                    logDirectory.Content = "Пристрій додано";
                }
                else
                {
                    logDirectory.Foreground = Brushes.Red;
                    logDirectory.Content = "Такий пистрій вже існує";
                }
            }
        }

        private void Click_EditDevice(object sender, RoutedEventArgs e)
        {
            if (device.SelectedItem as Device != null)
            {
                var tmp = device.SelectedItem as Device;
                Device editDevice = context.Devices.First(x => x.ID == tmp.ID);
                using (EfContext dbContext = new EfContext())
                {
                    dbContext.Devices.Attach(editDevice);
                    editDevice.Model = model.Text;
                    editDevice.Description_1 = description1.Text;
                    editDevice.Description_2 = description2.Text;
                    editDevice.Description_3 = description3.Text;
                    editDevice.Description_4 = description4.Text;
                    editDevice.Description_5 = description5.Text;
                    var list = dbContext.Device_ENUMs.ToList();
                    foreach (var item in list)
                        if (item.Name == deviceENUM.SelectedItem.ToString())
                            editDevice.Device_ENUM_ID = item.ID;
                    dbContext.Entry(editDevice).State = System.Data.Entity.EntityState.Modified;
                    dbContext.SaveChanges();
                }
                device.ItemsSource = context.Devices.ToList();
                model.Clear();
                description1.Clear();
                description2.Clear();
                description3.Clear();
                description4.Clear();
                description5.Clear();
                logDirectory.Foreground = Brushes.Green;
                logDirectory.Content = "Зміни внесено успішно";
            }
        }

        private void Click_DeleteDevice(object sender, RoutedEventArgs e)
        {
            if (device.SelectedItem as Device != null)
            {
                try
                {
                    var tmp = device.SelectedItem as Device;
                    Device delDevice = context.Devices.First(x => x.ID == tmp.ID);
                    context.Devices.Remove(delDevice);
                    context.SaveChanges();
                    device.ItemsSource = context.Devices.ToList();
                    model.Clear();
                    description1.Clear();
                    description2.Clear();
                    description3.Clear();
                    description4.Clear();
                    description5.Clear();
                    deviceENUM.ItemsSource = "";
                    logDirectory.Foreground = Brushes.BlueViolet;
                    logDirectory.Content = "Пристрій видалено";
                }
                catch
                {
                    logDirectory.Foreground = Brushes.Red;
                    logDirectory.Content = "Видалити неможливо! Треба від'єднати в БД ";
                }
            }
        }

        //Облік техніки========Selector_OnSelect = three
        private void MouseUp_DepartmentA(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Department editdep = departmentA.SelectedItem as Department;
                employeA.ItemsSource = context.Employes.Where(x => x.Departments.ID == editdep.ID).ToList();
            }
            catch
            {
                departmentA.ItemsSource = context.Departments.ToList();
            }
        }

        private void MouseUp_User(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Employe tmpUser = employeA.SelectedItem as Employe;
                user.Text = tmpUser.Name;
            }
            catch
            {
                user.Clear();
            }
        }

        private void MouseUp_Comp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Computer tmpComp = computerA.SelectedItem as Computer;
                comp.Text = tmpComp.NamePC;
            }
            catch
            {
                comp.Clear();
            }
        }

        private void MouseUp_DeviceENUMA(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Device_ENUM editDevEnum = device_ENUMA.SelectedItem as Device_ENUM;
                deviceA.ItemsSource = context.Devices.Where(x => x.Devices_ENUM.ID == editDevEnum.ID).ToList();
            }
            catch
            {
                device_ENUMA.ItemsSource = context.Device_ENUMs.ToList();
            }
        }

        private void MouseUp_Dev(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Device tmpDev = deviceA.SelectedItem as Device;
                dev.Text = tmpDev.Model;
            }
            catch
            {
                dev.Clear();
            }
        }

        private void MouseUp_Accounting(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Accounting tmp = accounting.SelectedItem as Accounting;
                if (tmp.Employes.Name != null)
                    user.Text = tmp.Employes.Name;
                if (tmp.Computers.NamePC.ToString() != null)
                    comp.Text = tmp.Computers.NamePC;
                if (tmp.Devices.Model != null)
                    dev.Text = tmp.Devices.Model;
            }
            catch { }
        }

        private void Click_AddAccounting(object sender, RoutedEventArgs e)
        {
            if (user.Text != "")
            {
                var tmp = context.Accountings.FirstOrDefault(x => x.Computers.NamePC == comp.Text);
                if (tmp == null)
                {
                    Accounting account = new Accounting();
                    var userTMP = context.Employes.FirstOrDefault(x => x.Name == user.Text);
                    if (userTMP != null)
                        account.EmployeID = userTMP.ID;
                    else
                    {
                        logAccounting.Foreground = Brushes.Red;
                        logAccounting.Content = "Користувач неможе бути порожнім";
                    }
                    var compTMP = context.Computers.FirstOrDefault(x => x.NamePC == comp.Text);
                    if (compTMP != null)
                        account.ComputerID = compTMP.ID;
                    var devTMP = context.Devices.FirstOrDefault(x => x.Model == dev.Text);
                    if (devTMP != null)
                        account.DeviceID = devTMP.ID;
                    context.Accountings.Add(account);
                    context.SaveChanges();
                    accounting.ItemsSource = context.Accountings.ToList();
                    user.Clear();
                    comp.Clear();
                    dev.Clear();
                    logAccounting.Foreground = Brushes.Green;
                    logAccounting.Content = "Додано успішно";
                }
                else
                {
                    logAccounting.Foreground = Brushes.Red;
                    logAccounting.Content = "Такий ПК вже існує";
                }
            }
        }

        private void Click_EditAccounting(object sender, RoutedEventArgs e)
        {
            if (accounting.SelectedItem as Accounting != null)
            {
                var tmp = accounting.SelectedItem as Accounting;
                Accounting editAccount = context.Accountings.First(x => x.ID == tmp.ID);
                if (editAccount != null)
                {
                    using (EfContext dbContext = new EfContext())
                    {
                        try
                        {
                            dbContext.Accountings.Attach(editAccount);
                            var userTMP = dbContext.Employes.FirstOrDefault(x => x.Name == user.Text);
                            if (userTMP != null)
                                editAccount.EmployeID = userTMP.ID;
                            else
                            {
                                logAccounting.Foreground = Brushes.Red;
                                logAccounting.Content = "Користувач неможе бути порожнім";
                            }
                            var compTMP = dbContext.Computers.First(x => x.NamePC == comp.Text);
                            if (compTMP != null)
                                editAccount.ComputerID = compTMP.ID;
                            else
                                editAccount.ComputerID = null;
                            var devTMP = dbContext.Devices.First(x => x.Model == dev.Text);
                            if (devTMP != null)
                                editAccount.DeviceID = devTMP.ID;
                            else
                                editAccount.DeviceID = null;
                        }
                        catch { };
                        dbContext.Entry(editAccount).State = System.Data.Entity.EntityState.Modified;
                        dbContext.SaveChanges();
                    }
                    accounting.ItemsSource = context.Accountings.ToList();
                    user.Clear();
                    comp.Clear();
                    dev.Clear();
                    logAccounting.Foreground = Brushes.Green;
                    logAccounting.Content = "Змінено успішно";
                }
                else
                {
                    accounting.ItemsSource = context.Accountings.ToList();
                    user.Clear();
                    comp.Clear();
                    dev.Clear();
                    logAccounting.Foreground = Brushes.Red;
                    logAccounting.Content = "Такого запису неіснує";
                }
            }
        }

        private void Click_DeleteAccounting(object sender, RoutedEventArgs e)
        {
            if (accounting.SelectedItem as Accounting != null)
            {
                var tmp = accounting.SelectedItem as Accounting;
                Accounting delAccount = context.Accountings.First(x => x.ID == tmp.ID);
                context.Accountings.Remove(delAccount);
                context.SaveChanges();
                accounting.ItemsSource = context.Accountings.ToList();
                user.Clear();
                comp.Clear();
                dev.Clear();
                logAccounting.Foreground = Brushes.BlueViolet;
                logAccounting.Content = "Видалено успішно";
            }
        }

        //Звітність========Selector_OnSelect = four
        private void MouseUp_Report(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Accounting reportSeach = reportA.SelectedItem as Accounting;
                reportCompA.ItemsSource = context.Computers.Where(x => x.ID == reportSeach.ComputerID).ToList();
            }
            catch
            {
                reportCompA.ItemsSource = context.Computers.ToList();
            }
        }

        private void Click_LoadAllToExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                var reportExcel = new MaketExcelGeneratorAll().Generate();
                CommonSaveFileDialog dialog = new CommonSaveFileDialog();
                dialog.DefaultFileName = "ReportAll_" + DateTime.Now.ToString("ddMMyyyy_hhmmss");
                dialog.DefaultExtension = "xlsx";
                dialog.Filters.Add(new CommonFileDialogFilter("Excel Documents", "*.xlsx"));
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    File.WriteAllBytes(dialog.FileName, reportExcel);
                    MessageBox.Show("Звіт вдало вигружений у " + dialog.FileName);
                    Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
                }
            }
            catch { }
        }

        private void Click_LoadCompToExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                Computer obj = new Computer();
                obj = reportCompA.SelectedItem as Computer;
                var reportExcel = new MaketExcelGeneratorComp().Generate(obj);
                CommonSaveFileDialog dialog = new CommonSaveFileDialog();
                dialog.DefaultFileName = "PasportComp_" + obj.NamePC + DateTime.Now.ToString("_ddMMyyyy_hhmmss");
                dialog.DefaultExtension = "xlsx";
                dialog.Filters.Add(new CommonFileDialogFilter("Excel Documents", "*.xlsx"));
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    File.WriteAllBytes(dialog.FileName, reportExcel);
                    MessageBox.Show("Звіт вдало вигружений у " + dialog.FileName);
                    Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
                }
            }
            catch { }
        }

        private void Click_LoadDepartmentToExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                Department obj = new Department();
                var dep = context.Departments.ToList();
                foreach (var item in dep)
                {
                    if (item.Name == depSelect.SelectedItem.ToString())
                        obj.ID = item.ID;
                }
                var reportExcel = new MaketExcelGeneratorDepartment().Generate(obj);
                CommonSaveFileDialog dialog = new CommonSaveFileDialog();
                dialog.DefaultFileName = "ReportDepartment_" + DateTime.Now.ToString("ddMMyyyy_hhmmss");
                dialog.DefaultExtension = "xlsx";
                dialog.Filters.Add(new CommonFileDialogFilter("Excel Documents", "*.xlsx"));
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    File.WriteAllBytes(dialog.FileName, reportExcel);
                    MessageBox.Show("Звіт вдало вигружений у " + dialog.FileName);
                    Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
                }
            }
            catch { }
        }

        private void Click_LoadEmployeToExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                Employe obj = new Employe();
                var emp = context.Employes.ToList();
                foreach (var item in emp)
                {
                    if (item.Name == empSelect.SelectedItem.ToString())
                        obj.ID = item.ID;
                }
                var reportExcel = new MaketExcelGeneratorEmploye().Generate(obj);
                CommonSaveFileDialog dialog = new CommonSaveFileDialog();
                dialog.DefaultFileName = "ReportEmploye_" + DateTime.Now.ToString("ddMMyyyy_hhmmss");
                dialog.DefaultExtension = "xlsx";
                dialog.Filters.Add(new CommonFileDialogFilter("Excel Documents", "*.xlsx"));
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    File.WriteAllBytes(dialog.FileName, reportExcel);
                    MessageBox.Show("Звіт вдало вигружений у " + dialog.FileName);
                    Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
                }
            }
            catch { }
        }

        private void Click_LoadDeviceToExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                Device_ENUM obj = new Device_ENUM();
                var dev = context.Device_ENUMs.ToList();
                foreach (var item in dev)
                {
                    if (item.Name == devSelect.SelectedItem.ToString())
                        obj.ID = item.ID;
                }
                var reportExcel = new MaketExcelGeneratorDevice().Generate(obj);
                CommonSaveFileDialog dialog = new CommonSaveFileDialog();
                dialog.DefaultFileName = "ReportDevice_" + DateTime.Now.ToString("ddMMyyyy_hhmmss");
                dialog.DefaultExtension = "xlsx";
                dialog.Filters.Add(new CommonFileDialogFilter("Excel Documents", "*.xlsx"));
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    File.WriteAllBytes(dialog.FileName, reportExcel);
                    MessageBox.Show("Звіт вдало вигружений у " + dialog.FileName);
                    Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
                }
            }
            catch { }
        }
    }
}