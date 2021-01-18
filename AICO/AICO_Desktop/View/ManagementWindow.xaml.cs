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
using OfficeOpenXml;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace AICO_Desktop
{
    public partial class ManagementWindow : Window
    {
        Excel.Workbook fileExcel;
        Computer myComp;
        EfContext context;
        ObservableCollection<Department> nodeDep;
        public Employe root { get; set; }
        public ManagementWindow(Employe user)
        {
            InitializeComponent();
            context = new EfContext();
            root = user;
            this.Title = root.Name + ", " + root.Work + ", " + root.Phone;
        }

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
                Computer newComp = new Computer();
                newComp = context.Computers.FirstOrDefault(x => x.NamePC == myComp.NamePC);
                Accounting userTmp = new Accounting();
                if (newComp != null)
                {
                    userTmp = context.Accountings.FirstOrDefault(x => x.ComputerID == newComp.ID);
                    //if (userTmp != null)
                    //{
                    //    userPCDB.Content = "Обладнання закріплене за ПК " + userTmp.Employes.Name;
                    //}
                    //else
                    //{
                    //    userPCDB.Content = "Обладнання закріплене за ПК. ";
                    //}
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
                employe.ItemsSource = context.Employes.ToList();
                if (context.Departments.Count() != 0)
                    departmentsName.ItemsSource = context.Departments.Select(x => x.Name).ToList();
                else
                    departmentsName.ItemsSource = "";
            }
            if ((tcSample.SelectedItem as TabItem).Name == "three")
            {
                device_ENUM.ItemsSource = context.Device_ENUMs.ToList();
                device.ItemsSource = context.Devices.ToList();
                if (context.Device_ENUMs.Count() != 0)
                    deviceENUM.ItemsSource = context.Device_ENUMs.Select(x => x.Name).ToList();
                else
                    deviceENUM.ItemsSource = "";
            }
            if ((tcSample.SelectedItem as TabItem).Name == "four")
            {
                departmentA.ItemsSource = context.Departments.ToList();
                employeA.ItemsSource = context.Employes.ToList();
                computerA.ItemsSource = context.Computers.ToList();
                device_ENUMA.ItemsSource = context.Device_ENUMs.ToList();
                deviceA.ItemsSource = context.Devices.ToList();
                accounting.ItemsSource = context.Accountings.ToList();
            }
            if ((tcSample.SelectedItem as TabItem).Name == "five")
            {
                reportA.ItemsSource = context.Accountings.ToList();
            }
        }
        //=====================================================Керування звітністю
        private void MouseUp_Report(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Accounting reportSeach = new Accounting();
                reportSeach = reportA.SelectedItem as Accounting;
                reportCompA.ItemsSource = context.Computers.Where(x => x.ID == reportSeach.ComputerID).ToList();
            }
            catch
            {
                reportCompA.ItemsSource = context.Computers.ToList();
            }
        }
        private void Click_LoadCompToExcel(object sender, RoutedEventArgs e)
        {
            var reportData = new MaketReport().GetReport();
            var reportExcel = new MaketExcelGenerator().Generate(reportData);
            File.WriteAllBytes("D:/ReportComp.xlsx", reportExcel);
            //Excel.Application excel = new Excel.Application();
            //fileExcel = excel.Workbooks.Open("D:/Report.xlsx");
            MessageBox.Show("Вигрузка Comp в Exel");
        }

        private void Click_LoadAllToExcel(object sender, RoutedEventArgs e)
        {
            var reportData = new MaketReport().GetReport();
            var reportExcel = new MaketExcelGenerator().Generate(reportData);
            File.WriteAllBytes("D:/ReportAll.xlsx", reportExcel);
            //Excel.Application excel = new Excel.Application();
            //fileExcel = excel.Workbooks.Open("D:/Report.xlsx");
            MessageBox.Show("Вигрузка All в Exel");
        }
        private void Expanded_DepTree(object sender, RoutedEventArgs e)
        {
            //TreeViewItem tvItem = (TreeViewItem)sender;
            //MessageBox.Show("Узел " + tvItem.Header.ToString() + " раскрыт");
        }

        private void Selected_EmpTree(object sender, RoutedEventArgs e)
        {
            //TreeViewItem tvItem = (TreeViewItem)sender;
            //MessageBox.Show("Выбран узел: " + tvItem.Header.ToString());
        }



        //=====================================================Керування обліком
        private void Click_AddAccounting(object sender, RoutedEventArgs e)
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

        private void Click_EditAccounting(object sender, RoutedEventArgs e)
        {
            Accounting editaccount = new Accounting();
            editaccount = accounting.SelectedItem as Accounting;
            if (editaccount != null)
            {
                using (EfContext dbContext = new EfContext())
                {
                    try
                    {
                        dbContext.Accountings.Attach(editaccount);
                        var userTMP = dbContext.Employes.FirstOrDefault(x => x.Name == user.Text);
                        if (userTMP != null)
                            editaccount.EmployeID = userTMP.ID;
                        else
                        {
                            logAccounting.Foreground = Brushes.Red;
                            logAccounting.Content = "Користувач неможе бути порожнім";
                        }
                        var compTMP = dbContext.Computers.FirstOrDefault(x => x.NamePC == comp.Text);
                        if (compTMP != null)
                            editaccount.ComputerID = compTMP.ID;
                        else
                            editaccount.ComputerID = null;
                        var devTMP = dbContext.Devices.FirstOrDefault(x => x.Model == dev.Text);
                        if (devTMP != null)
                            editaccount.DeviceID = devTMP.ID;
                        else
                            editaccount.DeviceID = null;
                    }
                    catch { };
                    dbContext.Entry(editaccount).State = System.Data.Entity.EntityState.Modified;
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

        private void Click_DeleteAccounting(object sender, RoutedEventArgs e)
        {
            Accounting delAccount = new Accounting();
            delAccount = accounting.SelectedItem as Accounting;
            context.Accountings.Remove(delAccount);
            context.SaveChanges();
            accounting.ItemsSource = context.Accountings.ToList();
            user.Clear();
            comp.Clear();
            dev.Clear();
            logAccounting.Foreground = Brushes.BlueViolet;
            logAccounting.Content = "Видалено успішно";
        }

        private void MouseUp_Accounting(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Accounting tmp = new Accounting();
                tmp = accounting.SelectedItem as Accounting;
                if (tmp.Employes.Name != null)
                    user.Text = tmp.Employes.Name;
                if (tmp.Computers.NamePC.ToString() != null)
                    comp.Text = tmp.Computers.NamePC;
                if (tmp.Devices.Model != null)
                    dev.Text = tmp.Devices.Model;
            }
            catch
            {
                //user.Clear();
                //comp.Clear();
                //dev.Clear();
            }
        }

        private void MouseUp_User(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Employe tmpUser = new Employe();
                tmpUser = employeA.SelectedItem as Employe;
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
                Computer tmpComp = new Computer();
                tmpComp = computerA.SelectedItem as Computer;
                comp.Text = tmpComp.NamePC;
            }
            catch
            {
                comp.Clear();
            }
        }
        private void MouseUp_Dev(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Device tmpDev = new Device();
                tmpDev = deviceA.SelectedItem as Device;
                dev.Text = tmpDev.Model;
            }
            catch
            {
                dev.Clear();
            }
        }
        //=====================================================Керування пристроями
        private void Click_AddDevice(object sender, RoutedEventArgs e)
        {
            var tmpDevice = context.Devices.FirstOrDefault(x => x.Model == model.Text);
            if (tmpDevice == null)
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
                logDevice.Foreground = Brushes.Green;
                logDevice.Content = "Пристрій додано";
            }
            else
            {
                logDevice.Foreground = Brushes.Red;
                logDevice.Content = "Такий пистрій вже існує";
            }
        }

        private void Click_EditDevice(object sender, RoutedEventArgs e)
        {
            Device newDevice = new Device();
            newDevice = device.SelectedItem as Device;
            using (var dbContext = new EfContext())
            {
                newDevice.Model = model.Text;
                newDevice.Description_1 = description1.Text;
                newDevice.Description_2 = description2.Text;
                newDevice.Description_3 = description3.Text;
                newDevice.Description_4 = description4.Text;
                newDevice.Description_5 = description5.Text;
                if (deviceENUM.SelectedItem != null)
                {
                    if (dbContext.Device_ENUMs.Count() != 0)
                        deviceENUM.ItemsSource = dbContext.Device_ENUMs.Select(x => x.Name).ToList();
                    else
                        deviceENUM.ItemsSource = "";
                }
                else
                {
                    var list = dbContext.Device_ENUMs.ToList();
                    foreach (var item in list)
                        if (item.Name == deviceENUM.SelectedItem.ToString())
                            newDevice.Device_ENUM_ID = item.ID;
                }
                dbContext.Devices.Attach(newDevice);
                dbContext.Entry(newDevice).State = System.Data.Entity.EntityState.Modified;
                dbContext.SaveChanges();
            }
            device.ItemsSource = context.Devices.ToList();
            model.Clear();
            description1.Clear();
            description2.Clear();
            description3.Clear();
            description4.Clear();
            description5.Clear();
            logDevice.Foreground = Brushes.Green;
            logDevice.Content = "Зміни внесено успішно";
        }

        private void Click_DeleteDevice(object sender, RoutedEventArgs e)
        {
            Device newDevice = new Device();
            newDevice = device.SelectedItem as Device;
            context.Devices.Remove(newDevice);
            context.SaveChanges();
            device.ItemsSource = context.Devices.ToList();
            model.Clear();
            description1.Clear();
            description2.Clear();
            description3.Clear();
            description4.Clear();
            description5.Clear();
            deviceENUM.ItemsSource = "";
            logDevice.Foreground = Brushes.BlueViolet;
            logDevice.Content = "Пристрій видалено";
        }

        private void MouseDuble_Device(object sender, MouseButtonEventArgs e)
        {
            logDevice.Content = "";
            try
            {
                Device editDevice = new Device();
                editDevice = device.SelectedItem as Device;
                model.Text = editDevice.Model;
                description1.Text = editDevice.Description_1;
                description2.Text = editDevice.Description_2;
                description3.Text = editDevice.Description_3;
                description4.Text = editDevice.Description_4;
                description5.Text = editDevice.Description_5;
                deviceENUM.SelectedItem = editDevice.Devices_ENUM.Name;
            }
            catch (NullReferenceException ex)
            {
                deviceENUM.SelectedItem = "";
            }
            catch
            {
                device.ItemsSource = context.Devices.ToList();
                model.Clear();
                description1.Clear();
                description2.Clear();
                description3.Clear();
                description4.Clear();
                description5.Clear();
                deviceENUM.ItemsSource = "";
            }
        }

        private void Click_AddDevENUM(object sender, RoutedEventArgs e)
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
                logDevice.Foreground = Brushes.Green;
                logDevice.Content = "Тип пристрою додано";
            }
            else
            {
                logDevice.Foreground = Brushes.Red;
                logDevice.Content = "Такий тип вже існує";
            }
        }
        private void Click_EditDevENUM(object sender, RoutedEventArgs e)
        {
            Device_ENUM newDevEnum = new Device_ENUM();
            newDevEnum = device_ENUM.SelectedItem as Device_ENUM;
            using (var dbContext = new EfContext())
            {
                newDevEnum.Name = devENUM_Text.Text;
                dbContext.Device_ENUMs.Attach(newDevEnum);
                dbContext.Entry(newDevEnum).State = System.Data.Entity.EntityState.Modified;
                dbContext.SaveChanges();
            }
            device_ENUM.ItemsSource = context.Device_ENUMs.ToList();
            devENUM_Text.Clear();
            logDevice.Foreground = Brushes.Green;
            logDevice.Content = "Зміни внесено успішно";
        }
        private void Click_DeleteDevENUM(object sender, RoutedEventArgs e)
        {
            Device_ENUM newDevEnum = new Device_ENUM();
            newDevEnum = device_ENUM.SelectedItem as Device_ENUM;
            context.Device_ENUMs.Remove(newDevEnum);
            context.SaveChanges();
            device_ENUM.ItemsSource = context.Device_ENUMs.ToList();
            devENUM_Text.Clear();
            logDevice.Foreground = Brushes.BlueViolet;
            logDevice.Content = "Тип видалено";
        }
        private void MouseUp_DeviceENUM(object sender, MouseButtonEventArgs e)
        {
            logDevice.Content = "";
            try
            {
                Device_ENUM editDevEnum = new Device_ENUM();
                editDevEnum = device_ENUM.SelectedItem as Device_ENUM;
                device.ItemsSource = context.Devices.Where(x => x.Devices_ENUM.ID == editDevEnum.ID).ToList();
            }
            catch
            {
                device_ENUM.ItemsSource = context.Device_ENUMs.ToList();
                devENUM_Text.Clear();
            }
        }

        private void MouseDuble_DeviceENUM(object sender, MouseButtonEventArgs e)
        {
            logDevice.Content = "";
            try
            {
                Device_ENUM editDevEnum = new Device_ENUM();
                editDevEnum = device_ENUM.SelectedItem as Device_ENUM;
                devENUM_Text.Text = editDevEnum.Name;
            }
            catch
            {
                device_ENUM.ItemsSource = context.Device_ENUMs.ToList();
                devENUM_Text.Clear();
            }
        }
        //=====================================================Керування працівниками та підрозділами
        private void Click_AddEmloye(object sender, RoutedEventArgs e)
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
                logUser.Foreground = Brushes.Green;
                logUser.Content = "Працівника додано";
            }
            else
            {
                logUser.Foreground = Brushes.Red;
                logUser.Content = "Такий працівник вже існує";
            }
        }

        private void Click_EditEmloye(object sender, RoutedEventArgs e)
        {
            Employe newUser = new Employe();
            newUser = employe.SelectedItem as Employe;
            using (var dbContext = new EfContext())
            {
                newUser.Name = name.Text;
                newUser.Work = work.Text;
                newUser.Phone = phone.Text;
                if (departmentsName.SelectedItem != null)
                {
                    if (context.Departments.Count() != 0)
                        departmentsName.ItemsSource = context.Departments.Select(x => x.Name).ToList();
                    else
                        departmentsName.ItemsSource = "";
                }
                else
                {
                    var list = dbContext.Departments.ToList();
                    foreach (var item in list)
                        if (item.Name == departmentsName.SelectedItem.ToString())
                            newUser.DepartmentID = item.ID;
                }
                dbContext.Employes.Attach(newUser);
                dbContext.Entry(newUser).State = System.Data.Entity.EntityState.Modified;
                context.SaveChanges();
            }
            employe.ItemsSource = context.Employes.ToList();
            name.Clear();
            work.Clear();
            phone.Clear();
            logUser.Foreground = Brushes.Green;
            logUser.Content = "Зміни внесено успішно";
        }

        private void Click_DeleteEmloye(object sender, RoutedEventArgs e)
        {
            Employe newUser = new Employe();
            newUser = employe.SelectedItem as Employe;
            context.Employes.Remove(newUser);
            context.SaveChanges();
            employe.ItemsSource = context.Employes.ToList();
            name.Clear();
            work.Clear();
            phone.Clear();
            departmentsName.SelectedItem = "";
            logUser.Foreground = Brushes.BlueViolet;
            logUser.Content = "Працівника видалено";
        }

        private void MouseDuble_Employe(object sender, MouseButtonEventArgs e)
        {
            logUser.Content = "";
            try
            {
                Employe editUser = new Employe();
                editUser = employe.SelectedItem as Employe;
                name.Text = editUser.Name;
                work.Text = editUser.Work;
                phone.Text = editUser.Phone;
                departmentsName.SelectedItem = editUser.Departments.Name;
            }
            catch (NullReferenceException ex)
            {
                departmentsName.SelectedItem = "";
            }
            catch
            {
                employe.ItemsSource = context.Employes.ToList();
                name.Clear();
                work.Clear();
                phone.Clear();
                departmentsName.SelectedItem = "";
            }
        }

        private void Click_AddDep(object sender, RoutedEventArgs e)
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
                logUser.Foreground = Brushes.Green;
                logUser.Content = "Відділ додано";
            }
            else
            {
                logUser.Foreground = Brushes.Red;
                logUser.Content = "Такий відділ вже існує";
            }
        }
        private void Click_EditDep(object sender, RoutedEventArgs e)
        {
            Department newDep = new Department();
            newDep = department.SelectedItem as Department;
            using (var dbContext = new EfContext())
            {
                newDep.Name = depText.Text;
                dbContext.Departments.Attach(newDep);
                dbContext.Entry(newDep).State = System.Data.Entity.EntityState.Modified;
                dbContext.SaveChanges();
            }
            department.ItemsSource = context.Departments.ToList();
            depText.Clear();
            logUser.Foreground = Brushes.Green;
            logUser.Content = "Зміни внесено успішно";
        }
        private void Click_DeleteDep(object sender, RoutedEventArgs e)
        {
            Department newDep = new Department();
            newDep = department.SelectedItem as Department;
            context.Departments.Remove(newDep);
            context.SaveChanges();
            department.ItemsSource = context.Departments.ToList();
            depText.Clear();
            logUser.Foreground = Brushes.BlueViolet;
            logUser.Content = "Відділ видалено";
        }
        private void MouseUp_Dep(object sender, MouseButtonEventArgs e)
        {
            logUser.Content = "";
            try
            {
                Department editdep = new Department();
                editdep = department.SelectedItem as Department;
                employe.ItemsSource = context.Employes.Where(x => x.Departments.ID == editdep.ID).ToList();
            }
            catch
            {
                department.ItemsSource = context.Departments.ToList();
                depText.Clear();
            }
        }

        private void MouseDuble_Dep(object sender, MouseButtonEventArgs e)
        {
            logUser.Content = "";
            try
            {
                Department editdep = new Department();
                editdep = department.SelectedItem as Department;
                depText.Text = editdep.Name;
            }
            catch
            {
                department.ItemsSource = context.Departments.ToList();
                depText.Clear();
            }
        }

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
            Computer compEdit = new Computer();
            compEdit = context.Computers.FirstOrDefault(x => x.NamePC == _lb1.Content.ToString());
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
        //=========================Облік техніки
        private void MouseUp_DepartmentA(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Department editdep = new Department();
                editdep = departmentA.SelectedItem as Department;
                employeA.ItemsSource = context.Employes.Where(x => x.Departments.ID == editdep.ID).ToList();
            }
            catch
            {
                departmentA.ItemsSource = context.Departments.ToList();
            }
        }

        private void MouseUp_DeviceENUMA(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Device_ENUM editDevEnum = new Device_ENUM();
                editDevEnum = device_ENUMA.SelectedItem as Device_ENUM;
                deviceA.ItemsSource = context.Devices.Where(x => x.Devices_ENUM.ID == editDevEnum.ID).ToList();
            }
            catch
            {
                device_ENUMA.ItemsSource = context.Device_ENUMs.ToList();
            }
        }
    }
}
