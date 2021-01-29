using AICO_CL.Entity;
using AICO_CL.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AICO_Desktop.View
{
    public class MaketExcelGeneratorAll
    {
        EfContext context = new EfContext();
        public byte[] Generate()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var reportAll = context.Accountings.ToList();
                var sheet = package.Workbook.Worksheets.Add("Maket Report");
                sheet.Cells["A1:Q1"].Merge = true;
                sheet.Cells["A1:Q1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A2:Q2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A1:Q1"].Style.Font.Size = 14;
                sheet.Cells["A1:Q1"].Style.Font.Bold = true;
                sheet.Cells["A2:Q2"].Style.Font.Bold = true;
                sheet.Cells["A1"].Value = "Список техніки закріпленої за працівниками";
                var row = 3;
                var column = 1;
                sheet.Cells["A2"].Value = "№п/п";
                sheet.Cells["B2"].Value = "Підрозділ";
                sheet.Cells["C2"].Value = "Працівник";
                sheet.Cells["D2"].Value = "Ім'я користувача";
                sheet.Cells["E2"].Value = "Ім'я комп'ютера";
                sheet.Cells["F2"].Value = "Процесор";
                sheet.Cells["G2"].Value = "Материнська плата";
                sheet.Cells["H2"].Value = "Оперативна пам'ять";
                sheet.Cells["I2"].Value = "Жорсткий диск";
                sheet.Cells["J2"].Value = "Операційна система";
                sheet.Cells["K2"].Value = "Пристрій";
                sheet.Cells["L2"].Value = "Модель пристрою";
                sheet.Cells["M2"].Value = "Опис 1";
                sheet.Cells["N2"].Value = "Опис 2";
                sheet.Cells["O2"].Value = "Опис 3";
                sheet.Cells["P2"].Value = "Опис 4";
                sheet.Cells["Q2"].Value = "Опис 5";
                for (int i = 0; i < reportAll.Count; i++)
                {
                    sheet.Cells[row, column++].Value = i + 1;
                    foreach (var e in context.Employes.ToList())
                        if (reportAll[i].EmployeID == e.ID)
                        {
                            foreach (var ed in context.Departments.ToList())
                                if (ed.ID == e.DepartmentID)
                                    sheet.Cells[row, column++].Value = ed.Name;
                            sheet.Cells[row, column++].Value = e.Name;
                        }
                    if (reportAll[i].ComputerID != null)
                    {
                        foreach (var c in context.Computers.ToList())
                            if (reportAll[i].ComputerID == c.ID)
                            {
                                sheet.Cells[row, column++].Value = c.UserNamePC;
                                sheet.Cells[row, column++].Value = c.NamePC;
                                sheet.Cells[row, column++].Value = c.CPUpc;
                                sheet.Cells[row, column++].Value = c.Motherboard;
                                sheet.Cells[row, column++].Value = c.RAMpc;
                                sheet.Cells[row, column++].Value = c.HDDpc;
                                sheet.Cells[row, column++].Value = c.OSVersion;
                            }
                    }
                    else
                        column = +11;
                    if (reportAll[i].DeviceID != null)
                    {
                        foreach (var d in context.Devices.ToList())
                            if (reportAll[i].DeviceID == d.ID)
                            {
                                foreach (var de in context.Device_ENUMs.ToList())
                                    if (de.ID == d.Device_ENUM_ID)
                                        sheet.Cells[row, column++].Value = de.Name;
                                sheet.Cells[row, column++].Value = d.Model;
                                sheet.Cells[row, column++].Value = d.Description_1;
                                sheet.Cells[row, column++].Value = d.Description_2;
                                sheet.Cells[row, column++].Value = d.Description_3;
                                sheet.Cells[row, column++].Value = d.Description_4;
                                sheet.Cells[row, column++].Value = d.Description_5;
                            }
                    }
                    else
                        column = +17;
                    row++;
                    column = 1;
                }
                sheet.Cells.AutoFitColumns();
                using (var range = sheet.Cells[2, 1, row - 1, 17])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
                return package.GetAsByteArray();
            };
        }
    }

    public class MaketExcelGeneratorComp
    {
        EfContext context = new EfContext();
        public byte[] Generate(Accounting obj)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var reportAll = context.Accountings.ToList();
                var sheet = package.Workbook.Worksheets.Add("Maket Report");
                sheet.Cells["A1:D1"].Merge = true;
                sheet.Cells["A1:D1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A1:D1"].Style.Font.Size = 18;
                sheet.Cells["A1:D1"].Style.Font.Bold = true;
                sheet.Cells["A1"].Value = "Паспорт ПК";
                sheet.Cells["B2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells["B2"].Style.Font.Bold = true;
                sheet.Cells["B2"].Value = "Підрозділ: ";
                sheet.Cells["C2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                Employe empl = new Employe();
                foreach (var item in context.Employes.ToList())
                {
                    if (item.ID == obj.EmployeID)
                    {
                        foreach (var d in context.Departments.ToList())
                            if (d.ID == item.DepartmentID)
                                sheet.Cells["C2"].Value = d.Name;
                    }
                }
                sheet.Cells["B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells["B3"].Style.Font.Bold = true;
                sheet.Cells["B3"].Value = "Працівник: ";
                sheet.Cells["C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells["C3"].Value = obj.Employes.Name;
                sheet.Cells["B5:B11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["B5:B11"].Style.Font.Bold = true;
                sheet.Cells["B5"].Value = "№ п/п";
                sheet.Cells["B6"].Value = "1";
                sheet.Cells["B7"].Value = "2";
                sheet.Cells["B8"].Value = "3";
                sheet.Cells["B9"].Value = "4";
                sheet.Cells["B10"].Value = "5";
                sheet.Cells["B11"].Value = "6";
                sheet.Cells["C5:D5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["C5:D5"].Style.Font.Bold = true;
                sheet.Cells["C5"].Value = "Назва комплектуючих";
                sheet.Cells["D5"].Value = "Модель комплектуючих";
                sheet.Cells["C6:C11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells["C6:C11"].Style.Font.Bold = true;
                sheet.Cells["C6"].Value = "Ім'я комп'ютера";
                sheet.Cells["C7"].Value = "Процесор";
                sheet.Cells["C8"].Value = "Материнська плата";
                sheet.Cells["C9"].Value = "Оперативна пам'ять";
                sheet.Cells["C10"].Value = "Жорсткий диск";
                sheet.Cells["C11"].Value = "Операційна система";
                Computer comp = new Computer();
                foreach (var item in context.Computers.ToList())
                {
                    if (item.ID == obj.ComputerID)
                        comp = item;
                }
                sheet.Cells["D6:D11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells["D6"].Value = comp.NamePC;
                sheet.Cells["D7"].Value = comp.CPUpc;
                sheet.Cells["D8"].Value = comp.Motherboard;
                sheet.Cells["D9"].Value = comp.RAMpc;
                sheet.Cells["D10"].Value = comp.HDDpc;
                sheet.Cells["D11"].Value = comp.OSVersion + " " + comp.BitOperating;
                sheet.Cells["B16:D16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["B16"].Value = "(Дата)";
                sheet.Cells["C16"].Value = "(Підпис)";
                sheet.Cells["D16"].Value = obj.Employes.Name;
                sheet.Cells.AutoFitColumns();
                using (var range = sheet.Cells[5, 2, 11, 4])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
                using (var range = sheet.Cells[15, 2, 15, 4])
                {
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }

                var sheet2 = package.Workbook.Worksheets.Add("Maket Reverse");
                sheet2.Row(1).Height = 33;
                sheet2.Column(1).Width = 1;
                sheet2.Column(2).Width = 7;
                sheet2.Column(3).Width = 9;
                sheet2.Column(4).Width = 32;
                sheet2.Column(5).Width = 14;
                sheet2.Column(6).Width = 14;
                sheet2.Column(7).Width = 9;
                sheet2.Cells["B1:G1"].Merge = true;
                sheet2.Cells["B1:G1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet2.Cells["B1:G1"].Style.Font.Size = 12;
                sheet2.Cells["B1:G1"].Style.Font.Bold = true;
                sheet2.Cells["B1:G1"].Style.WrapText = true;
                sheet2.Cells["B1"].Value = "Відомості про встановлення та видалення ПЗ, ремонт, технічному обслуговуванні, зміни програмної конфігурації ЕОМ";
                sheet2.Cells["B2:G2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet2.Cells["B2:G2"].Style.Font.Size = 12;
                sheet2.Cells["B2:G2"].Style.Font.Bold = true;
                sheet2.Cells["B2"].Value = "№ п/п";
                sheet2.Cells["C2"].Value = "Дата";
                sheet2.Cells["D2"].Value = "Дія";
                sheet2.Cells["E2"].Value = "Підстава";
                sheet2.Cells["F2"].Value = "Виконав, ПІБ";
                sheet2.Cells["G2"].Value = "Підписи";
                sheet2.Cells["B3:G3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet2.Cells["B3:G3"].Style.Font.Size = 12;
                sheet2.Cells["B3:G3"].Style.Font.Bold = true;
                sheet2.Cells["B3"].Value = "1";
                sheet2.Cells["C3"].Value = "2";
                sheet2.Cells["D3"].Value = "3";
                sheet2.Cells["E3"].Value = "4";
                sheet2.Cells["F3"].Value = "5";
                sheet2.Cells["G3"].Value = "6";
                using (var range = sheet2.Cells[2, 2, 48, 7])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
                return package.GetAsByteArray();
            };
        }
    }

    public class MaketExcelGeneratorDepartment
    {
        EfContext context = new EfContext();
        public byte[] Generate(Department obj)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var row = 3;
                var column = 1;
                int pp = 1;
                List<Employe> listEmp = new List<Employe>();
                var tmpEmp = context.Employes.ToList();
                for (int i = 0; i < tmpEmp.Count(); i++)
                {
                    if (obj.ID == tmpEmp[i].DepartmentID)
                        listEmp.Add(tmpEmp[i]);
                }
                List<Accounting> listAcc = new List<Accounting>();
                var tmpAcc = context.Accountings.ToList();
                var reportDepartment = context.Departments.ToList();
                var sheet = package.Workbook.Worksheets.Add("Maket Report");
                sheet.Cells["A1:P1"].Merge = true;
                sheet.Cells["A1:P1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A2:P2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A1:P1"].Style.Font.Size = 14;
                sheet.Cells["A1:P1"].Style.Font.Bold = true;
                sheet.Cells["A2:P2"].Style.Font.Bold = true;
                var _enum = context.Departments.ToList();
                foreach (var item in _enum)
                    if (obj.ID == item.ID)
                        sheet.Cells["A1"].Value = "Перелік оргтехніки підрозділу: " + item.Name;
                sheet.Cells["A2"].Value = "№п/п";
                sheet.Cells["B2"].Value = "Працівник";
                sheet.Cells["C2"].Value = "Ім'я користувача";
                sheet.Cells["D2"].Value = "Ім'я комп'ютера";
                sheet.Cells["E2"].Value = "Процесор";
                sheet.Cells["F2"].Value = "Материнська плата";
                sheet.Cells["G2"].Value = "Оперативна пам'ять";
                sheet.Cells["H2"].Value = "Жорсткий диск";
                sheet.Cells["I2"].Value = "Операційна система";
                sheet.Cells["J2"].Value = "Пристрій";
                sheet.Cells["K2"].Value = "Модель пристрою";
                sheet.Cells["L2"].Value = "Опис 1";
                sheet.Cells["M2"].Value = "Опис 2";
                sheet.Cells["N2"].Value = "Опис 3";
                sheet.Cells["O2"].Value = "Опис 4";
                sheet.Cells["P2"].Value = "Опис 5";
                for (int i = 0; i < tmpAcc.Count(); i++)
                {
                    foreach (var item in listEmp)
                    {
                        if (item.ID == tmpAcc[i].EmployeID)
                        {
                            sheet.Cells[row, column++].Value = pp++;
                            sheet.Cells[row, column++].Value = item.Name;
                            if (tmpAcc[i].ComputerID != null)
                            {
                                foreach (var c in context.Computers.ToList())
                                    if (tmpAcc[i].ComputerID == c.ID)
                                    {
                                        sheet.Cells[row, column++].Value = c.UserNamePC;
                                        sheet.Cells[row, column++].Value = c.NamePC;
                                        sheet.Cells[row, column++].Value = c.CPUpc;
                                        sheet.Cells[row, column++].Value = c.Motherboard;
                                        sheet.Cells[row, column++].Value = c.RAMpc;
                                        sheet.Cells[row, column++].Value = c.HDDpc;
                                        sheet.Cells[row, column++].Value = c.OSVersion;
                                    }
                            }
                            else
                                column = +10;
                            if (tmpAcc[i].DeviceID != null)
                            {
                                foreach (var d in context.Devices.ToList())
                                    if (tmpAcc[i].DeviceID == d.ID)
                                    {
                                        foreach (var de in context.Device_ENUMs.ToList())
                                            if (d.Device_ENUM_ID == de.ID)
                                                sheet.Cells[row, column++].Value = de.Name;
                                        sheet.Cells[row, column++].Value = d.Model;
                                        sheet.Cells[row, column++].Value = d.Description_1;
                                        sheet.Cells[row, column++].Value = d.Description_2;
                                        sheet.Cells[row, column++].Value = d.Description_3;
                                        sheet.Cells[row, column++].Value = d.Description_4;
                                        sheet.Cells[row, column++].Value = d.Description_5;
                                    }
                            }
                            else
                                column = +16;
                            row++;
                            column = 1;
                        }
                    }
                }
                sheet.Cells.AutoFitColumns();
                using (var range = sheet.Cells[2, 1, row - 1, 16])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
                return package.GetAsByteArray();
            };
        }
    }

    public class MaketExcelGeneratorEmploye
    {
        EfContext context = new EfContext();
        public byte[] Generate(Employe obj)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var row = 3;
                var column = 1;
                int pp = 1;
                var tmpAcc = context.Accountings.ToList();
                var sheet = package.Workbook.Worksheets.Add("Maket Report");
                sheet.Cells["A1:P1"].Merge = true;
                sheet.Cells["A1:P1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A2:P2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A1:P1"].Style.Font.Size = 14;
                sheet.Cells["A1:P1"].Style.Font.Bold = true;
                sheet.Cells["A2:P2"].Style.Font.Bold = true;
                var _enum = context.Employes.ToList();
                foreach (var item in _enum)
                    if (obj.ID == item.ID)
                        sheet.Cells["A1"].Value = "Перелік оргтехніки працівника: " + item.Name;
                sheet.Cells["A2"].Value = "№п/п";
                sheet.Cells["B2"].Value = "Ім'я користувача";
                sheet.Cells["C2"].Value = "Ім'я комп'ютера";
                sheet.Cells["D2"].Value = "Процесор";
                sheet.Cells["E2"].Value = "Материнська плата";
                sheet.Cells["F2"].Value = "Оперативна пам'ять";
                sheet.Cells["G2"].Value = "Жорсткий диск";
                sheet.Cells["H2"].Value = "Операційна система";
                sheet.Cells["I2"].Value = "Пристрій";
                sheet.Cells["J2"].Value = "Модель пристрою";
                sheet.Cells["K2"].Value = "Опис 1";
                sheet.Cells["L2"].Value = "Опис 2";
                sheet.Cells["M2"].Value = "Опис 3";
                sheet.Cells["N2"].Value = "Опис 4";
                sheet.Cells["O2"].Value = "Опис 5";
                for (int i = 0; i < tmpAcc.Count(); i++)
                {
                    if (obj.ID == tmpAcc[i].EmployeID)
                    {
                        sheet.Cells[row, column++].Value = pp++;
                        if (tmpAcc[i].ComputerID != null)
                        {
                            foreach (var c in context.Computers.ToList())
                                if (tmpAcc[i].ComputerID == c.ID)
                                {
                                    sheet.Cells[row, column++].Value = c.UserNamePC;
                                    sheet.Cells[row, column++].Value = c.NamePC;
                                    sheet.Cells[row, column++].Value = c.CPUpc;
                                    sheet.Cells[row, column++].Value = c.Motherboard;
                                    sheet.Cells[row, column++].Value = c.RAMpc;
                                    sheet.Cells[row, column++].Value = c.HDDpc;
                                    sheet.Cells[row, column++].Value = c.OSVersion;
                                }
                        }
                        else
                            column = +9;
                        if (tmpAcc[i].DeviceID != null)
                        {
                            foreach (var d in context.Devices.ToList())
                                if (tmpAcc[i].DeviceID == d.ID)
                                {
                                    foreach (var de in context.Device_ENUMs.ToList())
                                        if (d.Device_ENUM_ID == de.ID)
                                            sheet.Cells[row, column++].Value = de.Name;
                                    sheet.Cells[row, column++].Value = d.Model;
                                    sheet.Cells[row, column++].Value = d.Description_1;
                                    sheet.Cells[row, column++].Value = d.Description_2;
                                    sheet.Cells[row, column++].Value = d.Description_3;
                                    sheet.Cells[row, column++].Value = d.Description_4;
                                    sheet.Cells[row, column++].Value = d.Description_5;
                                }
                        }
                        else
                            column = +15;
                        row++;
                        column = 1;
                    }
                }
                sheet.Cells.AutoFitColumns();
                using (var range = sheet.Cells[2, 1, row - 1, 15])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
                return package.GetAsByteArray();
            };
        }
    }

    public class MaketExcelGeneratorDevice
    {
        EfContext context = new EfContext();
        public byte[] Generate(Device_ENUM obj)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var row = 3;
                var column = 1;
                int pp = 1;
                var reportDevice = context.Devices.ToList();
                var sheet = package.Workbook.Worksheets.Add("Maket Report");
                sheet.Cells["A1:G1"].Merge = true;
                sheet.Cells["A1:G1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A1:G1"].Style.Font.Size = 18;
                sheet.Cells["A1:G1"].Style.Font.Bold = true;
                var _enum = context.Device_ENUMs.ToList();
                foreach (var item in _enum)
                    if (obj.ID == item.ID)
                        sheet.Cells["A1"].Value = "Перелік оргтехніки типу: " + item.Name;
                sheet.Cells["A2:G2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A2:G2"].Style.Font.Bold = true;
                sheet.Cells["A2"].Value = "№п/п";
                sheet.Cells["B2"].Value = "Модель пристрою";
                sheet.Cells["C2"].Value = "Опис 1";
                sheet.Cells["D2"].Value = "Опис 2";
                sheet.Cells["E2"].Value = "Опис 3";
                sheet.Cells["F2"].Value = "Опис 4";
                sheet.Cells["G2"].Value = "Опис 5";
                foreach (var d in reportDevice)
                    if (obj.ID == d.Device_ENUM_ID)
                    {
                        sheet.Cells[row, column++].Value = pp++;
                        sheet.Cells[row, column++].Value = d.Model;
                        sheet.Cells[row, column++].Value = d.Description_1;
                        sheet.Cells[row, column++].Value = d.Description_2;
                        sheet.Cells[row, column++].Value = d.Description_3;
                        sheet.Cells[row, column++].Value = d.Description_4;
                        sheet.Cells[row, column++].Value = d.Description_5;
                        row++;
                        column = 1;
                    }
                sheet.Cells.AutoFitColumns();
                using (var range = sheet.Cells[2, 1, row - 1, 7])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
                return package.GetAsByteArray();
            };
        }
    }
}