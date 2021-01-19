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
                var row = 2;
                var column = 1;
                for (int i = 0; i < reportAll.Count; i++)
                {

                    sheet.Cells[row, column++].Value = reportAll[i].ID;
                    //sheet.Cells[row, column++].Value = reportAll[i].EmployeID;
                    foreach (var e in context.Employes.ToList())
                        if (reportAll[i].EmployeID == e.ID)
                        {
                            foreach (var ed in context.Departments.ToList())
                                if (ed.ID == e.DepartmentID)
                                    sheet.Cells[row, column++].Value = ed.Name;
                            sheet.Cells[row, column++].Value = e.Name;
                        }
                    //sheet.Cells[row, column++].Value = reportAll[i].ComputerID;
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
                    //sheet.Cells[row, column++].Value = reportAll[i].DeviceID;
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
                return package.GetAsByteArray();
            };
        }
    }
}
