using AICO_CL.Entity;
using AICO_CL.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AICO_Desktop.View
{
    public class ReportCompany
    {
        public DepartmentItem[] DepItem { get; set; }
        public EmployeItem[] EmpItem { get; set; }
    }

    public class DepartmentItem
    {
        public string Name { get; set; }
    }
    public class EmployeItem
    {
        public string Name { get; set; }
        public string Work { get; set; }
        public string Phone { get; set; }
    }

    public class MaketReport
    {
        EfContext context = new EfContext();
        public ReportCompany GetReport()
        {
            var listDep = context.Departments.ToList();
            var listEmp = context.Employes.ToList();
            return new ReportCompany
            {
                DepItem = new[]
                {
                new DepartmentItem{Name="IT"}
            },
                EmpItem = new[]
                {
                    new EmployeItem{Name="Admin", Work="Інженер комп систем", Phone="+380987202713"}
                }
            };
        }
    }

    public class MaketExcelGenerator
    {
        public byte[] Generate(ReportCompany report)
        {
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Maket Report");
            sheet.Cells["B1"].Value = "Список підрозділів та працівників";
            sheet.Cells[3,2, 3,3].LoadFromArrays(new object[][] { new[] { "Підрозділ", "Працівник" } });
            var row = 4;
            var column = 2;
            foreach(var item in report.DepItem)
            {
                sheet.Cells[row, column].Value = item.Name;
            }
            return package.GetAsByteArray();
        }
    }
        
}
