using AICO_CL.Models;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Text;

namespace AICO_CL.Entity
{
    public class EfContext : DbContext
    {
        static SqlConnectionStringBuilder connectDB = new SqlConnectionStringBuilder();
        static string[] file = File.ReadAllLines(PathString());
        public EfContext() : base(ConnectServer(file)) { }
        public DbSet<Accounting> Accountings { get; set; }
        public DbSet<Computer> Computers { get; set; }
        public DbSet<Department> Departments { get; set; }
        public DbSet<Employe> Employes { get; set; }
        public DbSet<Device_ENUM> Device_ENUMs { get; set; }
        public DbSet<Device> Devices { get; set; }
        static private string PathString()
        {
            string pach = @"./../../../../configDB.cfg";
            if (!File.Exists(pach))
                pach = @"configDB.cfg";
            return pach;
        }
        static private string ConnectServer(string[] file)
        {
            connectDB.DataSource = file[0].Substring(14);
            connectDB.InitialCatalog = file[1].Substring(9);
            if (file[2].Substring(30) == "0")
                connectDB.IntegratedSecurity = true;
            else
            {
                connectDB.IntegratedSecurity = false;
                connectDB.UserID = file[3].Substring(9);
                connectDB.Password = file[4].Substring(13);
            }
            return connectDB.ConnectionString;
        }
    }
}