using AICO_CL.Entity;
using AICO_CL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ConsoleAICO
{
    class Program
    {
        static Computer myComp;
        static List<Computer> newComp;
        static EfContext context;
        static void Main(string[] args)
        {
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
            myComp.ShowPC();
            var temp = context.Computers.FirstOrDefault(x => x.NamePC == myComp.NamePC);
            if (temp == null)
            {
                context.Computers.Add(myComp);
                context.SaveChanges();
            }

            newComp = new List<Computer>();
            newComp = context.Computers.ToList();
            foreach (var list in newComp)
            {
                Console.WriteLine(list.ID);
                Console.WriteLine("UserName PC: " + list.UserNamePC);
                Console.WriteLine("Name PC: " + list.NamePC);
                Console.WriteLine("Operating System: " + list.OSVersion);
                Console.WriteLine("Bit Operating System: " + list.BitOperating);
                Console.WriteLine("Motherboard: " + list.Motherboard);
                Console.WriteLine("CPU: " + list.CPUpc);
                Console.WriteLine("RAM: " + list.RAMpc);
                Console.WriteLine("HDD: " + list.HDDpc);
                Console.WriteLine("Video: " + list.Video);
                Console.WriteLine("=====================================================");
            }
        }
        public static string CodingGetHash(string password)
        {
            using (var hash = SHA1.Create())
                return string.Concat(hash.ComputeHash(Encoding.UTF8.GetBytes(password)).Select(x => x.ToString("X2")));
        }
    }
}
