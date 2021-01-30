using AICO_CL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace AICO_CL.Entity
{
    public class Start
    {
        EfContext context;
        public void StartFirst()
        {
            context = new EfContext();
            context.Database.CreateIfNotExists();
           int count = context.Employes.Count();
            if (count == 0)
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
        }

        public static string CodingGetHash(string password)
        {
            using (var hash = SHA1.Create())
                return string.Concat(hash.ComputeHash(Encoding.UTF8.GetBytes(password)).Select(x => x.ToString("X2")));
        }
    }
}
