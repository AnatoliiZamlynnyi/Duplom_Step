using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;

namespace AICO_CL.Models
{
    public class Employe
    {
        public int ID { get; set; }
        public string Work { get; set; }
        public string Name { get; set; }
        public string? Password { get; set; }
        public string Phone { get; set; }
        public int? DepartmentID { get; set; }
        [ForeignKey("DepartmentID")]
        public Department Departments { get; set; }

        public override string ToString()
        {
            return $"{Name}";
        }
    }
}
