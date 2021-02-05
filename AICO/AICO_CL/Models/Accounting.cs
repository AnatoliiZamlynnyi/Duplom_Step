using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;

namespace AICO_CL.Models
{
    public class Accounting
    {
        public int ID { get; set; }
        public int? EmployeID { get; set; }
        [ForeignKey("EmployeID")]
        public Employe Employes { get; set; }
        public int? ComputerID { get; set; }
        [ForeignKey("ComputerID")]
        public Computer Computers { get; set; }
        public int? DeviceID { get; set; }
        [ForeignKey("DeviceID")]
        public Device Devices { get; set; }
    }
}