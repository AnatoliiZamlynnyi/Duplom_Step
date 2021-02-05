using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;

namespace AICO_CL.Models
{
    public class Device
    {
        public int ID { get; set; }
        public string Model { get; set; }
        public string Description_1 { get; set; }
        public string Description_2 { get; set; }
        public string Description_3 { get; set; }
        public string Description_4 { get; set; }
        public string Description_5 { get; set; }
        public int Device_ENUM_ID { get; set; }
        [ForeignKey("Device_ENUM_ID")]
        public Device_ENUM Devices_ENUM { get; set; }
        public override string ToString()
        {
            return $"{Model}";
        }
    }
}