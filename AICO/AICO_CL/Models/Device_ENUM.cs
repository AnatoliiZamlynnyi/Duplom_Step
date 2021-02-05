using System;
using System.Collections.Generic;
using System.Text;

namespace AICO_CL.Models
{
    public class Device_ENUM
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public override string ToString()
        {
            return $"{Name}";
        }
    }
}