using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text;

namespace AICO_CL.Models
{
    public class Computer
    {
        public int ID { get; set; }
        public string UserNamePC { get; set; }
        public string NamePC { get; set; }
        public string OSVersion { get; set; }
        public string BitOperating { get; set; }
        public string Motherboard { get; set; }
        public string CPUpc { get; set; }
        public string RAMpc { get; set; }
        public string HDDpc { get; set; }
        public string Video { get; set; }
        public override string ToString()
        {
            return $"{NamePC}";
        }
        public void InfoSystem()
        {
            UserNamePC = Environment.UserName;
            NamePC = Environment.MachineName;
            var os = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem")
                .Get()
                .Cast<ManagementObject>()
                .First();
            OSVersion = (string)os["Caption"];
            if (Environment.Is64BitOperatingSystem == true)
                BitOperating = "x64";
            else
                BitOperating = "x86";
            var motherboard = new ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard")
                .Get()
                .Cast<ManagementObject>()
                .First();
            Motherboard = (string)motherboard["Manufacturer"] + " " + (string)motherboard["Product"];
            var cpu = new ManagementObjectSearcher("SELECT * FROM Win32_Processor")
               .Get()
               .Cast<ManagementObject>()
               .First();
            CPUpc = (string)cpu["Name"];
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
            ManagementObjectCollection results = searcher.Get();
            foreach (ManagementObject result in results)
                RAMpc = Convert.ToString(Math.Round(Convert.ToDecimal(result["TotalVisibleMemorySize"]) / (1024 * 1024))) + " Gb.";
            var hdd = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive")
                .Get()
                .Cast<ManagementObject>()
                .First();
            int size = (int)(Convert.ToDecimal(hdd["Size"]) / (1024 * 1024 * 1000));
            HDDpc = (string)hdd["Model"] + " - " + size.ToString() + " Gb.";
            var video = new ManagementObjectSearcher("SELECT * FROM Win32_VideoController")
           .Get()
           .Cast<ManagementObject>()
           .First();
            Video = (string)video["Name"];
        }

        public void ShowPC()
        {
            Console.WriteLine("My Computer");
            Console.WriteLine("UserName PC: " + UserNamePC);
            Console.WriteLine("Name PC: " + NamePC);
            Console.WriteLine("Operating System: " + OSVersion);
            Console.WriteLine("Bit Operating System: " + BitOperating);
            Console.WriteLine("Motherboard: " + Motherboard);
            Console.WriteLine("CPU: " + CPUpc);
            Console.WriteLine("RAM: " + RAMpc);
            Console.WriteLine("HDD: " + HDDpc);
            Console.WriteLine("Video: " + Video);
            Console.WriteLine("====================================================");
        }
    }
}