using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Management;

namespace macrosInstall
{

    class Program
    {

        static void Main(string[] args)
        {
            ManagementObjectSearcher searcher_soft =                    new ManagementObjectSearcher("root\\CIMV2",                       "SELECT * FROM Win32_Product");

            foreach (ManagementObject queryObj in searcher_soft.Get())
            {
                Console.WriteLine("<soft> Caption: {0} ; InstallDate: {1}</soft>",
                                  queryObj["Caption"], queryObj["InstallDate"]);
            }

            Console.Write("Press any key to continue . . . ");
            Console.ReadKey(true);
        }
    }
}
