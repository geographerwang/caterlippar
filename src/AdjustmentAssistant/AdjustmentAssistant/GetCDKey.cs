using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management;

namespace AdjustmentAssistant
{
    class GetCDKey
    {
        public static string GetCpuID()
        {
            ManagementClass mc = new ManagementClass("Win32_Processor");
            ManagementObjectCollection moc = mc.GetInstances();
            String strCpuID = null;
            foreach (ManagementObject mo in moc)
            {
                strCpuID = mo.Properties["ProcessorId"].Value.ToString();
                break;
            }
            return strCpuID;
        }
    }
}
