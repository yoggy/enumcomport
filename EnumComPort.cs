using System;
using System.Management;
using System.Collections;
using System.Text.RegularExpressions;

namespace enumcomport
{
    class ComPortComparer : System.Collections.IComparer
    {
        Regex reg = new Regex("COM(?<num>\\d+)");

        public int Compare(object a, object b)
        {
            int a_num = int.Parse(reg.Match((string)a).Groups["num"].Value);
            int b_num = int.Parse(reg.Match((string)b).Groups["num"].Value);

            if (a_num == b_num)
            {
                return 0;
            }
            if (a_num < b_num)
            {
                return -1;
            }

            return 1;
        }
    }
    
    class EnumComPort
    {
        public static string[] getComPorts()
        {
            var list = new ArrayList();

            ManagementClass win32_pnpentity = new ManagementClass("Win32_PnPEntity");
            ManagementObjectCollection col = win32_pnpentity.GetInstances();

            Regex reg = new Regex(".+\\((?<port>COM\\d+)\\)");

            foreach (ManagementObject obj in col)
            {
                // name : "USB Serial Port(COM??)"
                string name = (string)obj.GetPropertyValue("name");
                if (name != null && name.Contains("(COM"))
                {
                    // "USB Serial Port(COM??)" -> COM??
                    Match m = reg.Match(name);
                    string port = m.Groups["port"].Value;

                    // description : "USB Serial Port"
                    string desc = (string)obj.GetPropertyValue("Description");

                    // result string : "COM?? (USB Serial Port)"
                    list.Add(port + " (" + desc + ")");
                }
            }

            ComPortComparer comp = new ComPortComparer();
            list.Sort(comp);

            return (string[])list.ToArray(typeof(string));
        }

        public static void Main()
        {
            string[] ports = getComPorts();
            foreach (string port in ports)
            {
                Console.WriteLine(port);
            }
        }
    }
}
