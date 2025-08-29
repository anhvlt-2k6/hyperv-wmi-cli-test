using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Resources;
using System.Security;
using System.Text.RegularExpressions;

namespace ConsoleApp1
{
    internal class Program
    {
        public static (string computerName, string username, string password) credUserInput;

        public static SecureString securePwd;
        public static ConnectionOptions credential;
        public static ManagementScope scope;
        public static ManagementObjectSearcher searcherOfPciExpressData, searcherOfVMList;

        public static readonly Dictionary<int, string> vmStatusMap = new Dictionary<int, string>
        {
            {0,  "Unknown" },
            {1,  "Other" },
            {2,  "Running" },
            {3,  "Stopped" },
            {4,  "Shutting down" },
            {5,  "Not applicable" },
            {6,  "Enabled but Offline" },
            {7,  "In Test" },
            {8,  "Degraded" },
            {9,  "Quiesce" },
            {10, "Starting" }
        };


        public static Dictionary<string, (string deviceName, Double deviceMem, string deviceNote)> devices;
        //public static Dictionary<string, (string vmName, string vmStatus, List<string> devices)> osObjects;
        public static string scopePath;
        public static bool isLocal;

        public static void ConnectionEstablishment()
        {
            securePwd = new SecureString();
            foreach (char c in credUserInput.password)
            {
                securePwd.AppendChar(c);
            }

            credential = new ConnectionOptions
            {
                Username = credUserInput.username,
                SecurePassword = securePwd,
                Impersonation = ImpersonationLevel.Impersonate,
                Authentication = AuthenticationLevel.Default
            };

            scopePath = $"\\\\{credUserInput.computerName}\\root\\cimv2";
            scope = isLocal ? new ManagementScope(scopePath) : new ManagementScope(scopePath, credential);
            //searcherOfPciExpressData = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT * FROM Msvm_PciExpressSettingData"));
            //osObjects = new Dictionary<string, (string vmName, string vmStatus, List<string> devices)>();
            devices = new Dictionary<string, (string deviceName, Double deviceMem, string deviceNote)>();
            scope.Connect();

            /*
             * query types
             */
            ManagementObjectSearcher searcherofDeviceInstancePath = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT * FROM Win32_PnPEntity"));
            ManagementObjectSearcher searcherofResourcesAllocate = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT * FROM Win32_PNPAllocatedResource"));
            ManagementObjectSearcher searcherofDeviceMemAddress = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT * FROM Win32_DeviceMemoryAddress"));
            
            /*
             * perform adding data
             */
            foreach (var device in searcherofDeviceInstancePath.Get())
            {
                if (device["DeviceID"].ToString().StartsWith("PCI\\")) {

                    string deviceId = device["DeviceID"].ToString();
                    string startingaddress = String.Empty;
                    Double memoryGap = 0;

                    foreach (var resource in searcherofResourcesAllocate.Get())
                    {
                        if (resource["Dependent"].ToString().Contains(deviceId.Replace("\\", "\\\\")))
                        {
                            startingaddress = ((string[])resource["Antecedent"].ToString().Split('='))[1].Replace("\"", "");
                            break;
                        }
                    }

                    foreach (var deviceMemAddress in searcherofDeviceMemAddress.Get())
                    {
                        if (deviceMemAddress["StartingAddress"].ToString().Equals(startingaddress))
                        {
                            string[] addressRange = deviceMemAddress["Name"].ToString().Replace("0x", "").Split('-');
                            ulong startRange = ulong.Parse(addressRange[0], System.Globalization.NumberStyles.HexNumber);
                            ulong endRange = ulong.Parse(addressRange[1], System.Globalization.NumberStyles.HexNumber);
                            memoryGap += Math.Floor((endRange - startRange + 1) / (1048576.0));
                        }
                    }

                    devices[deviceId] = (device["Caption"].ToString(), memoryGap, String.Empty);
                    
                }
            }

            /*
             * display
             */
            foreach (var device in devices)
            {
                Console.WriteLine($"Device Instance ID: {device.Key}");
                Console.WriteLine($"\t...with name {device.Value.deviceName}");
                Console.WriteLine($"\t...with memory gap {device.Value.deviceMem} Mb\n");
            }
            
        }

        public static void Main(string[] args)
        {
            Console.WriteLine("Enter a computer name: ");
            credUserInput.computerName = Console.ReadLine();

            if (credUserInput.computerName.ToLower().Contains("localhost"))
            {
                isLocal = true;
            }

            if (!isLocal)
            {
                Console.WriteLine("Enter a username: ");
                credUserInput.username = Console.ReadLine()?.ToLower();

                Console.WriteLine("Enter a password: ");
                credUserInput.password = Console.ReadLine();
            }
            else
            {
                credUserInput.username = string.Empty;
                credUserInput.password = string.Empty;
            }

            ConnectionEstablishment();
        }
    }
}
