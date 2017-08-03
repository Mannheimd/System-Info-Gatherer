using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Windows;
using WUApiInterop;

namespace System_Info_Gatherer
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public void RunGatherSpecs_Button_Click(object sender, RoutedEventArgs e)
        {
            runSpecGather();
        }

        public void runSpecGather()
        {
            string userDesktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            using (StreamWriter writer = new StreamWriter(userDesktop + @"\System Information Report.txt"))
            {
                GetWindowsVersion(writer);
                GetSystemInformation(writer);
                GetOfficeVersions(writer);
                GetDotNetVersions(writer);
                GetWindowsUpdates(writer);

                MessageBox.Show("Results exported to 'System Information Report.txt', which is on your desktop.");
            }
        }

        [DllImport("kernel32.dll")]
        static extern bool GetBinaryType(string lpApplicationName, out BinaryType lpBinaryType);

        public string HKLM_GetString(string path, string key)
        {
            try
            {
                RegistryKey regKey = Registry.LocalMachine.OpenSubKey(path);
                if (regKey == null) return "";
                return (string)regKey.GetValue(key);
            }
            catch { return ""; }
        }

        public void GetWindowsVersion(StreamWriter writer)
        {
            writer.WriteLine("Windows Product: " + HKLM_GetString(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName"));
            writer.WriteLine("Windows Service Pack: " + HKLM_GetString(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CSDVersion"));
            writer.WriteLine("Windows Build: " + HKLM_GetString(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuild"));
            writer.WriteLine();
        }

        public void GetSystemInformation(StreamWriter writer)
        {
            writer.WriteLine("- System Information");

            //Win32_ComputerSystem
            ManagementObjectSearcher system_Searcher = new ManagementObjectSearcher("SELECT Name, Manufacturer, Model, SystemType FROM Win32_ComputerSystem");
            ManagementObjectCollection system_Collection = system_Searcher.Get();
            foreach (ManagementObject system_Object in system_Collection)
            {
                writer.WriteLine("System Name: " + (system_Object["Name"] != null ? system_Object["Name"].ToString() : ""));
                writer.WriteLine("System Manufacturer: " + (system_Object["Manufacturer"] != null ? system_Object["Manufacturer"].ToString() : ""));
                writer.WriteLine("System Model: " + (system_Object["Model"] != null ? system_Object["Model"].ToString() : ""));
                writer.WriteLine("System Type: " + (system_Object["SystemType"] != null ? system_Object["SystemType"].ToString() : ""));
            }

            //Win32_Processor
            ManagementObjectSearcher processor_Searcher = new ManagementObjectSearcher("SELECT Name, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors FROM Win32_Processor");
            ManagementObjectCollection processor_Collection = processor_Searcher.Get();
            foreach (ManagementObject processor_Object in processor_Collection)
            {
                writer.WriteLine("Processor Name: " + (processor_Object["Name"] != null ? processor_Object["Name"].ToString() : ""));
                writer.WriteLine("Max Clock Speed: " + (processor_Object["MaxClockSpeed"] != null ? processor_Object["MaxClockSpeed"].ToString() : ""));
                writer.WriteLine("Number Of Cores: " + (processor_Object["NumberOfCores"] != null ? processor_Object["NumberOfCores"].ToString() : ""));
                writer.WriteLine("Number Of Logical Processors: " + (processor_Object["NumberOfLogicalProcessors"] != null ? processor_Object["NumberOfLogicalProcessors"].ToString() : ""));
            }

            //Win32_OperatingSystem
            ManagementObjectSearcher memory_Searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
            ManagementObjectCollection memory_Collection = memory_Searcher.Get();
            foreach (ManagementObject memory_Object in memory_Searcher.Get())
            {
                writer.WriteLine("Physical Memory: " + (memory_Object["TotalVisibleMemorySize"] != null ? memory_Object["TotalVisibleMemorySize"].ToString() : ""));
                writer.WriteLine("Virtual Memory: " + (memory_Object["TotalVirtualMemorySize"] != null ? memory_Object["TotalVirtualMemorySize"].ToString() : ""));
            }

            //Locale
            writer.WriteLine("Installed Culture: " + CultureInfo.InstalledUICulture.DisplayName);
            writer.WriteLine("Current Culture: " + CultureInfo.CurrentUICulture.DisplayName);

            //Time Zone
            writer.WriteLine("Time Zone: " + (TimeZone.CurrentTimeZone.IsDaylightSavingTime(DateTime.Now) ? TimeZone.CurrentTimeZone.DaylightName : TimeZone.CurrentTimeZone.StandardName));

            writer.WriteLine();
        }

        public void GetOfficeVersions(StreamWriter writer)
        {
            writer.WriteLine("- Installed Office Versions:");

            RegistryKey localMachine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            RegistryKey software32BitKey = localMachine.OpenSubKey(@"Software\Microsoft\Office", false);
            RegistryKey software64BitKey = localMachine.OpenSubKey(@"Software\WOW6432Node\Microsoft\Office", false);

            foreach (KeyValuePair<string, string> officeVersion in OfficeVersions.versionNumbers)
            {
                try
                {
                    RegistryKey versionKey = software64BitKey.OpenSubKey(officeVersion.Key, false);
                    if (versionKey != null)
                    {
                        writer.WriteLine(officeVersion.Value);
                        writer.WriteLine("Reg Path: " + versionKey.ToString());

                        string[] versionSubKeys = versionKey.GetSubKeyNames();
                        bool applicationFound = false;
                        foreach (string versionSubKey in versionSubKeys)
                        {
                            if (OfficeVersions.productNames.ContainsKey(versionSubKey))
                            {
                                applicationFound = true;
                                writer.WriteLine("Application: " + versionSubKey);

                                string exeName = OfficeVersions.productNames[versionSubKey];
                                string installRootPath = null;
                                try
                                {
                                    RegistryKey productInstallRoot = versionKey.OpenSubKey(versionSubKey + @"\InstallRoot");
                                    installRootPath = productInstallRoot.GetValue("Path").ToString() + exeName;

                                    writer.WriteLine("Location: " + installRootPath);
                                }
                                catch
                                {
                                    writer.WriteLine("Location: " + "Registry reference found, but no install path");
                                }

                                if (installRootPath != null)
                                {
                                    try
                                    {
                                        GetBinaryType(installRootPath, out BinaryType binType);
                                        writer.WriteLine("Binary Type: " + binType);

                                        writer.WriteLine();
                                    }
                                    catch
                                    {
                                        writer.WriteLine("Binary Type: " + "Type not found");
                                    }
                                }
                            }
                        }

                        if (applicationFound == false)
                        {
                            writer.WriteLine("No applications found for this version");
                        }

                        writer.WriteLine();
                    }
                }
                catch { }

                try
                {
                    RegistryKey versionKey = software32BitKey.OpenSubKey(officeVersion.Key, false);
                    if (versionKey != null)
                    {
                        writer.WriteLine(officeVersion.Value);
                        writer.WriteLine("Reg Path: " + versionKey.ToString());

                        string[] versionSubKeys = versionKey.GetSubKeyNames();
                        bool applicationFound = false;
                        foreach (string versionSubKey in versionSubKeys)
                        {
                            if (OfficeVersions.productNames.ContainsKey(versionSubKey))
                            {
                                applicationFound = true;
                                writer.WriteLine("Application: " + versionSubKey);

                                string exeName = OfficeVersions.productNames[versionSubKey];
                                string installRootPath = null;
                                try
                                {
                                    RegistryKey productInstallRoot = versionKey.OpenSubKey(versionSubKey + @"\InstallRoot");
                                    installRootPath = productInstallRoot.GetValue("Path").ToString() + exeName;

                                    writer.WriteLine("Location: " + installRootPath);
                                }
                                catch
                                {
                                    writer.WriteLine("Location: " + "Registry reference found, but no install path");
                                }

                                if (installRootPath != null)
                                {
                                    try
                                    {
                                        GetBinaryType(installRootPath, out BinaryType binType);
                                        writer.WriteLine("Binary Type: " + binType);

                                        writer.WriteLine();
                                    }
                                    catch
                                    {
                                        writer.WriteLine("Binary Type: " + "Type not found");
                                    }
                                }
                            }
                        }

                        if (applicationFound == false)
                        {
                            writer.WriteLine("No applications found for this version");
                        }

                        writer.WriteLine();
                    }
                }
                catch { }
            }
        }

        public void GetDotNetVersions(StreamWriter writer)
        {
            writer.WriteLine("- Installed .NET Versions:");

            using (RegistryKey ndpKey = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, "").OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\"))
            {
                foreach (string versionKeyName in ndpKey.GetSubKeyNames())
                {
                    if (versionKeyName.StartsWith("v"))
                    {
                        RegistryKey versionKey = ndpKey.OpenSubKey(versionKeyName);
                        string fullVersion = (string)versionKey.GetValue("Version", "");
                        string servicePack = versionKey.GetValue("SP", "").ToString();
                        string installState = versionKey.GetValue("Install", "").ToString();
                        if (installState == "1")
                        {
                            writer.WriteLine("Version: " + fullVersion);
                            writer.WriteLine("Service pack:" + (servicePack != null ? servicePack : ""));
                            writer.WriteLine();
                        }
                    }
                }

                using (RegistryKey v4FullKey = ndpKey.OpenSubKey(@"v4\Full\"))
                {
                    string fullVersion = (string)v4FullKey.GetValue("Version", "");
                    string releaseNumber = v4FullKey.GetValue("Release", "").ToString();
                    string installState = v4FullKey.GetValue("Install", "").ToString();

                    if (installState == "1")
                    {
                        writer.WriteLine("Version: Full " + fullVersion);
                        writer.WriteLine("Release: " + releaseNumber);
                        writer.WriteLine();
                    }
                }

                using (RegistryKey v4ClientKey = ndpKey.OpenSubKey(@"v4\Client\"))
                {
                    string fullVersion = (string)v4ClientKey.GetValue("Version", "");
                    string releaseNumber = v4ClientKey.GetValue("Release", "").ToString();
                    string installState = v4ClientKey.GetValue("Install", "").ToString();

                    if (installState == "1")
                    {
                        writer.WriteLine("Version: Client " + fullVersion);
                        writer.WriteLine("Release: " + releaseNumber);
                        writer.WriteLine();
                    }
                }
            }
        }

        public void GetWindowsUpdates(StreamWriter writer)
        {
            writer.WriteLine("- Windows Update History:");

            UpdateSession updateSession = new UpdateSession();
            IUpdateSearcher updateSearcher = updateSession.CreateUpdateSearcher();
            IUpdateHistoryEntryCollection updateHistoryCollection = updateSearcher.QueryHistory(0, updateSearcher.GetTotalHistoryCount());
            foreach (IUpdateHistoryEntry updateHistory in updateHistoryCollection)
            {
                if (updateHistory.Date == new DateTime(1899, 12, 30))
                {
                    continue;
                }

                writer.WriteLine("Date: " + updateHistory.Date.ToLongDateString() + " " + updateHistory.Date.ToLongTimeString());
                writer.WriteLine("Title: " + (updateHistory.Title != null ? updateHistory.Title : ""));
                writer.WriteLine("Description: " + (updateHistory.Description != null ? updateHistory.Description : ""));
                writer.WriteLine("Operation: " + updateHistory.Operation);
                writer.WriteLine("Result Code: " + updateHistory.ResultCode);
                writer.WriteLine("Support URL: " + (updateHistory.SupportUrl != null ? updateHistory.SupportUrl : ""));
                writer.WriteLine("Update ID: " + (updateHistory.UpdateIdentity.UpdateID != null ? updateHistory.UpdateIdentity.UpdateID : ""));
                writer.WriteLine("Update Revision: " + (updateHistory.UpdateIdentity.RevisionNumber.ToString() != null ? updateHistory.UpdateIdentity.RevisionNumber.ToString() : ""));

                writer.WriteLine();
            }
        }
    }

    public class OfficeVersions
    {
        public static Dictionary<string, string> versionNumbers = new Dictionary<string, string>
        {
            { "1.0", "Office 1.0" },
            { "1.5", "Office 1.5" },
            { "1.6", "Office 1.6" },
            { "3.0", "Office 3.0" },
            { "4.0", "Office 4.0" },
            { "4.2", "Office for NT 4.2" },
            { "4.3", "Office 4.3" },
            { "7.0", "Office 95" },
            { "8.0", "Office 97" },
            { "8.5", "Office 97 Powered by Word 98" },
            { "9.0", "Office 2000" },
            { "10.0", "Office XP" },
            { "11.0", "Office 2003" },
            { "12.0", "Office 2007" },
            { "14.0", "Office 2010" },
            { "15.0", "Office 2013" },
            { "16.0", "Office 2016" }
        };

        public static Dictionary<string, string> productNames = new Dictionary<string, string>
        {
            { "Word", "winword.exe" },
            { "Excel", "excel.exe" },
            { "Outlook", "outlook.exe" },
            { "PowerPoint", "powerpnt.exe" },
            { "Access", "msaccess.exe" },
            { "Visio", "visio.exe" },
            { "Project", "winproj.exe" },
            { "Publisher", "mspub.exe" },
            { "FrontPage", "frontpg.exe" },
            { "Schedule+", "schdpl32.exe" },
            { "PhotoDraw", "photodrw.exe" },
            { "Binder", "binder.exe" },
            { "Photo Editor", "photoed.exe" },
            { "MapPoint Deluxe", "mappoint.exe" },
            { "InfoPath", "infopath.exe" },
            { "OneNote", "onenote.exe" },
            { "Communicator", "communicator.exe" },
            { "Groove/SharePoint Workspace/OneDrive", "groove.exe" },
            { "SharePoint Designer", "spdesign.exe" },
        };
    }

    public enum BinaryType : uint
    {
        SCS_32BIT_BINARY = 0, // A 32-bit Windows-based application
        SCS_64BIT_BINARY = 6, // A 64-bit Windows-based application.
        SCS_DOS_BINARY = 1, // An MS-DOS – based application
        SCS_OS216_BINARY = 5, // A 16-bit OS/2-based application
        SCS_PIF_BINARY = 3, // A PIF file that executes an MS-DOS – based application
        SCS_POSIX_BINARY = 4, // A POSIX – based application
        SCS_WOW_BINARY = 2 // A 16-bit Windows-based application 
    }
}
