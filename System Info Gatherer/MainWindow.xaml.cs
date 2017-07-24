﻿using Microsoft.Win32;
using System;
using System.Globalization;
using System.IO;
using System.Management;
using System.Windows;

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
                GetDotNetVersions(writer);
                GetWindowsUpdates(writer);

                MessageBox.Show("Results exported to 'System Information Report.txt', which is on your desktop.");
            }
        }

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
            writer.WriteLine("- Installed Windows Updates:");

            string query = "SELECT Caption, Description, HotFixID, InstalledBy, InstalledOn FROM Win32_QuickFixEngineering";
            ManagementObjectCollection resultCollection = null;
            try
            {
                ManagementObjectSearcher search = new ManagementObjectSearcher(query);
                resultCollection = search.Get();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }

            foreach (ManagementObject result in resultCollection)
            {
                try
                {
                    writer.WriteLine("Caption: " + (result["Caption"] != null ? result["Caption"].ToString() : ""));
                    writer.WriteLine("Description: " + (result["Description"] != null ? result["Description"].ToString() : ""));
                    writer.WriteLine("HotFixID: " + (result["HotFixID"] != null ? result["HotFixID"].ToString() : ""));
                    writer.WriteLine("InstalledBy: " + (result["InstalledBy"] != null ? result["InstalledBy"].ToString() : ""));
                    writer.WriteLine("InstalledOn: " + (result["InstalledOn"] != null ? result["InstalledOn"].ToString() : ""));
                    writer.WriteLine();
                }
                catch(Exception error)
                {
                    MessageBox.Show(error.Message);
                }
            }
        }
    }
}
