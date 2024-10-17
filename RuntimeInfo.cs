using System.Runtime.InteropServices;
using System;
using System.Reflection;
using System.Management;
public static class RuntimeInfo
{
    public static string getEnvironmentInfo()
    {

        string environmentInfo =
        $"Date    : {DateTime.Now.ToLongDateString()}  {DateTime.Now.ToLongTimeString()} \n" +
        $"Machine : {Environment.MachineName} \n" +
        $"Spec    : {getMachineInfo()} \n" +
        $"OS      : {RuntimeInformation.OSDescription.ToString()} \n" +
        $"Runtime : {RuntimeInformation.FrameworkDescription.ToString()} \n" +
        $"This    : {Assembly.GetExecutingAssembly().GetName().Name} \n" +
        $"Assembly: {getAssemblyName()}";

        Console.WriteLine(environmentInfo);
        return environmentInfo;
    }
    private static string getAssemblyName()
    {
        Type _type = typeof(GrapeCity.Documents.Excel.Workbook);
        string _assemblyName = "";
        _assemblyName += _type.Assembly.GetName().Name + " ";
        _assemblyName += _type.Assembly.GetName().Version;
        return _assemblyName;
    }
    private static string getMachineInfo()
    {
        string _machineInfo = "";
        if (System.Environment.OSVersion.Platform == PlatformID.Win32NT)
        {
            var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
            foreach (ManagementObject obj in searcher.Get())
            {
                string cpuName = obj["Name"].ToString();
                string core = obj["NumberOfCores"].ToString();
                string bit = obj["DataWidth"].ToString();
                _machineInfo += $"{cpuName} | {bit}bit| {core} コア |";
            }
            var memoryQuery = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            foreach (ManagementObject memory in memoryQuery.Get())
            {
                ulong totalPhysicalMemory = Convert.ToUInt64(memory["TotalPhysicalMemory"]);
                _machineInfo += $"メモリ:{totalPhysicalMemory / (1024 * 1024)} MB";
            }
        }
        return _machineInfo;

    }
}
