using System.Runtime.InteropServices;
using System;
public static class RuntimeInfo
{
    public static string getEnvironmentInfo()
    {

        var environmentInfo =
        $"Date : {DateTime.Now.ToLongDateString()}  {DateTime.Now.ToLongTimeString()} \n" +
        $"Machine : { Environment.MachineName} \n" +
        $"OS : { RuntimeInformation.OSDescription.ToString()} \n" +
        $"Runtime ： { RuntimeInformation.FrameworkDescription.ToString()}";

        return environmentInfo;
    }
}
