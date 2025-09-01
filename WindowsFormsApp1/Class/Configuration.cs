using Autodesk.Revit.UI;
using System;
using System.IO;

namespace WindowsFormsApp1.Class
{
    public static class Configuration
    {
        public static string GetLogFilePath(string key)
        {
            try
            {
                string[] lines = File.ReadAllLines(Environment.ExpandEnvironmentVariables("%AppData%\\Autodesk\\Revit\\Addins\\2025\\NewPlagin\\config.txt"));
                foreach (var line in lines)
                {
                    if (line.StartsWith(key))
                    {
                        return line.Split('=')[1].Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Title", ex.Message);
            }
            return null;
        }
    }
}
