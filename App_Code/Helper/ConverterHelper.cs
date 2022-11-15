using System;


namespace SIS.Helper
{
    public abstract class ConverterHelper
    {
        public static string BytesToHumanReadable(long bytes)
        {
            if (bytes < 0)
                return bytes.ToString();

            string[] suffix = { "B", "KB", "MB", "GB", "TB", "PB", "EB" };
            if (bytes == 0)
                return "0 " + suffix[0];
            bytes = Math.Abs(bytes);
            int place = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
            double num = Math.Round(bytes / Math.Pow(1024, place), 1);
            return (Math.Sign(bytes) * num).ToString() + " " + suffix[place];
        }
    }
}