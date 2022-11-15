using System;
using System.IO;


namespace SIS.Helper
{
    public abstract class FileHelper
    {
        // Gibt die Dateierweiterung ohne Punkt zurück
        public static string GetPureExtension(string filename)
        {
            return Path.GetExtension(filename).Remove(0, 1);
        }

        public static string GetUniqueFilename(string uploadPath, string fileExtension = "")
        {
            string uniqueFilename;
            do
            {
                uniqueFilename = Guid.NewGuid().ToString();
                if (fileExtension != "")
                {
                    uniqueFilename += "." + fileExtension;
                }
            } while (File.Exists(uploadPath + uniqueFilename));
            return uniqueFilename;
        }

        public static string GetBase64EncodedFile(string filename)
        {
            try
            {
                FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
                byte[] filebytes;
                try
                {
                    filebytes = new byte[fs.Length];
                    fs.Read(filebytes, 0, Convert.ToInt32(fs.Length));
                }
                finally
                {
                    fs.Close();
                }

                string base64File = Convert.ToBase64String(filebytes, Base64FormattingOptions.None);
                return base64File;
            }
            catch
            {
                return "";
            }
        }

        public static long GetFileSize(string filename)
        {
            FileInfo fileinfo = new FileInfo(filename);
            if (fileinfo.Exists)
                return fileinfo.Length;
            else
                return -1;
        }

        public static string GetFileSizeAsString(string filename, string invalidFileSizeString = "n/a")
        {
            long size = GetFileSize(filename);
            if (size == -1)
                return invalidFileSizeString;
            else
                return ConverterHelper.BytesToHumanReadable(size);
        }

        public static bool FileExists(string filename)
        {
            FileInfo fileInfo = new FileInfo(filename);
            return fileInfo.Exists;
        }
    }
}