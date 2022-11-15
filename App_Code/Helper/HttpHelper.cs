using System;
using System.Collections.Generic;


namespace SIS.Helper
{
    public abstract class HttpHelper
    {
        public static readonly Dictionary<string, string> MimeTypeMap = new Dictionary<string, string>
        {
            // vgl. https://referencesource.microsoft.com/#system.web/MimeMapping.cs
            { ".zip",  "application/x-zip-compressed" },
            { ".xls",  "application/vnd.ms-excel" },
            { ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
            { ".ppt",  "application/vnd.ms-powerpoint" },
            { ".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
            { ".pps",  "application/vnd.ms-powerpoint" },
            { ".doc",  "application/msword" },
            { ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
            { ".txt",  "text/plain" },
            { ".gif",  "image/gif" },
            { ".jpg",  "image/jpeg" },
            { ".jpeg", "image/jpeg" },
            { ".png",  "image/png" },
            { ".pdf",  "application/pdf" },
            { ".vcf",  "text/x-vcard" }
        };

        public const string DefaultFallbackMimeType = "application/octet-stream";


        // TODO: Upgrade des .Net Frameworks auf 4.6.2, denn dann kann eine interne Methode verwendet werden, siehe https://referencesource.microsoft.com/#System.Web
        public static string GetMimeType(string filename, string fallbackMimeType = DefaultFallbackMimeType)
        {
            if (string.IsNullOrEmpty(filename))
                throw new ArgumentNullException("filename");

            string extension = System.IO.Path.GetExtension(filename);
            return GetMimeTypeByExtension(extension, fallbackMimeType);
        }

        // Die Dateierweiterung ist mit führendem Punkt anzugeben, also z.B. ".txt"
        public static string GetMimeTypeByExtension(string extension, string fallbackMimeType = DefaultFallbackMimeType)
        {
            if (string.IsNullOrEmpty(extension))
                throw new ArgumentNullException("extension");

            string mimeType;
            if (!MimeTypeMap.TryGetValue(extension.ToLower(), out mimeType))
                mimeType = fallbackMimeType;
            return mimeType;

        }
    }
}