using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace GoogleSheetsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            string credentialsFileName = "japanese-417119-e94d80bf56c9.json";
            string credentialsFilePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"..\\..\\..\\{credentialsFileName}"));
            File.Exists(credentialsFilePath);
        }
    }
}