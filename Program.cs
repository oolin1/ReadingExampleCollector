﻿using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Text;

namespace GoogleSheetsExample {
    class Program {
        static void Main(string[] args) {
            Console.InputEncoding = Encoding.Unicode;
            Console.OutputEncoding = Encoding.Unicode;

            string credentialsFileName = "japanese-417119-e94d80bf56c9.json";
            string spreadsheetId = "1pJCzatA15yrxHMTlYcoxynaDEpdlMR07NNrmN462Ays";
            string wordSheetName = "単語";
            string readingSheetName = "読";

            string credentialsFilePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"..\\..\\..\\{credentialsFileName}"));

            SheetsService service = CreateGoogleSheetsService(credentialsFilePath);

            string range = $"{wordSheetName}!B3:B5";
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            ValueRange response = request.Execute();
        }

        private static SheetsService CreateGoogleSheetsService(string credPath) {
            GoogleCredential credential;

            using (var stream = new FileStream(credPath, FileMode.Open, FileAccess.Read)) {
                credential = GoogleCredential.FromStream(stream).CreateScoped(SheetsService.Scope.Spreadsheets);
            }

            var service = new SheetsService(new BaseClientService.Initializer() {
                HttpClientInitializer = credential,
                ApplicationName = AppDomain.CurrentDomain.FriendlyName
            });

            return service;
        }
    }
}