using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using ReadingExampleCollector;
using System.Text;
using GetRequest = Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.GetRequest;
using UpdateRequest = Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.UpdateRequest;

namespace GoogleSheetsExample {
    public class Program {
        static void Main(string[] args) {
            Console.InputEncoding = Encoding.Unicode;
            Console.OutputEncoding = Encoding.Unicode;

            List<ColumnSpan> wordRanges = new List<ColumnSpan> { new ColumnSpan("B2:B1800"), new ColumnSpan("J2:J1800") };
            List<ColumnSpan> kanjiRanges = new List<ColumnSpan> { new ColumnSpan("B3:B1300") };

            string credentialsFileName = "japanese-417119-e94d80bf56c9.json";
            string spreadsheetId = "1pJCzatA15yrxHMTlYcoxynaDEpdlMR07NNrmN462Ays";
            string wordSheetName = "単語";
            string readingSheetName = "読";

            string credentialsFilePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"..\\..\\..\\{credentialsFileName}"));
            SheetsService service = CreateGoogleSheetsService(credentialsFilePath);

            List<Cell> words = FetchWords(service, spreadsheetId, wordSheetName, wordRanges);

            List<IList<object>> newValues = new List<IList<object>>();
            newValues.Add(new List<object> { "Test" });
            newValues.Add(new List<object> { "Test2" });

            ValueRange valueRange = new ValueRange();
            valueRange.Values = newValues;

            UpdateRequest updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, $"{readingSheetName}!G3:G4");
            updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED; //?

            UpdateValuesResponse response2 = updateRequest.Execute();
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

        private static List<Cell> FetchWords(SheetsService service, string spreadsheetId, string wordSheetName, List<ColumnSpan> wordRanges) {
            List<Cell> words = new List<Cell>();

            foreach (ColumnSpan span in wordRanges) {
                string reqRange = $"{wordSheetName}!{span}";

                GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, reqRange);
                ValueRange response = request.Execute();

                for (int i = 0; i < response.Values.Count; i++) {
                    var value = response.Values[i].FirstOrDefault();
                    if (value != null && value.ToString() != "-") {
                        words.Add(new Cell(span.Column, span.FromRow + i, value.ToString()!));
                    }
                }
            }

            return words.DistinctBy(cell => cell.Content).ToList();
        }
    }
}