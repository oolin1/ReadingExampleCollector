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
        private static List<ColumnSpan> wordRanges = new List<ColumnSpan> { new ColumnSpan("B2:B1800"), new ColumnSpan("J2:J1800") };
        private static ColumnSpan kanjiRange = new ColumnSpan("B3:B1300");

        private static string credentialsFileName = "japanese-417119-e94d80bf56c9.json";
        private static string spreadsheetId = "1pJCzatA15yrxHMTlYcoxynaDEpdlMR07NNrmN462Ays";
        private static string wordSheetName = "単語";
        private static string readingSheetName = "読";
        private static string exampleColumn = "G";

        public static void Main(string[] args) {
            Console.InputEncoding = Encoding.Unicode;
            Console.OutputEncoding = Encoding.Unicode;

            string credentialsFilePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"..\\..\\..\\{credentialsFileName}"));
            SheetsService service = CreateGoogleSheetsService(credentialsFilePath);

            List<Cell> words = FetchCells(service, spreadsheetId, wordSheetName, wordRanges);
            List<Cell> kanjis = FetchCells(service, spreadsheetId, readingSheetName, new List<ColumnSpan> { kanjiRange });


            var newValues = new List<IList<object>>();

            foreach (Cell cell in kanjis) {
                List<string> containingWords = words.Where(word => word.Content.Contains(cell.Content)).Select(word => word.Content).ToList();

                if (containingWords.Count > 1) {
                    string result = containingWords.Aggregate((current, next) => current + ", " + next);
                    newValues.Add(new List<object> { result });
                }
                else if (containingWords.Count == 1) {
                    newValues.Add(new List<object> { containingWords.FirstOrDefault()! });
                }
                else {
                    newValues.Add(new List<object> { "" });
                }
            }
            ValueRange exampleValues = new ValueRange { Values = newValues };
            string exampleTargetRange = $"{readingSheetName}!{exampleColumn}{kanjiRange.FromRow}:{exampleColumn}{kanjiRange.ToRow}";

            UpdateRequest updateRequest = service.Spreadsheets.Values.Update(exampleValues, spreadsheetId, exampleTargetRange);
            updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;
            updateRequest.Execute();
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

        private static List<Cell> FetchCells(SheetsService service, string sheetId, string sheetName, List<ColumnSpan> spans) {
            List<Cell> words = new List<Cell>();

            foreach (ColumnSpan span in spans) {
                string reqRange = $"{sheetName}!{span}";

                GetRequest request = service.Spreadsheets.Values.Get(sheetId, reqRange);
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