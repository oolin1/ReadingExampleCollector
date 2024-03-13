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