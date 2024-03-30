namespace ReadingExampleCollector {
    public class ColumnSpan {
        public string Column { get; }
        public int FromRow { get; }
        public int ToRow { get; }

        public ColumnSpan(string range) {
            string[] parts = range.Split(':');

            (string, int) from = ParseCellReference(parts[0]);
            (string, int) to = ParseCellReference(parts[1]);

            this.Column = from.Item1;
            this.FromRow = from.Item2;
            this.ToRow = to.Item2;
        }

        private (string, int) ParseCellReference(string cellReference) {
            string column = "";
            string row = "";

            foreach (char character in cellReference) {
                if (char.IsLetter(character)) {
                    column += character;
                }
                else {
                    row += character;
                }
            }

            return new(column, Int32.Parse(row));
        }

        public override string ToString() {
            return $"{Column}{FromRow}:{Column}{ToRow}";
        }
    }
}