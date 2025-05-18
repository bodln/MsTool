using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace mersid.Models
{
    public record XlsRecord(string OriginalKey, double Value, string Marker, int Flag);
    public record CsvRecord(string OriginalKey, double SumValue, string Date1, string Date2, string Position, string Pib, int Flag);
    public record DiffRecord
    {
        public string Marker { get; set; }
        public string OriginalKey { get; init; }
        public double XlsValue { get; init; }
        public double CsvSumValue { get; init; }
        public string CsvOriginalKey { get; init; }
        public string CsvDate1 { get; init; }
        public string CsvDate2 { get; init; }
        public string Position { get; set; }
        public string Pib { get; set; }
    }
}
