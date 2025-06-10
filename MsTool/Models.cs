using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsTool.Models
{
    public record XlsRecord(
        string OriginalKey, 
        double Value, 
        string Date, // DATDOK
        string Marker, // VR
        int Flag // 1/2/3
    );

    public record CsvRecord(
        string OriginalKey, 
        double SumValue, 
        string Date1, // Datum PDV obaveze/evidentiranja
        string Date2, // Datum obrade
        string Position, 
        string Pib, 
        int Flag,
        string Status
    );

    public record DiffRecord
    {
        public string XlsMarker { get; set; }
        public string XlsOriginalKey { get; init; }
        public double XlsValue { get; init; }
        public double CsvSumValue { get; init; }
        public string CsvOriginalKey { get; init; }
        public string CsvDate1 { get; init; }
        public string CsvDate2 { get; init; }
        public string Position { get; set; }
        public string Pib { get; set; }
        public string CompanyName { get; set; }
        public bool DoubleTake { get; set; }
        public string Status { get; set; }
    }

    public record XlsAnalyticsRecord(
        string OriginalKey,
        double ValueMain, // DUGUJE
        double ValueRef, // POTRAZUJE
        string Date, // DATUM
        string Account, // NALOG
        bool Flag 
    );

    public record DiffAnalyticsRecord
    {
        public string OriginalMainKey { get; set; }
        public string OriginalRefKey { get; set; }
        public double ValueMain { get; set; }
        public double ValueRef { get; set; }
        public string DateMain { get; set; }
        public string DateRef { get; set; }
        public string AccountMain { get; set; }
        public string AccountRef { get; set; }
        public bool DoubleTake { get; set; } = false;
    }
}
