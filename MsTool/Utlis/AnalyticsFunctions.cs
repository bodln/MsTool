using MathNet.Numerics;
using MsTool.Models;
using NPOI.OpenXmlFormats.Vml;
using SixLabors.ImageSharp.ColorSpaces;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsTool.Utlis
{
    public static class AnalyticsFunctions
    {
        public static void Proceed(string xlsMainPath, string xlsRefPath, bool assumptions)
        {
            if (string.IsNullOrEmpty(xlsMainPath) || string.IsNullOrEmpty(xlsRefPath))
            {
                MessageBox.Show("Molim vas, prvo izaberite oba fajla.");
                return;
            }

            try
            {
                var xlsMainRecs = FileManipulator.LoadXlsAnalytics(xlsMainPath); // what csv was for BoB
                var xlsRefRecs = FileManipulator.LoadXlsAnalytics(xlsRefPath); 

                List<DiffAnalyticsRecord> diffs = new List<DiffAnalyticsRecord>();

                foreach (var key in xlsMainRecs.Keys)
                {
                    var xlsMain = xlsMainRecs[key];
                    xlsRefRecs.TryGetValue(key, out var xlsRef);

                    double refVal = xlsRef?.ValueRef ?? 0;

                    bool equal = Math.Abs(refVal - xlsMain.ValueMain) <= 5.0;
                    bool doubleTake = false;

                    if (!equal)
                    {
                        var matchingKey = xlsRefRecs
                            .Where(kvp =>
                            {
                                bool valueMatch = Math.Abs(kvp.Value.ValueRef - xlsMain.ValueMain) <= 5.0;

                                if (!valueMatch)
                                    return false;

                                if (!DateTime.TryParseExact(kvp.Value.Date, "dd-MM-yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var xlsRelativeDate))
                                    return false;

                                if (!DateTime.TryParseExact(xlsMain.Date, "dd-MM-yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var xlsMainDate))
                                    return false;

                                return xlsRelativeDate.Date == xlsMainDate.Date;
                            })
                            .Select(kvp => kvp.Key)
                            .FirstOrDefault();

                        if (matchingKey != null)
                        {
                            doubleTake = true;
                            xlsRef = xlsRefRecs[matchingKey];
                        }
                    }

                    if (xlsRef == null || !equal)
                    {
                        diffs.Add(new DiffAnalyticsRecord
                        {
                            OriginalMainKey = xlsMain.OriginalKey,
                            OriginalRefKey = xlsRef?.OriginalKey ?? "Nema",
                            ValueMain = xlsMain.ValueMain,
                            ValueRef = refVal,
                            DateMain = xlsMain.Date,
                            DateRef = xlsRef?.Date ?? "",
                            AccountMain = xlsMain.Account,
                            AccountRef = xlsRef?.Account ?? "",
                            DoubleTake = doubleTake
                        });
                    }
                }

                AnalyticsSaveDialog.ShowSaveDialog(diffs, assumptions);

            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
