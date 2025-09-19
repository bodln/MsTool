using MsTool.Models;

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
                var xlsMainRecs = FileManipulator.LoadXlsAnalytics(xlsMainPath, true); // What csv was for BoB
                var xlsRefRecs = FileManipulator.LoadXlsAnalytics(xlsRefPath, false);

                List<DiffAnalyticsRecord> diffs = new List<DiffAnalyticsRecord>();

                foreach (var key in xlsMainRecs.Keys)
                {
                    var xlsMain = xlsMainRecs[key];
                    xlsRefRecs.TryGetValue(key, out var xlsRef);

                    double refDebit = xlsRef?.ValueDebit ?? 0;
                    double refCredit = xlsRef?.ValueCredit ?? 0;

                    double mainDebit = xlsMain?.ValueDebit ?? 0;
                    double mainCredit = xlsMain?.ValueCredit ?? 0;

                    bool altCompFlag = xlsMain!.AltCompFlag;

                    bool equalityAssumption = false;

                    bool equal = true;

                    double difference = 0;

                    if (xlsRef == null)
                    {
                        equal = false;

                        if (!altCompFlag)
                        {
                            foreach (var refRec in xlsRefRecs)
                            {
                                if (refRec.Value.Date.ToString() == xlsMain.Date.ToString() &&
                                    refRec.Value.AltCompFlag == false &&
                                    (Math.Abs(mainDebit - refRec.Value.ValueCredit) <= 5.0) == true)
                                {
                                    equalityAssumption = true;
                                }
                            }
                        }
                    }
                    else if (altCompFlag)
                    {
                        if (!(Math.Abs(mainCredit - refDebit) <= 5.0))
                        {
                            equal = false;
                            difference = mainCredit - refDebit;
                        }
                    }
                    else if (!(Math.Abs(mainDebit - refCredit) <= 5.0))
                    {
                        equal = false;
                    }

                    if (!equal || equalityAssumption)
                    {
                        diffs.Add(new DiffAnalyticsRecord
                        {
                            OriginalMainKey = altCompFlag ? "-" : xlsMain.OriginalKey,
                            OriginalRefKey = xlsRef?.OriginalKey ?? "Nema",
                            ValueDebit = mainDebit,
                            ValueCreditDiff = difference,
                            DateMain = xlsMain.Date,
                            DateRef = xlsRef?.Date ?? "",
                            AccountMain = xlsMain.Account,
                            AccountRef = xlsRef?.Account ?? "",
                            AssumedEqual = equalityAssumption
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
