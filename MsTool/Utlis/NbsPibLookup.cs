using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace MsTool.Utlis
{
    public static class NbsPibLookup
    {
        // Single handler, but we disable auto-redirect (we parse the POST directly).
        private static readonly HttpClientHandler _handler = new HttpClientHandler
        {
            AllowAutoRedirect = false,
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
        };

        private static readonly HttpClient _http = new HttpClient(_handler)
        {
            Timeout = TimeSpan.FromSeconds(10)
        };

        static NbsPibLookup()
        {
            // A realistic User-Agent so the server treats us like a browser:
            _http.DefaultRequestHeaders.UserAgent.ParseAdd(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
                "AppleWebKit/537.36 (KHTML, like Gecko) " +
                "Chrome/112.0.0.0 Safari/537.36"
            );
            // Accept headers for HTML
            _http.DefaultRequestHeaders.Accept.ParseAdd("text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8");
        }

        public static async Task<string> LookupNameAsync(string pib)
        {
            if (string.IsNullOrWhiteSpace(pib))
                return "";

            const string url = "https://www.nbs.rs/rir_pn/pn_rir.html.jsp?type=rir_results&lang=SER_CIR&konverzija=yes";
            var form = new Dictionary<string, string>
            {
                ["pib"] = pib,
                ["Submit"] = "Pretraži"
            };

            for (int attempt = 1; attempt <= 3; attempt++)
            {
                try
                {
                    Debug.WriteLine($"[DEBUG] Attempt #{attempt}: POST PIB={pib}");
                    using var content = new FormUrlEncodedContent(form);
                    var resp = await _http.PostAsync(url, content);
                    Debug.WriteLine($"[DEBUG] Status: {(int)resp.StatusCode} {resp.StatusCode}");

                    var html = await resp.Content.ReadAsStringAsync();
                    Debug.WriteLine($"[DEBUG] HTML snippet:\n{html.Substring(0, Math.Min(200, html.Length))}...\n---");

                    var doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(html);
                    var node = doc.DocumentNode.SelectSingleNode("//input[@name='nazivULinku']");
                    if (node != null)
                    {
                        var raw = node.GetAttributeValue("value", "");
                        Debug.WriteLine($"[DEBUG] Found nazivULinku = {raw}");
                        return raw.Trim('\"');
                    }
                    Debug.WriteLine("[DEBUG] No nazivULinku input found in response.");
                    return "";
                }
                catch (HttpRequestException ex) when (ex.InnerException is IOException)
                {
                    Debug.WriteLine($"[WARN] Network error on attempt {attempt}: {ex.Message}");
                    // exponential backoff: 500ms, 1000ms, 2000ms
                    await Task.Delay(500 * (1 << attempt - 1));
                }
                catch (TaskCanceledException ex) when (!ex.CancellationToken.IsCancellationRequested)
                {
                    Debug.WriteLine($"[WARN] Timeout on attempt {attempt}: {ex.Message}");
                    await Task.Delay(500 * (1 << attempt - 1));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ERROR] Unexpected error on attempt {attempt}: {ex}");
                    break;
                }
            }

            Debug.WriteLine($"[ERROR] All attempts failed for PIB={pib}");
            return "";
        }
    }
}
