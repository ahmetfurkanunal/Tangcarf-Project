using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using ClosedXML.Excel;
using Newtonsoft.Json;
using WinTimer = System.Windows.Forms.Timer;
// ClosedXML ile System.Xml.Linq LoadOptions çakışmasın:
using XLoadOptions = System.Xml.Linq.LoadOptions;

namespace XmlToExcel
{
    public class Form1 : Form
    {
        // ---------- UI ----------
        private Button btnStartStop;
        private Label lblCountdown;
        private Label lblProducts;
        private TextBox txtLog;

        // ---------- Timer/State ----------
        private const int IntervalMinutes = 20;
        private const int IntervalSeconds = IntervalMinutes * 60;
        private readonly WinTimer _uiTimer = new WinTimer();
        private int _remainingSeconds = IntervalSeconds;
        private bool _running;
        private bool _processing;
        private CancellationTokenSource _cts;

        public Form1()
        {
            InitUi();

            _uiTimer.Interval = 1000;
            _uiTimer.Tick += async (s, e) =>
            {
                if (!_running) return;
                if (_remainingSeconds > 0)
                {
                    _remainingSeconds--;
                    UpdateCountdown();
                }
                else
                {
                    await RunOnceLoopAsync();
                }
            };

            UpdateCountdown();
            UpdateProductCount(0);
        }

        private void InitUi()
        {
            Text = "XML → Excel + Eşleme + Trendyol (WinForms)";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new System.Drawing.Size(1000, 600);

            var lblStart = new Label { Text = "Başlat", Left = 20, Top = 18, AutoSize = true, Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold) };
            Controls.Add(lblStart);

            btnStartStop = new Button { Text = "Start", Left = 20, Top = 45, Width = 120, Height = 36 };
            btnStartStop.Click += async (s, e) => { if (_running) await StopAsync(); else await StartAsync(); };
            Controls.Add(btnStartStop);

            var lblKalan = new Label { Text = "Kalan:", Left = 820, Top = 18, AutoSize = true, Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold) };
            Controls.Add(lblKalan);

            lblCountdown = new Label { Text = "00:20:00", Left = 820, Top = 45, AutoSize = true, Font = new System.Drawing.Font("Consolas", 18, System.Drawing.FontStyle.Bold) };
            Controls.Add(lblCountdown);

            var lblUrun = new Label { Text = "Ürün Sayısı:", Left = 820, Top = 100, AutoSize = true };
            Controls.Add(lblUrun);

            lblProducts = new Label { Text = "0", Left = 820, Top = 122, AutoSize = true, Font = new System.Drawing.Font("Segoe UI", 14, System.Drawing.FontStyle.Bold) };
            Controls.Add(lblProducts);

            txtLog = new TextBox { Left = 20, Top = 100, Width = 760, Height = 460, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true };
            Controls.Add(txtLog);
        }

        // ---------- Start/Stop ----------
        private async Task SendStocksToTSoftAsync(string trendyolExcelPath, CancellationToken ct)
        {
            await SendVariantStocksToTSoftAsync(trendyolExcelPath, ct);
        }
        private async Task StartAsync()
        {
            _running = true;
            btnStartStop.Text = "Stop";
            Log("Süreç başlatıldı. İlk tur çalışıyor...");
            _cts = new CancellationTokenSource();
            await RunJobAndReportAsync(_cts.Token);
            _remainingSeconds = IntervalSeconds;
            UpdateCountdown();
            _uiTimer.Start();
        }

        private async Task StopAsync()
        {
            _running = false;
            btnStartStop.Text = "Start";
            _uiTimer.Stop();
            if (_cts != null) { _cts.Cancel(); _cts.Dispose(); _cts = null; }
            Log("Süreç durduruldu.");
            await Task.CompletedTask;
        }

        private async Task RunOnceLoopAsync()
        {
            if (_cts == null) _cts = new CancellationTokenSource();
            await RunJobAndReportAsync(_cts.Token);
            _remainingSeconds = IntervalSeconds;
            UpdateCountdown();
        }

        private async Task RunJobAndReportAsync(CancellationToken ct)
        {
            if (_processing) { Log("Önceki tur bitmedi. Atlandı."); return; }
            var sw = System.Diagnostics.Stopwatch.StartNew();
            _processing = true;
            try
            {
                var result = await RunJobAsync(ct);
                sw.Stop();
                UpdateProductCount(result.Processed);
                Log($"Tur tamamlandı • Ürün: {result.Processed} • Süre: {sw.Elapsed:mm\\:ss} • {result.Message}");
            }
            catch (OperationCanceledException) { Log("Tur iptal edildi."); }
            catch (Exception ex) { Log("Hata: " + ex.Message); }
            finally { _processing = false; }
        }

        // =====================================================
        //                   ANA İŞ AKIŞI
        // =====================================================
        private async Task<(int Processed, string Message)> RunJobAsync(CancellationToken ct)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            string baseDir = FindProjectRoot(AppDomain.CurrentDomain.BaseDirectory);
            string outputs = Path.Combine(baseDir, "Outputs");
            Directory.CreateDirectory(outputs);

            // ---------- xml.txt ----------
            string xmlTxt = Path.Combine(baseDir, "Parameters", "xml.txt");
            if (!File.Exists(xmlTxt)) throw new InvalidOperationException("Parameters\\xml.txt yok.");

            var lines = File.ReadAllLines(xmlTxt, Encoding.UTF8)
                            .Select(t => t.Trim())
                            .Where(t => t.Length > 0 && !t.StartsWith("#"))
                            .ToArray();
            if (lines.Length == 0) throw new InvalidOperationException("xml.txt boş.");

            // 1. satır zorunlu (birincil kaynak)
            string primarySource = ResolvePath(baseDir, lines[0]);

            // 2. satır URL/dosya ise yedek; değilse çıktı adı
            string backupSource = null;
            string stemOverride = null;
            if (lines.Length >= 2)
            {
                var cand = lines[1];
                if (IsHttp(cand) || File.Exists(ResolvePath(baseDir, cand)))
                    backupSource = ResolvePath(baseDir, cand);
                else
                    stemOverride = Path.GetFileNameWithoutExtension(cand);
            }
            if (lines.Length >= 3 && string.IsNullOrWhiteSpace(stemOverride))
                stemOverride = Path.GetFileNameWithoutExtension(lines[2]);

            string stem = !string.IsNullOrWhiteSpace(stemOverride)
                            ? stemOverride
                            : "Products_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

            // ---------- XML yükle (fallback + cache) ----------
            Log("XML yükleniyor...");
            Log("Kaynak: " + primarySource);

            string cachePath = Path.Combine(outputs, "xml_cache.xml");
            var xdoc = await TryLoadWithFallbackAsync(primarySource, backupSource, cachePath, ct);

            // ---------- Parse ----------
            var variantTypes = new HashSet<string>(StringComparer.Ordinal);
            var stockLocations = new HashSet<string>(StringComparer.Ordinal);
            foreach (var v in xdoc.Descendants().Where(e => e.Name.LocalName == "variant"))
            {
                foreach (var vv in v.Descendants().Where(e => e.Name.LocalName == "variantValue"))
                { var t = GetChildValue(vv, "variantTypeName"); if (!string.IsNullOrWhiteSpace(t)) variantTypes.Add(t.Trim()); }
                foreach (var s in v.Descendants().Where(e => e.Name.LocalName == "stock"))
                { var loc = GetChildValue(s, "stockLocationName"); if (!string.IsNullOrWhiteSpace(loc)) stockLocations.Add(loc.Trim()); }
            }

            var rows = new List<Dictionary<string, string>>();
            foreach (var p in xdoc.Descendants().Where(e => e.Name.LocalName == "product"))
            {
                string productId = GetChildValue(p, "id");
                string productName = GetChildValue(p, "name");
                string googleTax = GetChildValue(p, "googleTaxonomyId");
                string salesChannels = GetChildValue(p, "salesChannelIds");
                string brandName = GetChildValue(GetChild(p, "brand"), "name");
                string productDesc = GetChildValue(p, "description");

                var catIds = new List<string>(); var catNames = new List<string>();
                foreach (var cat in p.Descendants().Where(e => e.Name.LocalName == "category"))
                {
                    var cid = GetChildValue(cat, "id"); if (!string.IsNullOrWhiteSpace(cid)) catIds.Add(cid.Trim());
                    foreach (var nm in cat.Elements().Where(x => x.Name.LocalName == "name"))
                    { var val = (nm.Value ?? "").Trim(); if (val.Length > 0) catNames.Add(val); }
                }

                foreach (var v in p.Descendants().Where(e => e.Name.LocalName == "variant"))
                {
                    var row = new Dictionary<string, string>(StringComparer.Ordinal);
                    row["Product.Id"] = productId;
                    row["Product.Name"] = productName;
                    row["Brand.Name"] = brandName;
                    row["Category.Names"] = string.Join(" | ", catNames.Distinct());
                    row["Category.Ids"] = string.Join(" | ", catIds.Distinct());
                    row["Product.GoogleTaxonomyId"] = googleTax;
                    row["Product.SalesChannelIds"] = salesChannels;
                    row["Product.Description"] = productDesc;

                    row["Variant.Id"] = GetChildValue(v, "id");
                    row["Variant.Sku"] = GetChildValue(v, "sku");
                    row["Variant.Description"] = GetChildValue(v, "description");

                    var barcodes = v.Descendants().Where(e => e.Name.LocalName == "barcode")
                        .Select(x => (x.Value ?? "").Trim()).Where(x => x.Length > 0).Distinct().ToList();
                    row["Variant.Barcodes"] = string.Join(" | ", barcodes);

                    var images = v.Descendants().Where(e => e.Name.LocalName == "image").ToList();
                    string mainUrl = "";
                    if (images.Count > 0)
                    {
                        var main = images.FirstOrDefault(img => string.Equals(GetChildValue(img, "isMain"), "true", StringComparison.OrdinalIgnoreCase))
                                   ?? images.FirstOrDefault(img => GetChildValue(img, "order") == "0")
                                   ?? images.First();
                        mainUrl = GetChildValue(main, "imageUrl");
                    }
                    var allUrls = images.Select(img => GetChildValue(img, "imageUrl")).Where(u => !string.IsNullOrWhiteSpace(u)).Distinct().ToList();
                    row["Variant.Image.MainUrl"] = mainUrl;
                    row["Variant.Image.AllUrls"] = string.Join(" | ", allUrls);
                    row["Variant.Image.Count"] = images.Count.ToString();

                    var price = v.Descendants().FirstOrDefault(e => e.Name.LocalName == "price");
                    row["Variant.Price.SellPrice"] = GetChildValue(price, "sellPrice");
                    row["Variant.Price.DiscountPrice"] = GetChildValue(price, "discountPrice");

                    int total = 0; var details = new List<string>();
                    foreach (var s in v.Descendants().Where(e => e.Name.LocalName == "stock"))
                    {
                        string loc = (GetChildValue(s, "stockLocationName") ?? "").Trim();
                        int cnt; int.TryParse((GetChildValue(s, "stockCount") ?? "").Trim(), out cnt);
                        total += cnt;
                        if (loc.Length > 0)
                        { details.Add(loc + ":" + cnt.ToString()); row["Variant.Stock." + loc] = cnt.ToString(); }
                    }
                    row["Variant.Stock.Total"] = total.ToString();
                    row["Variant.Stock.Details"] = string.Join(" | ", details);

                    var vvalues = v.Descendants().Where(e => e.Name.LocalName == "variantValue").ToList();
                    var vvPairs = new List<string>();
                    foreach (var vv in vvalues)
                    {
                        string t = (GetChildValue(vv, "variantTypeName") ?? "").Trim();
                        string val = (GetChildValue(vv, "variantValueName") ?? "").Trim();
                        if (t.Length > 0 && val.Length > 0)
                        { vvPairs.Add(t + "=" + val); row["Variant." + t] = MergeCell(row, "Variant." + t, val); }
                    }
                    row["Variant.VariantValues"] = string.Join(" | ", vvPairs);

                    rows.Add(row);
                }
            }

            var productKeys = new HashSet<string>(rows.SelectMany(r => r.Keys), StringComparer.Ordinal);
            var rowByBarcode = BuildRowByBarcode(rows);

            // ---------- Products.xlsx ----------
            var allCols = new HashSet<string>(rows.SelectMany(r => r.Keys), StringComparer.Ordinal);
            foreach (var t in variantTypes) allCols.Add("Variant." + t);
            foreach (var loc in stockLocations) allCols.Add("Variant.Stock." + loc);
            var ordered = OrderColumns(allCols, variantTypes, stockLocations);

            var dt = new DataTable("Products");
            foreach (var c in ordered) dt.Columns.Add(c, typeof(string));
            foreach (var r in rows)
            {
                var dr = dt.NewRow();
                foreach (var c in ordered) dr[c] = r.ContainsKey(c) ? r[c] : "";
                dt.Rows.Add(dr);
            }

            // >>> güvenli kaydetme
            string productsPath = SaveProductsExcel(dt, outputs, stem);
            Log($"Products yazıldı → {productsPath} (satır: {dt.Rows.Count})");

            // ===== TSoft Excel üret =====
            string tsoftExcel = WriteTSoftExcel(rows, outputs);
            Log($"TSoft Excel yazıldı → {tsoftExcel}");


            // ---------- XML stok sözlüğü ----------
            var stockByBarcode = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var r in rows)
            {
                int.TryParse(GetVal(r, "Variant.Stock.Total"), out int total);
                foreach (var b in (GetVal(r, "Variant.Barcodes") ?? "").Split('|'))
                {
                    var norm = NormalizeBarcode(b);
                    if (norm.Length == 0) continue;
                    stockByBarcode[norm] = total;
                }
            }

            // ---------- Eşleme + Trendyol ----------
            string eslemeIn = Path.Combine(baseDir, "Parameters", "esleme.xlsx");
            int trendyolCount; 
            string eslemeOut, trendyolOut;

            if (File.Exists(eslemeIn))
            {
                (eslemeOut, trendyolOut, trendyolCount) =
                    await FillEslemeAndTrendyolAsync(eslemeIn, outputs, stockByBarcode, rowByBarcode, productKeys, ct);
                Log($"Eşleme dolduruldu → {eslemeOut}");
                Log($"Trendyol tablosu yazıldı ({trendyolCount} satır) → {trendyolOut}");
            }
            else
            {
                trendyolOut = Path.Combine(outputs, "Trendyol_Stok_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
                var pairs = stockByBarcode.OrderBy(k => k.Key)
                                          .Select(kv => new KeyValuePair<string, int>(kv.Key, kv.Value));
                var template = Path.Combine(baseDir, "Parameters", "Trendyol.xlsx");
                if (File.Exists(template)) WriteTrendyolFromTemplate(template, trendyolOut, pairs);
                else WriteTrendyolFromTemplateFallback(trendyolOut, pairs);
                trendyolCount = stockByBarcode.Count;
            }

            // ---------- Trendyol POST ----------
            int sent = 0;
            try
            {
                var cfg = await LoadTrendyolConfigAsync(baseDir, ct);
                var items = ReadTrendyolItemsFromExcel(trendyolOut);
                foreach (var chunk in Chunk(items, cfg.BatchSize))
                {
                    var id = await PostInventoryAsync(cfg, chunk, ct);
                    sent += chunk.Count;
                    Log($"POST OK • adet: {chunk.Count} • batchRequestId: {id}");
                    await Task.Delay(400, ct);
                }
            }
            catch (Exception ex) { Log("Trendyol gönderim uyarı: " + ex.Message); }

            // ----------- TSOFT POST EKLENDİ -----------
            try
            {
                await SendVariantStocksToTSoftAsync(productsPath, ct);
            }
            catch (Exception ex)

            {
                Log("TSOFT gönderim uyarı: " + ex.Message);
            }


            return (sent, $"Products:{dt.Rows.Count}, Trendyol gönderilen:{sent}");
        }

        // =============== XML yükleme + fallback ===============
        private static async Task<XDocument> LoadXDocumentAsync(string input, CancellationToken ct)
        {
            if (!IsHttp(input))
            {
                using (var fs = new FileStream(input, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, true))
                {
                    return await Task.Run(() =>
                        XDocument.Load(fs, XLoadOptions.PreserveWhitespace | XLoadOptions.SetLineInfo), ct);
                }
            }

            using (var handler = new HttpClientHandler
            { AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate })
            using (var http = new HttpClient(handler))
            {
                http.Timeout = TimeSpan.FromSeconds(60);
                using (var resp = await http.GetAsync(input, HttpCompletionOption.ResponseHeadersRead, ct))
                {
                    resp.EnsureSuccessStatusCode();
                    using (var stream = await resp.Content.ReadAsStreamAsync())
                    {
                        return await Task.Run(() =>
                            XDocument.Load(stream, XLoadOptions.PreserveWhitespace | XLoadOptions.SetLineInfo), ct);
                    }
                }
            }
        }

        private static async Task<XDocument> TryLoadWithFallbackAsync(
            string primary, string backup, string cachePath, CancellationToken ct)
        {
            // önce yedek
            if (!string.IsNullOrWhiteSpace(backup) && (IsHttp(backup) || File.Exists(backup)))
            {
                try
                {
                    var xd = await LoadXDocumentAsync(backup, ct);
                    try { xd.Save(cachePath); } catch { }
                    return xd;
                }
                catch { /* primary'ye geç */ }
            }

            // sonra primary
            try
            {
                var xd = await LoadXDocumentAsync(primary, ct);
                try { xd.Save(cachePath); } catch { }
                return xd;
            }
            catch
            {
                if (File.Exists(cachePath))
                    return XDocument.Load(cachePath, XLoadOptions.PreserveWhitespace | XLoadOptions.SetLineInfo);
                throw;
            }
        }

        private static bool IsHttp(string s) =>
            s.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            s.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

        private static string ResolvePath(string baseDir, string p)
        {
            if (string.IsNullOrWhiteSpace(p)) return p;
            if (IsHttp(p)) return p;
            if (Path.IsPathRooted(p)) return p;
            return Path.Combine(baseDir, p);
        }

        // ================= Dinamik eşleme + Trendyol =================
        // Eşleme: bulunmayanlar boş kalsın.
        // Trendyol Excel: her satır yazılsın, bulunmayanlar 0 olsun.
        private static async Task<(string eslemeOut, string trendyolOut, int count)>
            FillEslemeAndTrendyolAsync(
                string eslemeIn,
                string outputs,
                Dictionary<string, int> stockByXmlBarcode,
                Dictionary<string, Dictionary<string, string>> rowByBarcode,
                HashSet<string> productKeys,
                CancellationToken ct)
        {
            string eslemeOut = Path.Combine(outputs, "Eslestirme_Stok_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
            string trendyolOut = Path.Combine(outputs, "Trendyol_Stok_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
            int written = 0;

            await Task.Run(() =>
            {
                using (var wb = new XLWorkbook(eslemeIn))
                {
                    var ws = wb.Worksheets.First();
                    FindEslemeHeaders(ws, out int headerRow, out int colEan, out int colXml, out int colTotal);

                    // Dinamik kolonlar
                    var dynamicCols = new List<(int col, string key)>();
                    int headerLastCol = ws.Row(headerRow).LastCellUsed()?.Address.ColumnNumber ?? 0;

                    var reserved = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    { "ean13","ean-13","ean","barkod","barcode","xmlbarcode","xmlbarkod","barcodexml","xml",
                      "variant.stock.total","urunstokadedi","ürünstokadedi","stock","stok" };

                    for (int c = 1; c <= headerLastCol; c++)
                    {
                        string h = ws.Cell(headerRow, c).GetString();
                        if (!string.IsNullOrWhiteSpace(h) && productKeys.Contains(h) && !reserved.Contains(h.Trim().ToLowerInvariant()))
                            dynamicCols.Add((c, h));
                    }

                    // Eşleme için stok sütunu yoksa ekle
                    if (colTotal == 0)
                    {
                        colTotal = Math.Max(headerLastCol + 1, Math.Max(colEan, colXml) + 1);
                        ws.Cell(headerRow, colTotal).Value = "Variant.Stock.Total";
                    }

                    int last = ws.LastRowUsed()?.RowNumber() ?? headerRow;
                    var trendyolPairs = new List<KeyValuePair<string, int>>();

                    for (int r = headerRow + 1; r <= last; r++)
                    {
                        string ean = NormalizeBarcode(ws.Cell(r, colEan).Value.ToString());
                        if (ean.Length == 0) continue;

                        string xmlB = (colXml > 0) ? NormalizeBarcode(ws.Cell(r, colXml).GetString()) : ean;
                        if (xmlB.Length == 0) xmlB = ean;

                        // --- EŞLEME: sadece bulursa yaz (boş kalabilir) ---
                        if (stockByXmlBarcode.TryGetValue(xmlB, out int stockFound))
                        {
                            ws.Cell(r, colTotal).Value = stockFound;
                            written++;
                        }

                        // Dinamik alanlar
                        if (rowByBarcode.TryGetValue(xmlB, out var prodRow))
                        {
                            foreach (var (col, key) in dynamicCols)
                                ws.Cell(r, col).Value = (prodRow.TryGetValue(key, out var v) ? v : "") ?? "";
                        }

                        // --- TRENDYOL: her satır eklensin; bulunamazsa 0 ---
                        int qtyForTrendyol = stockByXmlBarcode.TryGetValue(xmlB, out var s) ? s : 0;
                        trendyolPairs.Add(new KeyValuePair<string, int>(ean, qtyForTrendyol));
                    }

                    // Eşleme kaydet
                    ws.Columns().AdjustToContents();
                    wb.SaveAs(eslemeOut);

                    // Trendyol dosyası
                    var template = Path.Combine(Path.GetDirectoryName(outputs) ?? "", "Parameters", "Trendyol.xlsx");
                    if (File.Exists(template)) WriteTrendyolFromTemplate(template, trendyolOut, trendyolPairs);
                    else WriteTrendyolFromTemplateFallback(trendyolOut, trendyolPairs);
                }
            }, ct);

            return (eslemeOut, trendyolOut, written);
        }

        private static void FindEslemeHeaders(IXLWorksheet ws, out int headerRow, out int colEan, out int colXml, out int colTotal)
        {
            headerRow = 1; colEan = 1; colXml = 0; colTotal = 0;
            Func<string, string> N = s => (s ?? "").Trim().ToLowerInvariant()
                .Replace('ç', 'c').Replace('ğ', 'g').Replace('ı', 'i').Replace('ö', 'o').Replace('ş', 's').Replace('ü', 'u')
                .Replace(" ", "").Replace("-", "").Replace("_", "");
            int lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
            int maxRow = Math.Min(10, lastRow);
            for (int r = 1; r <= maxRow; r++)
            {
                int lastCol = ws.Row(r).LastCellUsed()?.Address.ColumnNumber ?? 0;
                if (lastCol == 0) continue;
                int e = 0, x = 0, t = 0;
                for (int c = 1; c <= lastCol; c++)
                {
                    string h = N(ws.Cell(r, c).GetString());
                    if (h == "ean13" || h == "ean-13" || h == "ean" || h == "barkod" || h == "barcode") e = c;
                    if (h == "xmlbarcode" || h == "xmlbarkod" || h == "barcodexml" || h == "xml") x = c;
                    if (h == "variant.stock.total" || h == "urunstokadedi" || h == "ürünstokadedi" || h == "stock" || h == "stok") t = c;
                }
                if (e > 0) { headerRow = r; colEan = e; colXml = x; colTotal = t; return; }
            }
        }

        // =============== Trendyol sheet — ŞABLONLA BİREBİR ===============
        private static void WriteTrendyolFromTemplate(string templatePath, string outputPath, IEnumerable<KeyValuePair<string, int>> pairs)
        {
            using (var wb = new XLWorkbook(templatePath))
            {
                var ws = wb.Worksheets.FirstOrDefault(s => s.Name.Equals("Güncelleme Bilgileri", StringComparison.OrdinalIgnoreCase))
                         ?? wb.Worksheets.First();

                FindTrendyolHeadersExact(ws, out int headerRow, out int colBarcode, out int colQty, out int colList, out int colSale);

                int lastRow = ws.LastRowUsed()?.RowNumber() ?? headerRow;
                if (lastRow > headerRow)
                    ws.Range(headerRow + 1, 1, lastRow, 4).Clear();

                int r = headerRow + 1;
                foreach (var kv in pairs)
                {
                    ws.Cell(r, colBarcode).Value = kv.Key;
                    ws.Cell(r, colBarcode).Style.NumberFormat.Format = "@";
                    ws.Cell(r, colQty).Value = kv.Value;
                    r++;
                }

                ws.Columns(colBarcode, colQty).AdjustToContents();
                wb.SaveAs(outputPath);
            }
        }

        private static void FindTrendyolHeadersExact(IXLWorksheet ws, out int headerRow, out int colBarcode, out int colQty, out int colList, out int colSale)
        {
            headerRow = 1; colBarcode = 1; colQty = 4; colList = 2; colSale = 3;

            const string H_BAR = "Barkod";
            const string H_LIST = "Piyasa Satış Fiyatı (KDV Dahil)";
            const string H_SALE = "Trendyol'da  Satılacak Fiyat (KDV Dahil)";
            const string H_QTY = "Ürün Stok Adedi";

            int lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
            int scanRows = Math.Min(15, lastRow);

            for (int r = 1; r <= scanRows; r++)
            {
                int lastCol = ws.Row(r).LastCellUsed()?.Address.ColumnNumber ?? 0;
                if (lastCol == 0) continue;

                int b = 0, lp = 0, sp = 0, q = 0;
                for (int c = 1; c <= lastCol; c++)
                {
                    string h = ws.Cell(r, c).GetString().Trim();
                    if (h == H_BAR) b = c;
                    else if (h == H_LIST) lp = c;
                    else if (h == H_SALE) sp = c;
                    else if (h == H_QTY) q = c;
                }

                if (b > 0 && q > 0)
                {
                    headerRow = r; colBarcode = b; colQty = q; colList = lp; colSale = sp;
                    return;
                }
            }

            throw new InvalidOperationException("Şablon başlıkları bulunamadı. Lütfen Parameters/Trendyol.xlsx birebir olsun.");
        }

        // Şablon yoksa sade dosya
        private static void WriteTrendyolFromTemplateFallback(string outputPath, IEnumerable<KeyValuePair<string, int>> pairs)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Güncelleme Bilgileri");
                ws.Cell(1, 1).Value = "Barkod";
                ws.Cell(1, 2).Value = "Piyasa Satış Fiyatı (KDV Dahil)";
                ws.Cell(1, 3).Value = "Trendyol'da  Satılacak Fiyat (KDV Dahil)";
                ws.Cell(1, 4).Value = "Ürün Stok Adedi";

                ws.Column(1).Style.NumberFormat.Format = "@";
                ws.Column(4).Style.NumberFormat.Format = "0";

                int r = 2;
                foreach (var kv in pairs)
                {
                    ws.Cell(r, 1).Value = kv.Key;
                    ws.Cell(r, 4).Value = kv.Value;
                    r++;
                }
                ws.Columns(1, 4).AdjustToContents();
                wb.SaveAs(outputPath);
            }
        }

        // =============== Trendyol POST ===============
        private sealed class TrendyolCfg
        {
            public string BaseUrl;
            public string SellerId;
            public string ApiKey;
            public string ApiSecret;
            public string UserAgent = "XmlToExcelWinForms";
            public int TimeoutSeconds = 60;
            public int BatchSize = 100;
        }

        private static async Task<TrendyolCfg> LoadTrendyolConfigAsync(string baseDir, CancellationToken ct)
        {
            string path = Path.Combine(baseDir, "Parameters", "trendyol.postman.json.txt");
            if (File.Exists(path))
            {
                string json;
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, true))
                using (var sr = new StreamReader(fs, Encoding.UTF8))
                { json = await sr.ReadToEndAsync(); }

                Func<string, string> Get = key =>
                {
                    var m = Regex.Match(json, "\"" + Regex.Escape(key) + "\"\\s*:\\s*\"([^\"]*)\"", RegexOptions.IgnoreCase);
                    return m.Success ? m.Groups[1].Value.Trim() : null;
                };

                var cfg = new TrendyolCfg
                {
                    BaseUrl = Get("baseUrl"),
                    SellerId = Get("sellerId"),
                    ApiKey = Get("apiKey"),
                    ApiSecret = Get("apiSecret")
                };

                if (!string.IsNullOrWhiteSpace(cfg.BaseUrl) &&
                    cfg.BaseUrl.IndexOf("{sellerId}", StringComparison.OrdinalIgnoreCase) >= 0)
                    cfg.BaseUrl = cfg.BaseUrl.Replace("{sellerId}", cfg.SellerId);

                if (string.IsNullOrWhiteSpace(cfg.BaseUrl) ||
                    string.IsNullOrWhiteSpace(cfg.SellerId) ||
                    string.IsNullOrWhiteSpace(cfg.ApiKey) ||
                    string.IsNullOrWhiteSpace(cfg.ApiSecret))
                    throw new InvalidOperationException("trendyol.postman.json.txt eksik.");

                cfg.BaseUrl = cfg.BaseUrl
                    .Replace("https://api.trendyol.com/sapigw", "https://apigw.trendyol.com")
                    .Replace("https://stageapi.trendyol.com/sapigw", "https://stageapigw.trendyol.com");

                return cfg;
            }

            // Fallback: trendyol.txt
            string txt = Path.Combine(baseDir, "Parameters", "trendyol.txt");
            if (File.Exists(txt))
            {
                var dict = File.ReadAllLines(txt, Encoding.UTF8)
                    .Select(l => l.Trim()).Where(l => l.Length > 0 && !l.StartsWith("#"))
                    .Select(l => l.Split(new[] { '=' }, 2)).Where(a => a.Length == 2)
                    .ToDictionary(a => a[0].Trim().ToUpperInvariant(), a => a[1].Trim());

                var cfg = new TrendyolCfg
                {
                    BaseUrl = dict.ContainsKey("URL") ? dict["URL"] : null,
                    SellerId = dict.ContainsKey("SELLER_ID") ? dict["SELLER_ID"] : null,
                    ApiKey = dict.ContainsKey("API_KEY") ? dict["API_KEY"] : null,
                    ApiSecret = dict.ContainsKey("API_SECRET") ? dict["API_SECRET"] : null,
                    UserAgent = dict.ContainsKey("USER_AGENT") ? dict["USER_AGENT"] : "XmlToExcelWinForms",
                    TimeoutSeconds = dict.ContainsKey("TIMEOUT_SECONDS") && int.TryParse(dict["TIMEOUT_SECONDS"], out int s) ? s : 60,
                    BatchSize = dict.ContainsKey("BATCH_SIZE") && int.TryParse(dict["BATCH_SIZE"], out int b) ? Math.Max(1, b) : 100
                };

                if (cfg.BaseUrl.IndexOf("{sellerId}", StringComparison.OrdinalIgnoreCase) >= 0)
                    cfg.BaseUrl = cfg.BaseUrl.Replace("{sellerId}", cfg.SellerId);

                if (string.IsNullOrWhiteSpace(cfg.BaseUrl) ||
                    string.IsNullOrWhiteSpace(cfg.SellerId) ||
                    string.IsNullOrWhiteSpace(cfg.ApiKey) ||
                    string.IsNullOrWhiteSpace(cfg.ApiSecret))
                    throw new InvalidOperationException("trendyol.txt eksik.");

                return cfg;
            }

            throw new InvalidOperationException("Trendyol ayar dosyası bulunamadı (postman.json.txt veya trendyol.txt).");
        }

        private sealed class PostItem
        {
            public string barcode;
            public int quantity;
            public decimal? listPrice;
            public decimal? salePrice;
            public string currencyType;
        }

        private static List<PostItem> ReadTrendyolItemsFromExcel(string excelPath)
        {
            var list = new List<PostItem>();
            using (var wb = new XLWorkbook(excelPath))
            {
                var ws = wb.Worksheets.First();
                FindTrendyolHeaders(ws, out int h, out int cBar, out int cQty, out int cList, out int cSale);
                int last = ws.LastRowUsed()?.RowNumber() ?? h;
                for (int r = h + 1; r <= last; r++)
                {
                    string b = ws.Cell(r, cBar).GetString().Trim();
                    if (b.StartsWith("'")) b = b.Substring(1);
                    b = new string(b.Where(char.IsDigit).ToArray());
                    if (b.Length == 0) continue;

                    int q = 0;
                    var sQ = ws.Cell(r, cQty).GetString().Trim();
                    if (!int.TryParse(sQ, NumberStyles.Integer, CultureInfo.InvariantCulture, out q))
                        int.TryParse(Convert.ToString(ws.Cell(r, cQty).Value, CultureInfo.InvariantCulture), out q);
                    q = Math.Max(0, q);

                    decimal? lp = null, sp = null;
                    if (cList > 0 && decimal.TryParse(ws.Cell(r, cList).GetString().Trim(), NumberStyles.Any, CultureInfo.GetCultureInfo("tr-TR"), out decimal v1)) lp = v1;
                    if (cSale > 0 && decimal.TryParse(ws.Cell(r, cSale).GetString().Trim(), NumberStyles.Any, CultureInfo.GetCultureInfo("tr-TR"), out decimal v2)) sp = v2;

                    list.Add(new PostItem { barcode = b, quantity = q, listPrice = lp, salePrice = sp, currencyType = (lp.HasValue || sp.HasValue) ? "TRY" : null });
                }
            }
            return list;
        }

        private static void FindTrendyolHeaders(IXLWorksheet ws, out int headerRow, out int colBarcode, out int colQty, out int colList, out int colSale)
        {
            headerRow = 1; colBarcode = 1; colQty = 4; colList = 0; colSale = 0;
            Func<string, string> N = s => (s ?? "").Trim().ToLowerInvariant()
                .Replace('ç', 'c').Replace('ğ', 'g').Replace('ı', 'i').Replace('ö', 'o').Replace('ş', 's').Replace('ü', 'u')
                .Replace(" ", "").Replace("-", "").Replace("_", "");
            int lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
            int maxRow = Math.Min(15, lastRow);
            for (int r = 1; r <= maxRow; r++)
            {
                int lastCol = ws.Row(r).LastCellUsed()?.Address.ColumnNumber ?? 0;
                if (lastCol == 0) continue;
                int b = 0, q = 0, lp = 0, sp = 0;
                for (int c = 1; c <= lastCol; c++)
                {
                    string h = N(ws.Cell(r, c).GetString());
                    if (h == "barkod" || h == "barcode") b = c;
                    if (h == "urunstokadedi" || h == "ürünstokadedi") q = c;
                    if (h.StartsWith("piyasasatisfiyatikdvdahil") || h == "listprice") lp = c;
                    if (h.StartsWith("trendyoldasatilacakfiyatkdvdahil") || h == "saleprice") sp = c;
                }
                if (b > 0 && q > 0) { headerRow = r; colBarcode = b; colQty = q; colList = lp; colSale = sp; return; }
            }
        }

        private static async Task<string> PostInventoryAsync(TrendyolCfg cfg, List<PostItem> items, CancellationToken ct)
        {
            var sb = new StringBuilder();
            sb.Append("{\"items\":[");
            bool first = true;
            foreach (var it in items)
            {
                if (!first) sb.Append(',');
                first = false;
                sb.Append("{\"barcode\":\"").Append(it.barcode).Append("\",\"quantity\":").Append(it.quantity);
                if (it.listPrice.HasValue) sb.Append(",\"listPrice\":").Append(it.listPrice.Value.ToString(CultureInfo.InvariantCulture));
                if (it.salePrice.HasValue) sb.Append(",\"salePrice\":").Append(it.salePrice.Value.ToString(CultureInfo.InvariantCulture));
                if (!string.IsNullOrEmpty(it.currencyType)) sb.Append(",\"currencyType\":\"").Append(it.currencyType).Append('\"');
                sb.Append('}');
            }
            sb.Append("]}");

            using (var handler = new HttpClientHandler { AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate })
            using (var http = new HttpClient(handler))
            {
                http.Timeout = TimeSpan.FromSeconds(cfg.TimeoutSeconds);
                http.DefaultRequestHeaders.Accept.ParseAdd("application/json");
                http.DefaultRequestHeaders.UserAgent.ParseAdd(cfg.UserAgent);
                string token = Convert.ToBase64String(Encoding.ASCII.GetBytes((cfg.ApiKey ?? "").Trim() + ":" + (cfg.ApiSecret ?? "").Trim()));
                http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", token);

                using (var content = new StringContent(sb.ToString(), Encoding.UTF8, "application/json"))
                using (var resp = await http.PostAsync(cfg.BaseUrl, content, ct))
                {
                    string body = await resp.Content.ReadAsStringAsync();
                    if (!resp.IsSuccessStatusCode)
                        throw new InvalidOperationException("Trendyol POST hata: " + (int)resp.StatusCode + " " + resp.ReasonPhrase + " | " + body);

                    var m = Regex.Match(body, "\"batchRequestId\"\\s*:\\s*\"([^\"]+)\"");
                    return m.Success ? m.Groups[1].Value : "";
                }
            }
        }

        private static IEnumerable<List<PostItem>> Chunk(List<PostItem> src, int size)
        {
            if (size <= 0) size = 100;
            for (int i = 0; i < src.Count; i += size)
                yield return src.Skip(i).Take(Math.Min(size, src.Count - i)).ToList();
        }

        // =============== Helpers ===============
        private static string GetVal(Dictionary<string, string> dict, string key) => dict.TryGetValue(key, out var v) ? v : "";
        private static string MergeCell(Dictionary<string, string> dict, string key, string val)
        { if (dict.TryGetValue(key, out var ex)) return ex.Contains(val) ? ex : ex + " | " + val; return val; }

        private static List<string> OrderColumns(HashSet<string> cols, HashSet<string> variantTypes, HashSet<string> stockLocs)
        {
            var order = new List<string>
            {
                "Product.Id","Product.Name","Brand.Name","Category.Names","Category.Ids",
                "Product.GoogleTaxonomyId","Product.SalesChannelIds","Product.Description",
                "Variant.Id","Variant.Sku","Variant.Barcodes",
                "Variant.Price.SellPrice","Variant.Price.DiscountPrice",
                "Variant.Stock.Total","Variant.Stock.Details",
                "Variant.Image.MainUrl","Variant.Image.Count","Variant.Image.AllUrls",
                "Variant.Description","Variant.VariantValues"
            };
            foreach (var t in variantTypes.OrderBy(x => x, StringComparer.Ordinal)) order.Add("Variant." + t);
            foreach (var s in stockLocs.OrderBy(x => x, StringComparer.Ordinal)) order.Add("Variant.Stock." + s);
            foreach (var c in cols.OrderBy(x => x, StringComparer.Ordinal)) if (!order.Contains(c)) order.Add(c);
            return order;
        }

        private static XElement GetChild(XElement parent, string localName) =>
            parent?.Elements().FirstOrDefault(e => e.Name.LocalName == localName);

        private static string GetChildValue(XElement parent, string localName)
        { var e = GetChild(parent, localName); return e != null ? (e.Value ?? "").Trim() : ""; }

       private static string NormalizeBarcode(string s)
{
    if (string.IsNullOrWhiteSpace(s)) return "";

    // sadece rakamları al
    string digits = new string(s.Where(char.IsDigit).ToArray());

    if (digits.Length == 0) return "";

    // 🔥 TSoft için EAN-13 zorla
    if (digits.Length < 13)
        digits = digits.PadLeft(13, '0');

    // Fazlaysa son 13 haneyi al
    if (digits.Length > 13)
        digits = digits.Substring(digits.Length - 13);

    return digits;
}


        private static string EnsureXlsx(string name) =>
            name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ? name : name + ".xlsx";

        public static string FindProjectRoot(string startDir)
        {
            var dir = new DirectoryInfo(startDir);
            for (int i = 0; i < 8 && dir != null; i++)
            {
                bool hasParams = Directory.Exists(Path.Combine(dir.FullName, "Parameters"));
                bool hasOutputs = Directory.Exists(Path.Combine(dir.FullName, "Outputs"));
                if (hasParams && hasOutputs) return dir.FullName;
                dir = dir.Parent;
            }
            return startDir;
        }

        private static Dictionary<string, Dictionary<string, string>> BuildRowByBarcode(
            List<Dictionary<string, string>> productRows)
        {
            var map = new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);
            foreach (var r in productRows)
            {
                var bars = (r.TryGetValue("Variant.Barcodes", out var s) ? s : "")
                            .Split('|').Select(b => new string(b.Where(char.IsDigit).ToArray()))
                            .Where(b => !string.IsNullOrEmpty(b)).Distinct();

                foreach (var b in bars)
                    if (!map.ContainsKey(b)) map[b] = r;
            }
            return map;
        }

        // ---------- Güvenli Products kaydet ----------
        // ---------- Güvenli Products kaydet (timestamp'lı isim) ----------
        // ---------- Güvenli Products kaydet (timestamp'lı, 'ikas' sözcüğünü siler) ----------
        private static string SaveProductsExcel(DataTable dt, string outputs, string stem)
        {
            Directory.CreateDirectory(outputs);

            // Temel ad: xml.txt 2. (veya 3.) satırdan; yoksa "Products"
            string baseNameRaw = string.IsNullOrWhiteSpace(stem) ? "Products" : SanitizeFileName(stem);

            // 'ikas' sözcüğünü tamamen kaldır, kalan gereksiz alt çizgi/boşlukları düzelt
            string baseName = System.Text.RegularExpressions.Regex.Replace(baseNameRaw, "ikas", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            baseName = System.Text.RegularExpressions.Regex.Replace(baseName, @"[_\-\.\s]+", "_").Trim('_', '-', '.', ' ');
            if (string.IsNullOrEmpty(baseName)) baseName = "Products";

            // Diğer çıktılarla uyumlu: <base>_yyyyMMdd_HHmmss.xlsx
            string fileName = $"{baseName}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            string path = Path.Combine(outputs, EnsureXlsx(fileName));

            if (dt.Columns.Count == 0)
            {
                dt.Columns.Add("Info", typeof(string));
                var dr = dt.NewRow();
                dr["Info"] = "XML boş/uyumsuz";
                dt.Rows.Add(dr);
            }

            try
            {
                if (File.Exists(path)) File.Delete(path);
            }
            catch (IOException)
            {
                path = Path.Combine(outputs, EnsureXlsx($"{baseName}_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.xlsx"));
            }

            using (var wb = new ClosedXML.Excel.XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Products");
                ws.Cell(1, 1).InsertTable(dt, true);
                ws.Columns().AdjustToContents();
                wb.SaveAs(path);
            }

            return path;
        }



        private static string SanitizeFileName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);
            foreach (var ch in name)
                sb.Append(invalid.Contains(ch) ? '_' : ch);

            string result = sb.ToString().Trim();
            if (result.Length == 0) result = "Products";
            return result.Length > 80 ? result.Substring(0, 80) : result;
        }

        private void UpdateCountdown()
        { TimeSpan ts = TimeSpan.FromSeconds(_remainingSeconds); lblCountdown.Text = ts.ToString(@"hh\:mm\:ss"); }
        private void UpdateProductCount(int n) => lblProducts.Text = n.ToString();
        private void Log(string s) => txtLog.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + s + Environment.NewLine);

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(282, 253);
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private static Dictionary<string, string> LoadBarcodeMapping(string eslemePath)
        {
            var map = new Dictionary<string, string>();

            if (!File.Exists(eslemePath))
                return map;

            using (var wb = new XLWorkbook(eslemePath))
            {
                var ws = wb.Worksheets.First();

                FindEslemeHeaders(ws, out int h, out int colEan, out int colXml, out _);

                int last = ws.LastRowUsed()?.RowNumber() ?? h;

                for (int r = h + 1; r <= last; r++)
                {
                    // XML EAN (869 / 8 ile başlar)
                    string xmlEAN = NormalizeBarcode(ws.Cell(r, colEan).GetString());

                    // TSoft Barcode (06 ile başlar)
                    string tsoftBarcode = NormalizeBarcode(ws.Cell(r, colXml).GetString());

                    if (xmlEAN.Length > 0 && tsoftBarcode.Length > 0)
                        map[xmlEAN] = tsoftBarcode;
                }
            }

            return map;
        }

        private sealed class TSoftCfg
        {
            public string Url;
            public string Token;
            public int TimeoutSeconds = 60;
        }

        private sealed class TSoftResponse
        {
            public bool? success;
            public List<TSoftMessage> message;
        }

        private sealed class TSoftMessage
        {
            public List<string> text;
        }

        private sealed class TSoftStockPayload
        {
            public string MainProductCode;
            public string SubProductCode;
            public string Stock;
            public string IsActive;
        }

        private const string DefaultTSoftSubProductUrl = "https://tangcarf.tsoft.biz/rest1/subProduct/setSubProducts";

        private static string NormalizeTSoftUrl(string rawUrl)
        {
            try
            {
                var input = (rawUrl ?? "").Trim();
                if (input.Length == 0) return DefaultTSoftSubProductUrl;

                input = input.Trim('"', '\'').TrimEnd(';');

                // Eğer kullanıcı tüm request satırını yapıştırdıysa, endpoint'i içinden ayıkla.
                var endpoint = Regex.Match(
                    input,
                    @"https?://[^/\s""']+/rest1/subProduct/setSubProducts",
                    RegexOptions.IgnoreCase);
                var anyUrl = Regex.Match(input, @"https?://[^\s""']+", RegexOptions.IgnoreCase);
                var candidate = endpoint.Success
                    ? endpoint.Value
                    : (anyUrl.Success
                        ? anyUrl.Value
                        : input);

                int q = candidate.IndexOf('?');
                if (q >= 0) candidate = candidate.Substring(0, q);

                if (!Uri.TryCreate(candidate, UriKind.Absolute, out var uri))
                    return DefaultTSoftSubProductUrl;

                var ub = new UriBuilder(uri)
                {
                    Scheme = Uri.UriSchemeHttps,
                    Port = -1,
                    Query = "",
                    Fragment = ""
                };

                var normalized = ub.Uri.GetLeftPart(UriPartial.Path).TrimEnd('/');
                return normalized.Length <= 512 ? normalized : DefaultTSoftSubProductUrl;
            }
            catch
            {
                return DefaultTSoftSubProductUrl;
            }
        }

        private static TSoftCfg LoadTSoftConfig(string baseDir)
        {
            string path = Path.Combine(baseDir, "Parameters", "tsoft.txt");
            if (!File.Exists(path))
                throw new InvalidOperationException("tsoft.txt bulunamadı");

            var lines = File.ReadAllLines(path, Encoding.UTF8)
                .Select(l => l.Trim())
                .Where(l => l.Length > 0 && !l.StartsWith("#") && !l.StartsWith("//"))
                .ToList();

            if (lines.Count == 0)
                throw new InvalidOperationException("tsoft.txt boş.");

            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in lines)
            {
                int idx = line.IndexOf('=');
                if (idx <= 0 || idx >= line.Length - 1) continue;
                string key = line.Substring(0, idx).Trim();
                string val = line.Substring(idx + 1).Trim();
                if (key.Length == 0 || val.Length == 0) continue;
                dict[key] = val;
            }

            string token = dict.TryGetValue("TOKEN", out var tok) ? tok : null;
            if (string.IsNullOrWhiteSpace(token) && lines.Count == 1 && !lines[0].Contains("="))
                token = lines[0];

            if (string.IsNullOrWhiteSpace(token))
                throw new InvalidOperationException("tsoft.txt içinde TOKEN bulunamadı.");

            string rawUrl = dict.TryGetValue("URL", out var u) ? u : DefaultTSoftSubProductUrl;
            string url = NormalizeTSoftUrl(rawUrl);
            int timeout = 60;
            if (dict.TryGetValue("TIMEOUT_SECONDS", out var t) && int.TryParse(t, out int parsed))
                timeout = Math.Max(10, parsed);

            return new TSoftCfg { Url = url, Token = token, TimeoutSeconds = timeout };
        }

        private static string BuildTSoftError(TSoftResponse resp)
        {
            var texts = resp?.message?
                .Where(m => m?.text != null)
                .SelectMany(m => m.text)
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .Select(t => t.Trim())
                .Distinct()
                .ToList();

            return (texts != null && texts.Count > 0) ? string.Join(" | ", texts) : "Bilinmeyen TSoft hatası";
        }

        private async Task SendVariantStocksToTSoftAsync(string productsExcelPath, CancellationToken ct)
        {
            string baseDir = FindProjectRoot(AppDomain.CurrentDomain.BaseDirectory);
            var cfg = LoadTSoftConfig(baseDir);
            Log($"TSOFT endpoint: {cfg.Url}");

            var stockByPair = new Dictionary<string, TSoftStockPayload>(StringComparer.OrdinalIgnoreCase);
            int skippedMainCode = 0;
            int skippedSubCode = 0;

            using (var wb = new XLWorkbook(productsExcelPath))
            {
                var ws = wb.Worksheets.First();

                // 🔥 HEADER BUL
                int headerRow = 1;
                int colMain = 0;
                int colSku = 0;
                int colStock = 0;

                int lastCol = ws.Row(headerRow).LastCellUsed()?.Address.ColumnNumber ?? 0;
                if (lastCol == 0)
                    throw new InvalidOperationException("Products Excel başlık satırı boş.");

                for (int c = 1; c <= lastCol; c++)
                {
                    string h = ws.Cell(headerRow, c).GetString().Trim();

                    if (h.Equals("Product.Id", StringComparison.OrdinalIgnoreCase) ||
                        h.Equals("MainProductCode", StringComparison.OrdinalIgnoreCase) ||
                        h.Equals("ProductCode", StringComparison.OrdinalIgnoreCase))
                        colMain = c;

                    if (h.Equals("Variant.Sku", StringComparison.OrdinalIgnoreCase))
                        colSku = c;

                    if (h.Equals("Variant.Stock.Total", StringComparison.OrdinalIgnoreCase))
                        colStock = c;
                }

                if (colMain == 0 || colSku == 0 || colStock == 0)
                    throw new Exception("Product.Id(MainProductCode), Variant.Sku veya Variant.Stock.Total sütunu bulunamadı");

                int lastRow = ws.LastRowUsed()?.RowNumber() ?? headerRow;

                for (int r = headerRow + 1; r <= lastRow; r++)
                {
                    string mainCode = ws.Cell(r, colMain).GetString().Trim();
                    string subCode = ws.Cell(r, colSku).GetString().Trim();

                    if (string.IsNullOrEmpty(mainCode))
                    {
                        skippedMainCode++;
                        continue;
                    }

                    if (string.IsNullOrEmpty(subCode))
                    {
                        skippedSubCode++;
                        continue;
                    }

                    int stock = 0;
                    string sStock = ws.Cell(r, colStock).GetString().Trim();
                    if (!int.TryParse(sStock, NumberStyles.Integer, CultureInfo.InvariantCulture, out stock))
                        int.TryParse(Convert.ToString(ws.Cell(r, colStock).Value, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out stock);

                    string key = mainCode + "||" + subCode;
                    stockByPair[key] = new TSoftStockPayload
                    {
                        MainProductCode = mainCode,
                        SubProductCode = subCode,
                        Stock = Math.Max(0, stock).ToString(CultureInfo.InvariantCulture),
                        IsActive = "1"
                    };
                }
            }

            var items = stockByPair.Values.ToList();

            Log($"TSOFT gönderilecek ürün: {items.Count} (MainProductCode boş atlanan: {skippedMainCode}, SubProductCode boş atlanan: {skippedSubCode})");

            string dataJson = JsonConvert.SerializeObject(items);

            var content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string,string>("token", cfg.Token),
                new KeyValuePair<string,string>("data", dataJson)
            });

            using (var http = new HttpClient())
            {
                http.Timeout = TimeSpan.FromSeconds(cfg.TimeoutSeconds);
                HttpResponseMessage resp;
                try
                {
                    resp = await http.PostAsync(cfg.Url, content, ct);
                }
                catch (UriFormatException)
                {
                    throw new InvalidOperationException("tsoft.txt URL hatalı. Sadece endpoint yazın: https://.../rest1/subProduct/setSubProducts");
                }

                string body = await resp.Content.ReadAsStringAsync();

                if (!resp.IsSuccessStatusCode)
                    throw new InvalidOperationException(body);

                TSoftResponse parsed = null;
                try { parsed = JsonConvert.DeserializeObject<TSoftResponse>(body); } catch { }

                if (parsed != null && parsed.success.HasValue && !parsed.success.Value)
                    throw new InvalidOperationException(BuildTSoftError(parsed));

                Log("TSOFT stok güncelleme başarılı.");
            }
        }



        private static string WriteTSoftExcel(
     List<Dictionary<string, string>> rows,
     string outputs)
        {
            string baseDir = FindProjectRoot(AppDomain.CurrentDomain.BaseDirectory);
            string eslemePath = Path.Combine(baseDir, "Parameters", "esleme.xlsx");

            // 🔥 XML EAN13 → 06 Barkod map
            var barcodeMap = new Dictionary<string, string>();

            if (File.Exists(eslemePath))
            {
                using (var wb = new XLWorkbook(eslemePath))
                {
                    var ws = wb.Worksheets.First();

                    FindEslemeHeaders(ws, out int h, out int colEan, out int colXml, out _);

                    int last = ws.LastRowUsed()?.RowNumber() ?? h;

                    for (int r = h + 1; r <= last; r++)
                    {
                        string trendyolBarcode = NormalizeBarcode(ws.Cell(r, colEan).GetString()); // 06
                        string xmlBarcode = NormalizeBarcode(ws.Cell(r, colXml).GetString());      // 869

                        if (!string.IsNullOrEmpty(xmlBarcode) &&
                            !string.IsNullOrEmpty(trendyolBarcode))
                        {
                            barcodeMap[xmlBarcode] = trendyolBarcode;
                        }
                    }
                }
            }

            string path = Path.Combine(outputs,
                "TSoft_Stok_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

            var dt = new DataTable("TSoft");

            dt.Columns.Add("ProductCode");
            dt.Columns.Add("Barcode"); // 🔥 SADECE 06 barkod
            dt.Columns.Add("Stock");

            foreach (var r in rows)
            {
                string productCode = GetVal(r, "Product.Id");
                string sku = GetVal(r, "Variant.Sku");
                int.TryParse(GetVal(r, "Variant.Stock.Total"), out int stock);

                var xmlBarcodes = GetVal(r, "Variant.Barcodes")
                    .Split('|')
                    .Select(b => NormalizeBarcode(b))
                    .Where(b => !string.IsNullOrEmpty(b));

                foreach (var xmlB in xmlBarcodes)
                {
                    // 🔥 eşleşmeden 06 barcode al
                    if (!barcodeMap.TryGetValue(xmlB, out string tsoftBarcode))
                        continue; // eşleşmeyenleri gönderme

                    var dr = dt.NewRow();
                    dr["ProductCode"] = string.IsNullOrWhiteSpace(productCode) ? sku : productCode;
                    dr["Barcode"] = tsoftBarcode; // 🔥 06 barkod
                    dr["Stock"] = stock;

                    dt.Rows.Add(dr);
                }
            }

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Products");
                ws.Cell(1, 1).InsertTable(dt, true);

                ws.Column(2).Style.NumberFormat.Format = "@";

                ws.Columns().AdjustToContents();
                wb.SaveAs(path);
            }

            return path;
        }


    }
}
