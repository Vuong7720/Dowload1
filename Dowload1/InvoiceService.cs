using Newtonsoft.Json;
using OpenAI.Chat;
using OpenCvSharp;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.ClientModel;
using System.IO.Compression;
using System.Net;
using System.Net.Sockets;
using Titanium.Web.Proxy;
using Titanium.Web.Proxy.EventArguments;
using Titanium.Web.Proxy.Models;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Tesseract;

namespace Dowload1
{
    public class InvoiceService : IInvoiceService
    {
        // Đã hoàn thành download file có cả xml và pdf
        public async Task<(byte[] fileData, string fileName)> DownloadInvoice(string url, string invoiceCode)
        {
            string requestId = Guid.NewGuid().ToString();
            string rootDownloadPath = Path.Combine(Directory.GetCurrentDirectory(), "Downloads");
            string userDownloadPath = Path.Combine(rootDownloadPath, requestId);
            if (!Directory.Exists(userDownloadPath)) Directory.CreateDirectory(userDownloadPath);

            ChromeOptions options = new ChromeOptions();

            //options.AddArgument("--headless=new");

            // Cấu hình tải file & In ấn
            options.AddUserProfilePreference("download.default_directory", userDownloadPath);
            options.AddUserProfilePreference("download.prompt_for_download", false);

            options.AddUserProfilePreference("download.directory_upgrade", true);
            // Thêm vào cùng chỗ với các options khác
            options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);

            // QUAN TRỌNG: Tắt các tính năng chặn tải file của Chrome
            options.AddUserProfilePreference("safebrowsing.enabled", true);
            options.AddUserProfilePreference("safebrowsing.disable_download_protection", true);
            options.AddArgument("--safebrowsing-disable-download-protection");
            options.AddArgument("--allow-running-insecure-content");

            options.AddArgument("--kiosk-printing"); // Quan trọng nhất để tự động lưu PDF
            options.AddUserProfilePreference("printing.print_preview_sticky_settings",
                "{\"recentDestinations\":[{\"id\":\"Save as PDF\",\"origin\":\"local\",\"account\":\"\"}],\"selectedDestinationId\":\"Save as PDF\",\"version\":2}");

            // Start a local HTTP(S) proxy to capture PDF responses from the browser.
            var proxyPort = new Random().Next(20000, 30000);
            var proxyServer = new ProxyServer();
            var explicitEndPoint = new ExplicitProxyEndPoint(System.Net.IPAddress.Loopback, proxyPort, true);
            proxyServer.AddEndPoint(explicitEndPoint);
            proxyServer.Start();

            // Configure Chrome to use the proxy
            options.AddArgument($"--proxy-server=127.0.0.1:{proxyPort}");
            // Ignore certificate errors for automation (will allow proxy MITM without trusting cert)
            options.AddArgument("--ignore-certificate-errors");

            // Hook response event to capture PDFs
            proxyServer.BeforeResponse += async (object sender, SessionEventArgs e) =>
            {
                try
                {
                    var headers = e.HttpClient.Response.Headers;
                    string contentType = string.Empty;
                    foreach (var h in headers)
                    {
                        if (string.Equals(h.Name, "Content-Type", StringComparison.OrdinalIgnoreCase))
                        {
                            contentType = h.Value ?? string.Empty;
                            break;
                        }
                    }

                    if (!string.IsNullOrEmpty(contentType) && contentType.ToLowerInvariant().Contains("application/pdf"))
                    {
                        try
                        {
                            var bodyBytes = await e.GetResponseBody();
                            if (bodyBytes != null && bodyBytes.Length > 0)
                            {
                                var pdfPath = Path.Combine(userDownloadPath, $"Invoice_{invoiceCode}_proxy.pdf");
                                await File.WriteAllBytesAsync(pdfPath, bodyBytes);
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            };

            using (var driver = new ChromeDriver(options))
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                bool loginSuccess = false;

                for (int i = 0; i < 10; i++)
                {
                    driver.Navigate().GoToUrl(url);

                    // Nhập thông tin (Mã hóa đơn & Captcha)
                    var inputMaso = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("QueryExtender_txtText")));
                    inputMaso.Clear();
                    inputMaso.SendKeys(invoiceCode);
                     
                    var imgElement = driver.FindElement(By.Id("QueryExtender_Image"));
                    string base64Data = imgElement.GetAttribute("src").Split(',')[1];

                    string captchaText = SolveCaptchaLocalModelAI(base64Data);
                    // Lưu ý: Cần làm sạch captchaText một lần nữa trong code C#
                    captchaText = Regex.Replace(captchaText, @"[^a-zA-Z0-9]", "");

                    var inputCaptcha = driver.FindElement(By.Id("QueryExtender_txtCode"));
                    inputCaptcha.SendKeys(captchaText);
                    driver.FindElement(By.Id("QueryExtender_Ok")).Click();

                    await Task.Delay(3000); // Chờ trang load kết quả

                    try
                    {
                        // Kiểm tra sự xuất hiện của nút download để xác nhận đăng nhập thành công
                        wait.Until(ExpectedConditions.ElementIsVisible(By.Id("QueryExtender_btnDownload")));
                        loginSuccess = true;
                        break;
                    }
                    catch
                    {
                        Console.WriteLine($"Sai captcha lần {i + 1}, đang thử lại...");
                    }
                }

                if (!loginSuccess) return (null, "FAILED_LOGIN");

                try
                {
                    // --- BẮT ĐẦU TẢI ---

                    // 1. Tải XML
                    var btnXml = driver.FindElement(By.Id("QueryExtender_btnDownload"));
                    driver.ExecuteScript("arguments[0].click();", btnXml);
                    await Task.Delay(2000);

                    // 2. Tải PDF (Dùng Script để đảm bảo focus vào trang trước khi in)
                    //string pdfUrl = ((IJavaScriptExecutor)driver).ExecuteScript(@"
                    //                const embed = document.querySelector('embed[original-url]');
                    //                return embed ? embed.getAttribute('original-url') : null;
                    //            ") as string;

                    //if (string.IsNullOrEmpty(pdfUrl))
                    //    throw new Exception("Không tìm thấy original-url của PDF");
                    //driver.ExecuteScript($@"
                    //        var link = document.createElement('a');
                    //        link.href = '{pdfUrl}';
                    //        link.download = 'Invoice_{invoiceCode}.pdf';
                    //        document.body.appendChild(link);
                    //        link.click();
                    //        document.body.removeChild(link);
                    //    ");

                    // Chờ một chút để Chrome xử lý tải file
                    //await Task.Delay(3000);


                    // NOTE: Disabled performance.getEntries -> apiUrl path because it commonly returns null/blob and
                    // causes fallback to printToPDF (screenshot-like PDF). We'll rely on other fallbacks (blob fetch / Playwright).
                    string apiUrl = null;

                    var cookieContainer = new System.Net.CookieContainer();
                    var seleniumCookies = driver.Manage().Cookies.AllCookies;
                    foreach (var cookie in seleniumCookies)
                    {
                        cookieContainer.Add(new System.Net.Cookie(cookie.Name, cookie.Value, cookie.Path, cookie.Domain));
                    }

                    bool pdfSaved = false;

                    // Directly try Playwright fallback (skip other fallbacks)
                    try
                    {
                        var pwPdf = await TryDownloadPdfWithPlaywrightAsync(url, invoiceCode, cookieContainer);
                        if (pwPdf != null && pwPdf.Length > 0)
                        {
                            string pdfFilePath = Path.Combine(userDownloadPath, $"Invoice_{invoiceCode}.pdf");
                            await File.WriteAllBytesAsync(pdfFilePath, pwPdf);
                            pdfSaved = true;
                            Console.WriteLine($"Lưu PDF gốc bằng Playwright: {pwPdf.Length} bytes");
                        }
                        else
                        {
                            Console.WriteLine("Playwright không lấy được PDF gốc.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Lỗi Playwright fallback: " + ex.Message);
                    }



                    // 3. Vòng lặp chờ file (Thông minh hơn)
                    int timeout = 5;
                    while (timeout > 0)
                    {
                        var files = Directory.GetFiles(userDownloadPath);
                        // Lọc bỏ file .crdownload và kiểm tra số lượng
                        var completedFiles = files.Where(f => !f.EndsWith(".crdownload") && !f.EndsWith(".tmp")).ToList();

                        if (completedFiles.Count >= 2) break; // Đã đủ XML và PDF

                        await Task.Delay(1000);
                        timeout--;
                    }

                    // Đóng trình duyệt sớm để giải phóng handle file
                    driver.Quit();

                    // Stop proxy and dispose
                    try
                    {
                        proxyServer.BeforeResponse -= null; // best-effort detach
                        proxyServer.Stop();
                        proxyServer.Dispose();
                    }
                    catch { }

                    // 4. Nén và trả về dữ liệu
                    string zipPath = Path.Combine(rootDownloadPath, $"Invoice_{invoiceCode}.zip");
                    if (File.Exists(zipPath)) File.Delete(zipPath); // Xóa bản cũ nếu có

                    ZipFile.CreateFromDirectory(userDownloadPath, zipPath);
                    byte[] archiveBytes = await File.ReadAllBytesAsync(zipPath);

                    // Dọn dẹp
                    Directory.Delete(userDownloadPath, true);
                    File.Delete(zipPath);

                    return (archiveBytes, $"Invoice_{invoiceCode}.zip");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Lỗi thực thi: " + ex.Message);
                    return (null, null);
                }
            }
        }

        // Fallback using Playwright to intercept the PDF response and save the original PDF bytes
        private async Task<byte[]?> TryDownloadPdfWithPlaywrightAsync(string pageUrl, string invoiceCode, CookieContainer cookieContainer)
        {
            try
            {
                using var playwright = await Microsoft.Playwright.Playwright.CreateAsync();
                var browser = await playwright.Chromium.LaunchAsync(new Microsoft.Playwright.BrowserTypeLaunchOptions { Headless = true });
                var context = await browser.NewContextAsync();

                // Transfer cookies from CookieContainer to Playwright context (for the target page domain)
                try
                {
                    var uri = new Uri(pageUrl);
                    var coll = cookieContainer.GetCookies(uri);
                    var pwCookies = new List<Microsoft.Playwright.Cookie>();
                    foreach (System.Net.Cookie c in coll)
                    {
                        pwCookies.Add(new Microsoft.Playwright.Cookie { Name = c.Name, Value = c.Value, Domain = c.Domain, Path = c.Path, HttpOnly = c.HttpOnly, Secure = c.Secure });
                    }
                    if (pwCookies.Count > 0)
                    {
                        await context.AddCookiesAsync(pwCookies.ToArray());
                    }
                }
                catch { }

                // Navigate and attempt to reproduce the download action so we can capture the network response with PDF bytes
                var page = await context.NewPageAsync();

                // Ensure page navigation after cookies set
                await page.GotoAsync(pageUrl, new Microsoft.Playwright.PageGotoOptions { WaitUntil = Microsoft.Playwright.WaitUntilState.NetworkIdle, Timeout = 30000 });

                // Prepare network log file
                var logPath = Path.Combine(Directory.GetCurrentDirectory(), "playwright_netlog.txt");
                try { File.WriteAllText(logPath, $"Playwright network log started at {DateTime.UtcNow:o}\n"); } catch {}

                // Try to wait for a PDF response that matches the invoice PDF request
                Microsoft.Playwright.IResponse? pdfResponse = null;
                var waitTask = page.WaitForResponseAsync(resp =>
                    {
                        var u = (resp.Url ?? string.Empty).ToLowerInvariant();
                        var headers = resp.Headers ?? new Dictionary<string, string>();
                        var ct = headers.ContainsKey("content-type") ? headers["content-type"] : (headers.GetValueOrDefault("Content-Type") ?? string.Empty);
                        return (ct != null && ct.ToLowerInvariant().Contains("application/pdf")) || u.Contains("einvoicequery.ashx") || u.Contains("t=4") || u.EndsWith(".pdf");
                    }, new Microsoft.Playwright.PageWaitForResponseOptions { Timeout = 30000 });

                // Log every network response for diagnosis
                page.Response += async (sender, resp) =>
                {
                    try
                    {
                        var urlLower = (resp.Url ?? string.Empty);
                        var headers = resp.Headers ?? new Dictionary<string, string>();
                        var ct = headers.ContainsKey("content-type") ? headers["content-type"] : (headers.GetValueOrDefault("Content-Type") ?? string.Empty);
                        var status = resp.Status;
                        var lenHeader = headers.ContainsKey("content-length") ? headers["content-length"] : string.Empty;
                        var line = $"{DateTime.UtcNow:o} URL={urlLower} Status={status} Content-Type={ct} Content-LengthHeader={lenHeader}" + Environment.NewLine;
                        await File.AppendAllTextAsync(logPath, line);
                    }
                    catch { }
                };

                // Attempt to perform the site's form submission (including captcha) inside Playwright
                try
                {
                    for (int attempt = 0; attempt < 6; attempt++)
                    {
                        try
                        {
                            var inputMaso = await page.QuerySelectorAsync("#QueryExtender_txtText");
                            if (inputMaso == null)
                            {
                                await File.AppendAllTextAsync(logPath, $"Attempt {attempt + 1}: input field not found\n");
                                await page.WaitForTimeoutAsync(1000);
                                continue;
                            }

                            await page.FillAsync("#QueryExtender_txtText", invoiceCode);

                            var imgEl = await page.QuerySelectorAsync("#QueryExtender_Image");
                            if (imgEl == null)
                            {
                                await File.AppendAllTextAsync(logPath, $"Attempt {attempt + 1}: captcha image not found\n");
                                await page.WaitForTimeoutAsync(1000);
                                continue;
                            }

                            var src = await imgEl.GetAttributeAsync("src");
                            if (string.IsNullOrEmpty(src) || !src.Contains(','))
                            {
                                await File.AppendAllTextAsync(logPath, $"Attempt {attempt + 1}: captcha src missing or invalid\n");
                                await page.WaitForTimeoutAsync(1000);
                                continue;
                            }

                            var base64Part = src.Split(',')[1];
                            var captchaText = SolveCaptchaLocalModelAI(base64Part);
                            captchaText = Regex.Replace(captchaText ?? string.Empty, "[^a-zA-Z0-9]", "");

                            await page.FillAsync("#QueryExtender_txtCode", captchaText);
                            await page.ClickAsync("#QueryExtender_Ok");

                            // Wait briefly for page to update
                            await page.WaitForTimeoutAsync(2000);

                            // Check if download button visible
                            try
                            {
                                var downloadEl = await page.QuerySelectorAsync("#QueryExtender_btnDownload");
                                if (downloadEl != null)
                                {
                                    bool vis = false;
                                    try { vis = await downloadEl.IsVisibleAsync(); } catch { }
                                    if (vis)
                                    {
                                        await File.AppendAllTextAsync(logPath, $"Playwright login success on attempt {attempt + 1}\n");
                                        break;
                                    }
                                }
                            }
                            catch { }
                        }
                        catch (Exception ex)
                        {
                            try { await File.AppendAllTextAsync(logPath, $"Login attempt {attempt + 1} error: {ex.Message}\n"); } catch { }
                        }
                    }

                    // Now try to click the download button to produce network response
                    try
                    {
                        var btn = await page.QuerySelectorAsync("#QueryExtender_btnDownload");
                        if (btn != null) await btn.ClickAsync(new Microsoft.Playwright.ElementHandleClickOptions { Timeout = 10000 });
                        else await page.EvaluateAsync("document.querySelector('#QueryExtender_btnDownload')?.click()");
                    }
                    catch (Exception ex)
                    {
                        try { await File.AppendAllTextAsync(logPath, "Error clicking download button: " + ex.Message + "\n"); } catch { }
                    }

                    try
                    {
                        pdfResponse = await page.WaitForResponseAsync(resp =>
                        {
                            var u = (resp.Url ?? string.Empty).ToLowerInvariant();
                            var headers = resp.Headers ?? new Dictionary<string, string>();
                            var ct = headers.ContainsKey("content-type") ? headers["content-type"] : (headers.GetValueOrDefault("Content-Type") ?? string.Empty);
                            return (ct != null && ct.ToLowerInvariant().Contains("application/pdf")) || u.Contains("einvoicequery.ashx") || u.Contains("t=4") || u.EndsWith(".pdf");
                        }, new Microsoft.Playwright.PageWaitForResponseOptions { Timeout = 45000 });
                    }
                    catch { }
                }
                catch (Exception ex)
                {
                    try { await File.AppendAllTextAsync(logPath, "Playwright login+click flow error: " + ex.Message + "\n"); } catch { }
                }

                if (pdfResponse != null)
                {
                    try
                    {
                        var body = await pdfResponse.BodyAsync();
                        await File.AppendAllTextAsync(logPath, $"Captured PDF response URL={pdfResponse.Url} Status={pdfResponse.Status} Length={body?.LongLength}\n");
                        await browser.CloseAsync();
                        return body;
                    }
                    catch (Exception ex)
                    {
                        try { await File.AppendAllTextAsync(logPath, "Error reading PDF response body: " + ex.Message + "\n"); } catch { }
                    }
                }

                // If we didn't capture a network PDF, try to inspect the page for embed/iframe/object/data or blob URLs
                try
                {
                    // This script looks for embed/iframe/object elements and tries to return a base64 string
                    string fetchPdfScript = @"(async () => {
                        try {
                            function toBase64(buf) {
                                let binary = '';
                                const bytes = new Uint8Array(buf);
                                const chunk = 0x8000;
                                for (let i = 0; i < bytes.length; i += chunk) {
                                    binary += String.fromCharCode.apply(null, Array.prototype.slice.call(bytes, i, Math.min(i + chunk, bytes.length)));
                                }
                                return btoa(binary);
                            }

                            // Candidate selectors
                            const candidates = [];
                            document.querySelectorAll('embed, iframe, object').forEach(e => candidates.push(e));
                            // anchors that might directly link to pdf
                            document.querySelectorAll('a').forEach(a => { if (a.href) candidates.push(a); });

                            for (const e of candidates) {
                                try {
                                    const src = e.src || e.getAttribute('data') || e.getAttribute('data-src') || e.getAttribute('original-url') || e.href || '';
                                    if (!src) continue;

                                    // data:application/pdf;base64,...
                                    if (src.startsWith('data:')) {
                                        const parts = src.split(',');
                                        if (parts.length > 1) return parts[1];
                                    }

                                    // blob: URL - fetch it in-page
                                    if (src.startsWith('blob:')) {
                                        try {
                                            const r = await fetch(src);
                                            const ab = await r.arrayBuffer();
                                            return toBase64(ab);
                                        } catch (e) { }
                                    }

                                    // regular URL - try fetch if it looks like a pdf link
                                    if (src.toLowerCase().includes('.pdf') || src.toLowerCase().includes('t=4') || src.toLowerCase().includes('einvoice')) {
                                        try {
                                            const r = await fetch(src, { credentials: 'include' });
                                            const ct = r.headers.get('content-type') || '';
                                            if (ct.includes('application/pdf') || src.toLowerCase().includes('.pdf')) {
                                                const ab = await r.arrayBuffer();
                                                return toBase64(ab);
                                            }
                                        } catch (e) { }
                                    }
                                } catch (e) { }
                            }

                            // If PDF generated client-side into a canvas, try to find canvas and export as PDF-ish image (fallback)
                            const canv = document.querySelector('canvas');
                            if (canv) {
                                try {
                                    const dataUrl = canv.toDataURL('image/png');
                                    const b64 = dataUrl.split(',')[1];
                                    return b64;
                                } catch (e) { }
                            }

                            return null;
                        } catch (err) { return null; }
                    })();";

                    var maybeBase64 = await page.EvaluateAsync<string>(fetchPdfScript);
                    if (!string.IsNullOrEmpty(maybeBase64))
                    {
                        try
                        {
                            var bytes = Convert.FromBase64String(maybeBase64);
                            await File.AppendAllTextAsync(logPath, $"Captured PDF via in-page fetch, bytes={bytes.Length}\n");
                            await browser.CloseAsync();
                            return bytes;
                        }
                        catch (Exception ex)
                        {
                            try { await File.AppendAllTextAsync(logPath, "Error decoding base64 from page: " + ex.Message + "\n"); } catch { }
                        }
                    }
                    else
                    {
                        try { await File.AppendAllTextAsync(logPath, "No in-page PDF/base64 found after clicks. Saving page html for analysis.\n"); } catch { }
                        try { var html = await page.ContentAsync(); await File.WriteAllTextAsync(Path.Combine(Directory.GetCurrentDirectory(), $"playwright_page_{invoiceCode}.html"), html); } catch { }
                    }
                }
                catch (Exception ex)
                {
                    try { await File.AppendAllTextAsync(logPath, "Error during in-page PDF extraction: " + ex.Message + "\n"); } catch { }
                }

                await browser.CloseAsync();
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Playwright fallback error: " + ex.Message);
                return null;
            }
        }


        // giải captcha
        private string SolveCaptchaLocalModelAI(string base64Data)
        {
            try
            {
                // 1. Khởi tạo Credential (bắt buộc cho thư viện OpenAI mới)
                ApiKeyCredential credential = new ApiKeyCredential("lm-studio");

                // 2. Cấu hình Endpoint trỏ về LM Studio
                var options = new OpenAI.OpenAIClientOptions
                {
                    Endpoint = new Uri("http://localhost:1234/v1")
                };

                // 3. Khởi tạo ChatClient với đúng ID từ ảnh của bạn
                // Lưu ý: Copy chính xác identifier từ mục "API Usage" trong ảnh
                var client = new ChatClient("mistralai/ministral-3-3b", credential, options);

                // 4. Xử lý dữ liệu ảnh
                byte[] imageBytes = Convert.FromBase64String(base64Data);
                BinaryData imageBinary = BinaryData.FromBytes(imageBytes, "image/png");

                // 5. Tạo Prompt tối ưu cho việc giải Captcha
                List<ChatMessage> messages = new List<ChatMessage>
        {
            new UserChatMessage(
                ChatMessageContentPart.CreateTextPart("OCR this image. Return only the text/characters found, no talking."),
                ChatMessageContentPart.CreateImagePart(imageBinary, "image/png")
            )
        };

                // 6. Gọi API (Dùng CompleteChat nếu muốn đợi kết quả ngay)
                ChatCompletion completion = client.CompleteChat(messages);

                string captchaResult = completion.Content[0].Text.Trim();

                // In ra console để bạn kiểm tra model có đọc đúng không
                Console.WriteLine($"[AI OCR]: {captchaResult}");

                return captchaResult;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi kết nối LM Studio: {ex.Message}");
                return string.Empty;
            }
        }
    }
}
