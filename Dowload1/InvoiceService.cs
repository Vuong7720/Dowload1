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

            options.AddArgument("--headless=new");

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
            var proxyServer = new ProxyServer();

            var explicitEndPoint = new ExplicitProxyEndPoint(System.Net.IPAddress.Loopback, 0, true);
            proxyServer.AddEndPoint(explicitEndPoint);
            proxyServer.Start();

            // 2. Lấy cổng thực tế SAU KHI đã Start
            var actualPort = explicitEndPoint.Port;

            // 3. Truyền cổng đó vào Chrome
            options.AddArgument($"--proxy-server=127.0.0.1:{actualPort}");

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

                    //string captchaText = SolveCaptchaLocalModelAI(base64Data);
                    string captchaText = SolveCaptchaWithQwen2VL(base64Data);
                    // Lưu ý: Cần làm sạch captchaText một lần nữa trong code C#
                    captchaText = Regex.Replace(captchaText, @"[^a-zA-Z0-9]", "");

                    var inputCaptcha = driver.FindElement(By.Id("QueryExtender_txtCode"));
                    inputCaptcha.SendKeys(captchaText);
                    driver.FindElement(By.Id("QueryExtender_Ok")).Click();

                    await Task.Delay(3000); // Chờ trang load kết quả

                    try
                    {
                        wait.Until(ExpectedConditions.ElementIsVisible(By.Id("QueryExtender_btnDownload")));
                        loginSuccess = true;
                        break;
                    }
                    catch
                    {
                        Console.WriteLine($"Sai captcha lan {i + 1}, dang thu lai...");
                    }
                }

                if (!loginSuccess) return (null, "FAILED_LOGIN");

                try
                {
                    // --- BẮT ĐẦU TẢI ---

                    // 1. Tải XML
                    var btnXml = driver.FindElement(By.Id("QueryExtender_btnDownload"));
                    driver.ExecuteScript("arguments[0].click();", btnXml);
                    //await Task.Delay(2000);

                    // 2. Tải PDF (Dùng Script để đảm bảo focus vào trang trước khi in)
                   
                    string apiUrl = null;

                    var cookieContainer = new System.Net.CookieContainer();
                    var seleniumCookies = driver.Manage().Cookies.AllCookies;
                    foreach (var cookie in seleniumCookies)
                    {
                        cookieContainer.Add(new System.Net.Cookie(cookie.Name, cookie.Value, cookie.Path, cookie.Domain));
                    }

                    bool pdfSaved = false;

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

        // giải captcha với ministral 3-3b
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

        // giải captcha với qwen2-vl-2b-instruct:2 với tốc độ nhanh hơn so với model ministral
        private string SolveCaptchaWithQwen2VL(string base64Data)
        {
            try
            {
                ApiKeyCredential credential = new ApiKeyCredential("lm-studio");
                var options = new OpenAI.OpenAIClientOptions { Endpoint = new Uri("http://localhost:1234/v1") };

                // 1. Thay đổi ID tương ứng với Model Qwen bạn vừa tải
                // Lưu ý: Kiểm tra ID chính xác trong tab "Local Server" của LM Studio
                var client = new ChatClient("lmstudio-community/Qwen2-VL-2B-Instruct-GGUF", credential, options);

                byte[] imageBytes = Convert.FromBase64String(base64Data);
                BinaryData imageBinary = BinaryData.FromBytes(imageBytes, "image/png");

                // 2. Prompt tối ưu cho Qwen2-VL để giải mã captcha nhiễu
                List<ChatMessage> messages = new List<ChatMessage>
        {
            new UserChatMessage(
                ChatMessageContentPart.CreateTextPart("Read the characters in this image. Return ONLY the text found."),
                ChatMessageContentPart.CreateImagePart(imageBinary, "image/png")
            )
        };

                // 3. Thiết lập Temperature thấp (0 hoặc 0.1) để giảm việc đoán sai ký tự
                ChatCompletionOptions completionOptions = new ChatCompletionOptions()
                {
                    Temperature = 0.1f,
                    MaxOutputTokenCount = 10 // Captcha thường rất ngắn, giới hạn để tiết kiệm tài nguyên
                };

                ChatCompletion completion = client.CompleteChat(messages, completionOptions);

                // Làm sạch kết quả (loại bỏ khoảng trắng hoặc xuống dòng thừa)
                string captchaResult = completion.Content[0].Text.Trim();

                Console.WriteLine($"[Qwen2-VL OCR]: {captchaResult}");
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
