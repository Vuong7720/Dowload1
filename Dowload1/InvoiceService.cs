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
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Tesseract;

namespace Dowload1
{
    public class InvoiceService : IInvoiceService
    {
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

            using (var driver = new ChromeDriver(options))
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                bool loginSuccess = false;

                for (int i = 0; i < 3; i++)
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

        private async Task DownloadPdfHttpAsync(
      IWebDriver driver,
      string pdfUrl,
      string saveFolder,
      string invoiceCode
  )
        {
            var handler = new HttpClientHandler
            {
                UseCookies = true,
                CookieContainer = new CookieContainer(),
                AutomaticDecompression = DecompressionMethods.All
            };

            var pdfUri = new Uri(pdfUrl);

            // Copy cookie
            foreach (var c in driver.Manage().Cookies.AllCookies)
            {
                handler.CookieContainer.Add(
                    pdfUri,
                    new System.Net.Cookie(c.Name, c.Value, c.Path, c.Domain)
                );
            }

            using var http = new HttpClient(handler);

            // ⚠️ HEADER BẮT BUỘC
            http.DefaultRequestHeaders.Add("User-Agent",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/143.0");
            http.DefaultRequestHeaders.Add("Accept",
                "application/pdf,application/octet-stream;q=0.9,*/*;q=0.8");
            http.DefaultRequestHeaders.Add("Referer",
                "https://einvoice.fast.com.vn/");
            http.DefaultRequestHeaders.Add("Accept-Language", "vi-VN,vi;q=0.9");

            // ⚠️ NHIỀU SERVER PDF CẦN RANGE
            http.DefaultRequestHeaders.Range =
                new System.Net.Http.Headers.RangeHeaderValue(0, null);

            var response = await http.GetAsync(pdfUri);

            var pdfBytes = await response.Content.ReadAsByteArrayAsync();

            Console.WriteLine($"[PDF] Status = {response.StatusCode}");
            Console.WriteLine($"[PDF] ContentType = {response.Content.Headers.ContentType}");
            Console.WriteLine($"[PDF] Size = {pdfBytes.Length}");

            // Validate PDF
            if (!response.IsSuccessStatusCode ||
                pdfBytes.Length < 1000 ||
                Encoding.ASCII.GetString(pdfBytes.Take(4).ToArray()) != "%PDF")
            {
                // Dump lỗi để debug
                var dump = Encoding.UTF8.GetString(pdfBytes.Take(500).ToArray());
                Console.WriteLine("[PDF INVALID CONTENT]");
                Console.WriteLine(dump);

                throw new Exception("Tải PDF HTTP thất bại hoặc nội dung không hợp lệ");
            }

            string pdfPath = Path.Combine(saveFolder, $"Invoice_{invoiceCode}.pdf");
            await File.WriteAllBytesAsync(pdfPath, pdfBytes);

            Console.WriteLine($"[PDF] Saved OK: {pdfPath}");
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
