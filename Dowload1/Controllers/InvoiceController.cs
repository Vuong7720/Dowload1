using Dowload1.requests;
using Microsoft.AspNetCore.Mvc;

namespace Dowload1.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class InvoiceController : ControllerBase
    {
        private readonly IInvoiceService _invoiceService;

        public InvoiceController(IInvoiceService invoiceService)
        {
            _invoiceService = invoiceService;
        }

        [HttpPost("download")]
        public async Task<IActionResult> GetInvoice([FromBody] DownloadInvoiceRequest request)
        {
            var result = await _invoiceService.DownloadInvoice(request.link, request.code);

            if (result.fileData == null)
            {
                return BadRequest(new { message = "Không thể tải hóa đơn. Vui lòng kiểm tra mã hoặc captcha." });
            }

            // Trả về file Zip chứa cả XML và PDF cho client
            return File(result.fileData, "application/zip", result.fileName);
        }
    }
}
