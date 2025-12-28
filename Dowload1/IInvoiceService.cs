namespace Dowload1
{
    public interface IInvoiceService
    {
        Task<(byte[] fileData, string fileName)> DownloadInvoice(string url, string invoiceCode);
    }
}
