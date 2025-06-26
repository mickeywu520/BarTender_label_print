namespace LabelPrintManager.Models
{
    /// <summary>
    /// 進貨單據資料模型
    /// </summary>
    public class PurchaseReceiptModel
    {
        public string DocType { get; set; }
        public string DocNumber { get; set; }
        public string DocItem { get; set; }
        public string PurchaseDate { get; set; }
        public string SupplierCode { get; set; }
        public string SupplierName { get; set; }
        public string ProductCode { get; set; }
        public string ProductName { get; set; }
        public string Specification { get; set; }
        public decimal Quantity { get; set; }
    }
}
