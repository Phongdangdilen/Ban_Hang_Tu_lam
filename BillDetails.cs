namespace Ontap_Net104_319.Models
{
    public class BillDetails // là thông tin về 1 sản phẩm trong hóa đơn
    {
        public Guid Id { get; set; }
        public Guid ProductId { get; set; }
        public string BillId { get; set; }
        public decimal ProductPrice { get; set; }
        public int Quantity { get; set; }
        public int Status { get; set; }
        public virtual Product? Product { get; set; }
        public virtual Bill? Bill { get; set; }
    }
}
