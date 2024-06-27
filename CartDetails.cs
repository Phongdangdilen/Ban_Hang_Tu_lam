namespace Ontap_Net104_319.Models
{
    public class CartDetails
    {
        public Guid Id { get; set; }
        public Guid ProductId { get; set; }
        public string Username { get; set; }
        public int Quantity { get; set; }
        public int Status { get; set; }
        public virtual Cart? Cart { get; set; }
        public virtual Product? Product { get; set; }
    }
}
