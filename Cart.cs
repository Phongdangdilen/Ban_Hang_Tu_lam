namespace Ontap_Net104_319.Models
{
    public class Cart
    {
        public string Username { get; set; }
        public int Status { get; set; }
        public virtual Account? Account{ get; set; } 
        public virtual List<CartDetails>? Details { get; set; }
    }
}
