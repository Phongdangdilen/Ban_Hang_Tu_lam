using System.ComponentModel.DataAnnotations;

namespace Ontap_Net104_319.Models
{
    public class Account
    {
        // Data Anotation Validation được sử dụng để thực hiện validate dữ liệu ngay từ model
        [Required] //  Bắt buộc phải có 
        [StringLength(450, MinimumLength = 10, ErrorMessage = "Độ dài phải từ 10 đến 450 kí tự")]
        public string Username { get; set; }
        [Required]
        [StringLength(450, MinimumLength = 10, ErrorMessage = "Độ dài phải từ 10 đến 450 kí tự")]
        public string Password { get; set; }
        [RegularExpression("^(\\+\\d{1,2}\\s)?\\(?\\d{3}\\)?[\\s.-]\\d{3}[\\s.-]\\d{4}$",
            ErrorMessage = "Số điện thoại phải đúng format và có 10 chữ số")] // xxx-xxx-xxxx
        public string PhoneNumber { get; set; }
        [EmailAddress(ErrorMessage = "Email chưa đúng định dạng")]
        public string Email { get; set; }
        public string Address { get; set; }
        public virtual Cart? Cart { get; set; }
        public virtual List<Bill>? Bills { get; set; }
    }
}
