using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Ontap_Net104_319.Models;
using System.Net.Http;

namespace Ontap_Net104_319.Controllers
{
    public class AccountController : Controller
    {
        AppDbContext context; // Khai báo context và tạo constructor
        HttpClient client;
        private readonly IHttpClientFactory _clientFactory;
        public AccountController(IHttpClientFactory httpClientFactory)
        {
            context = new AppDbContext(); // khỏi tạo context
            client = new HttpClient();
            _clientFactory = httpClientFactory;
        }

        public IActionResult Login(string username, string password)
        {

            string requesstURl = $"https://localhost:7001/api/Account/login?username={username}&password={password}";
            var response = client.PostAsJsonAsync(requesstURl, requesstURl).Result;
            if (response.IsSuccessStatusCode)
            {
                HttpContext.Session.SetString("username", username);
                return RedirectToAction("Index", "Home");
            }
            else return View();

        }

        public IActionResult SignUp() // Action này đơn thuần để mở View cần thực hiện
        {
            return View();
        }
        [HttpPost]
        public IActionResult SignUp(Account account) // Action này thực hiện việc tạo ra tài khoản mới
        {
            
            string requesstURl = "https://localhost:7001/api/Account/SignUp";
            var response = client.PostAsJsonAsync(requesstURl, account).Result;
            if (response.IsSuccessStatusCode)
            {
                return RedirectToAction("Index", "Home");
            }
            return BadRequest();
        }
        public IActionResult LogOut()
        {
            HttpContext.Session.Remove("username"); // Xóa dữ liệu của username đã login
            return RedirectToAction("Index", "Home");
        }
    }
}
