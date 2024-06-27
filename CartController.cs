using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Ontap_Net104_319.Models;

namespace Ontap_Net104_319.Controllers
{
	public class CartController : Controller
	{
		HttpClient _client;
		AppDbContext _context;
		public CartController()
		{
			_client = new HttpClient();
			_context = new AppDbContext();
		}
		public IActionResult Index() // Hiển thị tất cả danh sách các sản phẩm có trong giỏ hàng của 1 user
		{
			if (HttpContext.Session.GetString("username") != null)
			{
				string requesstURl = $"https://localhost:7001/api/Cart/get-cart?username={HttpContext.Session.GetString("username")}";
				var response = _client.GetStringAsync(requesstURl).Result;
				var data = JsonConvert.DeserializeObject<List<CartDetails>>(response);
				return View(data);
			}
			return RedirectToAction("Login", "Account");
		}


		public IActionResult Pay()
		{
			var p = HttpContext.Session.GetString("pvc");
			var data = _context.CartDetails.Include(a => a.Product).Include(a => a.Cart).Where(a => a.Cart.Username == HttpContext.Session.GetString("username")).ToList();
			if (HttpContext.Session.GetString("username") != null)
			{
				string requesstURl = $"https://localhost:7001/api/Cart/pay/{(p)}";
				var response = _client.PutAsJsonAsync(requesstURl, HttpContext.Session.GetString("username")).Result;
				if (!response.IsSuccessStatusCode)
				{
					foreach (var item in data)
					{
						var cartDetails = _context.CartDetails.Include(a => a.Product).FirstOrDefault(a => a.Id == item.Id);
						if (cartDetails.Product.Amount < cartDetails.Quantity || cartDetails.Product.Amount == 0)
						{
							TempData["Invalid"] = $"{cartDetails.Product.Name}";
							return RedirectToAction("Index", "Cart");

						}
					}
				}
				return RedirectToAction("Index", "Bill");
			}
			return View();
		}


		public IActionResult Delete(Guid id)
		{
			string requesstURl = $"https://localhost:7001/api/Cart/delete-by-id?id={id}";
			var response = _client.DeleteAsync(requesstURl).Result;
			return RedirectToAction("Index", "Cart");

		}
	}
}
