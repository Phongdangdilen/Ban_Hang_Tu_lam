using Microsoft.AspNetCore.Http.Metadata;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Newtonsoft.Json;
using Ontap_Net104_319.Models;
using System.Collections.Generic;

namespace Ontap_Net104_319.Controllers
{
	public class ProductController : Controller
	{
		HttpClient _client;
		AppDbContext _context;
		public ProductController()
		{
			_client = new HttpClient();
			_context = new AppDbContext();
		}
		public IActionResult Index()
		{
			string requesstURl = "https://localhost:7001/api/Product/get-all";
			var response = _client.GetStringAsync(requesstURl).Result;
			var data = JsonConvert.DeserializeObject<List<Product>>(response);
			return View(data);

		}

		public ActionResult Details(Guid id)
		{
			string requesstURl = $"https://localhost:7001/api/Product/get-by-id?id={id}";
			var response = _client.GetStringAsync(requesstURl).Result;
			var data = JsonConvert.DeserializeObject<Product>(response);
			return View(data);
		}
		public IActionResult Create()
		{
			return View();
		}
		[HttpPost]
		public IActionResult Create(Product product)
		{
			string requesstURl = "https://localhost:7001/api/Product/create-product";
			var response = _client.PostAsJsonAsync(requesstURl, product).Result;
			return RedirectToAction("Index");
		}
		public IActionResult Edit(Guid id)
		{
			string requesstURl = $"https://localhost:7001/api/Product/get-by-id?id={id}";
			var response = _client.GetStringAsync(requesstURl).Result;
			var data = JsonConvert.DeserializeObject<Product>(response);
			return View(data);
		}
		[HttpPost]
		public IActionResult Edit(Product product)
		{
			string requesstURl = "https://localhost:7001/api/Product/update-product";
			var response = _client.PutAsJsonAsync(requesstURl, product).Result;
			return RedirectToAction("Index");
		}
		public IActionResult Delete(Guid id)
		{
			string requesstURl = $"https://localhost:7001/api/Product/delete-by-id?id={id}";
			var response = _client.DeleteAsync(requesstURl).Result;
			return RedirectToAction("Index");
		}
		public IActionResult AddtoCart(Guid id, int quantity)
		{
			var username = HttpContext.Session.GetString("username");
			var cartItem = _context.CartDetails.FirstOrDefault(p => p.ProductId == id && p.Username == username);
			var productItem = _context.Products.FirstOrDefault(p => p.Id == id);
			if (username != null)
			{
				if (cartItem != null)
				{
					if (cartItem.Quantity + quantity > productItem.Amount)
					{
						TempData["Invalid"] = $"{cartItem.Product.Name}";
						return RedirectToAction("Index", "Cart");
					}
				}
				string requesstURl = $"https://localhost:7001/api/Product/Add-to-Cart?username={username}&id={id}&quantity={quantity}";
				var response = _client.PutAsJsonAsync(requesstURl, requesstURl).Result;
				return RedirectToAction("Index", "Cart");

			}
			return BadRequest();
		}
	}
}
