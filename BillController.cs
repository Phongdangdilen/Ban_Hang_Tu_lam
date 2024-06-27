using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using NuGet.Protocol.Plugins;
using Ontap_Net104_319.Models;
using System.Xml.Schema;

namespace Ontap_Net104_319.Controllers
{
	public class BillController : Controller
	{
		HttpClient _client;

		AppDbContext _context;
		public BillController()
		{
			_client = new HttpClient();
			_context = new AppDbContext();
		}
        // GET: BillCOntroller
        public IActionResult ExportToExcel()
        {
            var username = HttpContext.Session.GetString("username") ?? "Guest";
            var bills = _context.Bills
                .Where(b => b.Username == username)
                .OrderByDescending(b => b.CreateDate)
                .ToList();

            if (bills == null || !bills.Any())
            {
                return Content("No bills available to export.");
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Bills");

                worksheet.Cell(1, 1).Value = "Bill ID";
                worksheet.Cell(1, 2).Value = "Create Date";
                worksheet.Cell(1, 3).Value = "Product Name";
                worksheet.Cell(1, 4).Value = "Price";
                worksheet.Cell(1, 5).Value = "Quantity";
                worksheet.Cell(1, 6).Value = "Total";
                worksheet.Cell(1, 7).Value = "Status";

                int currentRow = 2;
                decimal totalSum = 0;

                foreach (var bill in bills)
                {
                    foreach (var detail in bill.Details)
                    {
                        var product = _context.Products.FirstOrDefault(p => p.Id == detail.ProductId);
                        if (product == null)
                        {
                            continue; // Bỏ qua nếu không tìm thấy sản phẩm
                        }

                        worksheet.Cell(currentRow, 1).Value = bill.Id;
                        worksheet.Cell(currentRow, 2).Value = bill.CreateDate;
                        worksheet.Cell(currentRow, 3).Value = product.Name ?? "Unknown";
                        worksheet.Cell(currentRow, 4).Value = detail.ProductPrice;
                        worksheet.Cell(currentRow, 5).Value = detail.Quantity;
                        worksheet.Cell(currentRow, 6).Value = detail.ProductPrice * detail.Quantity;
                        worksheet.Cell(currentRow, 7).Value = bill.Status;

                        totalSum += detail.ProductPrice * detail.Quantity;
                        currentRow++;
                    }
                }

                worksheet.Cell(currentRow, 5).Value = "Total Sum";
                worksheet.Cell(currentRow, 6).Value = totalSum;

                var range = worksheet.Range(1, 1, currentRow, 7);
                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Bills.xlsx");
                }
            }
        }


            public IActionResult Index()
		{
			if (HttpContext.Session.GetString("username") != null)
			{
				string requesstURl = $"https://localhost:7001/api/Bill/get-bill?username={HttpContext.Session.GetString("username")}";
				var response = _client.GetStringAsync(requesstURl).Result;
				var data = JsonConvert.DeserializeObject<List<Bill>>(response);
				return View(data);
			}
			return RedirectToAction("Login", "Account");
		}

		public IActionResult CancelBill(string username, string id)
		{
			if (HttpContext.Session.GetString("username") != null)
			{
				string requesstURl = $"https://localhost:7001/api/Bill/Cancel-bill?username={HttpContext.Session.GetString("username")}&id={id}";
				var response = _client.PutAsJsonAsync(requesstURl, requesstURl).Result;
				return RedirectToAction("Index", "Bill");
			}
			return BadRequest("Lỗi");
		}

		public IActionResult RepurchaseBill(string username, string id)
		{
			var bill = _context.Bills.Include(b => b.Details).FirstOrDefault(b => b.Id == id && b.Username == HttpContext.Session.GetString("username"));
			if (HttpContext.Session.GetString("username") != null)
			{
				foreach (var item in bill.Details)
				{
					var billDetails = _context.BillDetails.Include(a => a.Product).FirstOrDefault(a => a.Id == item.Id);
					if (billDetails.Product.Amount < billDetails.Quantity || billDetails.Product.Amount == 0)
					{
						TempData["Invalid"] = $"{billDetails.Product.Name}";
						return RedirectToAction("Index", "Bill");
					}
				}
				string requesstURl = $"https://localhost:7001/api/Bill/repurchase-bill?username={HttpContext.Session.GetString("username")}&id={id}";
				var response = _client.PutAsJsonAsync(requesstURl, requesstURl).Result;
				return RedirectToAction("Index", "cart");
			}
			return BadRequest("Lỗi");
		}
	}

}
