using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;

namespace Ontap_Net104_319.Controllers
{
    public class LocationController : Controller
    {
        HttpClient _client;
        public LocationController()
        {
            _client = new HttpClient();
        }
        // GET: LocationController
        public ActionResult GetProvince()
        {
            string requestURL = "https://online-gateway.ghn.vn/shiip/public-api/master-data/province";
            _client.DefaultRequestHeaders.Add("Token", "11e923e7-257d-11ef-ad6a-e6aec6d1ae72");
            var response = _client.GetStringAsync(requestURL).Result;
            JObject data = JObject.Parse(response);
            var pokemonData = new JObject
            {
                ["data"] = data["data"],
            };
            return View(pokemonData);
        }

        public ActionResult GetDistrict(int id)
        {

            if (id != 0)
            {
                HttpContext.Session.SetString("DistrictUser", id.ToString());
                string requestURL = $"https://online-gateway.ghn.vn/shiip/public-api/master-data/district?province_id={id}";
                _client.DefaultRequestHeaders.Add("Token", "11e923e7-257d-11ef-ad6a-e6aec6d1ae72");
                var response = _client.GetStringAsync(requestURL).Result;
                JObject data = JObject.Parse(response);
                Console.WriteLine(id);
                var pokemonData = new JObject
                {
                    ["data"] = data["data"],
                };
                return View(pokemonData);
            }
            else
            {
                return BadRequest();
            }
        }

        public ActionResult GetWard(int id)
        {
            if (id != 0)
            {
                var DistrictUser = HttpContext.Session.GetString("DistrictUser");
                string requestURL = $"https://online-gateway.ghn.vn/shiip/public-api/master-data/ward?district_id={id}";
                _client.DefaultRequestHeaders.Add("Token", "11e923e7-257d-11ef-ad6a-e6aec6d1ae72");
                var response = _client.GetStringAsync(requestURL).Result;
                JObject data = JObject.Parse(response);
                Console.WriteLine(id);
                var pokemonData = new JObject
                {
                    ["data"] = data["data"],
                };
                return View(pokemonData);
            }
            else
            {
                return BadRequest();
            }
        }
        public ActionResult Total(int id)
        {
            if (id != 0)
            {

                var DistrictUser = HttpContext.Session.GetString("DistrictUser");
                string requestURL = $"https://online-gateway.ghn.vn/shiip/public-api/v2/shipping-order/fee?service_id=53321&insurance_value=500000&from_district_id=1804&to_district_id={DistrictUser}&to_ward_code={id}&height=15&length=15&weight=1000&width=15";
                _client.DefaultRequestHeaders.Add("Token", "11e923e7-257d-11ef-ad6a-e6aec6d1ae72");
                var response = _client.GetStringAsync(requestURL).Result;
                JObject data = JObject.Parse(response);
                var a = Convert.ToDecimal(data["data"]["total"]).ToString();
                HttpContext.Session.SetString("phi", a);
                return RedirectToAction("Pay", "Cart");
            }
            else
            {
                return BadRequest();
            }
        }


    }
}
