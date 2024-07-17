using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ModelBinding.Binders;
using ShoppingList.Models;
using ShoppingList.Services;
using HtmlAgilityPack;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Net;
using System.Text;
using System.IO;
using Azure;
using System.Xml;
using System;


namespace ShoppingList.Controllers
{
    public class ProductsController : Controller
    {
        private readonly ApplicationDbContext context;
        private readonly IWebHostEnvironment environment;

        public ProductsController(ApplicationDbContext context, IWebHostEnvironment environment)
        {
            this.context = context;
            this.environment = environment;
        }
        public IActionResult Index()
        {
            var products = context.Products.OrderByDescending(p => p.Id).ToList();
            return View(products);
        }

        public IActionResult Create()
        {
            return View();
        }


        // submit button for creating new product
        [HttpPost]
        public IActionResult Create(ProductDto productDto)
        {
            if (productDto.ImageFile == null)
            {
                ModelState.AddModelError("ImageFile", "The image file is required.");
            }

            if (!ModelState.IsValid)
            {
                return View(productDto);
            }

            // save img
            string newFileName = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            newFileName += Path.GetExtension(productDto.ImageFile!.FileName);
            string fullPath = environment.WebRootPath + "/pictures/" + newFileName;

            using (var stream = System.IO.File.Create(fullPath))
            {
                productDto.ImageFile.CopyTo(stream);
            }

            // save to db
            Product product = new Product()
            {
                Name = productDto.Name,
                Brand = productDto.Brand,
                Category = productDto.Category,
                Price = productDto.Price,
                Description = productDto.Description,
                ImageFileName = newFileName,
                Created = DateTime.Now
            };

            context.Products.Add(product);
            context.SaveChanges();

            return RedirectToAction("Index", "Products");




        }




        public IActionResult Edit(int id)
        {
            var product = context.Products.Find(id);

            if (product == null)
            {
                return RedirectToAction("Index", "Products");
            }

            // create productDto from product
            var productDto = new ProductDto()
            {
                Name = product.Name,
                Brand = product.Brand,
                Category = product.Category,
                Price = product.Price,
                Description = product.Description
            };

            ViewData["ProductId"] = product.Id;
            ViewData["ImageFileName"] = product.ImageFileName;
            ViewData["Created"] = product.Created.ToString("MM/dd/yyyy");


            return View(productDto);
        }

        [HttpPost]
        public IActionResult Edit(int id, ProductDto productDto)
        {
            var product = context.Products.Find(id);

            if (product == null)
            {
                return RedirectToAction("Index", "Products");
            }

            if (!ModelState.IsValid)
            {
                ViewData["ProductId"] = product.Id;
                ViewData["ImageFileName"] = product.ImageFileName;
                ViewData["Created"] = product.Created.ToString("MM/dd/yyyy");

                return View(productDto);
            }

            // update img, if changed
            string newFileName = product.ImageFileName;
            if (productDto.ImageFile != null)
            {
                newFileName = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                newFileName += Path.GetExtension(productDto.ImageFile.FileName);

                string fullPath = environment.WebRootPath + "/pictures/" + newFileName;

                using (var stream = System.IO.File.Create(fullPath))
                {
                    productDto.ImageFile.CopyTo(stream);
                }

                // delete old img
                string oldPath = environment.WebRootPath + "/pictures/" + product.ImageFileName;
                System.IO.File.Delete(oldPath);

            }
            // update product in DB
            product.Name = productDto.Name;
            product.Brand = productDto.Brand;
            product.Category = productDto.Category;
            product.Price = productDto.Price;
            product.Description = productDto.Description;
            product.ImageFileName = newFileName;

            context.SaveChanges();

            return RedirectToAction("Index", "Products");

        }


        public IActionResult Delete(int id)
        {
            var product = context.Products.Find(id);

            if (product == null)
            {
                return RedirectToAction("Index", "Products");
            }

            // delete old img
            string oldPath = environment.WebRootPath + "/pictures/" + product.ImageFileName;
            System.IO.File.Delete(oldPath);

            context.Products.Remove(product);
            context.SaveChanges();
            return RedirectToAction("Index", "Products");
        }
        public IActionResult LinkToProduct()
        {
            return View();
        }
        [HttpPost]
        public IActionResult LinkToProduct(LinkToProduct lnk)
        {
            if (!ModelState.IsValid)
            {
                return View(lnk);
            }

            // if amazon link provided, scrape and attempt to auto fill.
            WebScrapeAmazon(lnk);

            return RedirectToAction("Index", "Products");
        }
        public void WebScrapeAmazon(LinkToProduct linkProduct)
        {
            if (!linkProduct.Link.StartsWith("https://www.amazon"))
            {
                return;
            }

            HttpClientHandler handler = new HttpClientHandler()
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };
            var client = new HttpClient(handler);
            var response = CallUrl(client, linkProduct.Link);

            List<string> parsedData = ParseHtmlElemAsync(response.Result);


            // price
            decimal priceConverted = Convert.ToDecimal(parsedData[2]);


            Product product = new Product()
            {
                Name = parsedData[0],
                Brand = parsedData[1],
                Category = linkProduct.Category,
                Price = priceConverted,
                Description = linkProduct.Description,
                ImageFileName = parsedData[3],
                Created = DateTime.Now
            };
            context.Products.Add(product);
            context.SaveChanges();

        }

        private static async Task<string> CallUrl(HttpClient client, string fullUrl)
        {
            var response = await client.GetStringAsync(fullUrl);
            return response;
        }
        private List<string> ParseHtmlElemAsync(string html)
        {
            // scrape relevant data
            List<string> data = new List<string>();
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);

            var name = htmlDoc.DocumentNode.SelectSingleNode("//span[contains(concat(' ', normalize-space(@id), ' '), 'productTitle')]");
            var brand = htmlDoc.DocumentNode.SelectSingleNode("//span[contains(concat(' ', normalize-space(@class), ' '), 'a-size-base po-break-word')]");
            var price = htmlDoc.DocumentNode.SelectSingleNode("//span[contains(concat(' ', normalize-space(@class), ' '), 'a-offscreen')]");
            var imgLink = htmlDoc.DocumentNode.SelectSingleNode("//img[contains(concat(' ', normalize-space(@id), ' '), 'landingImage')]");
            string src = imgLink.GetAttributeValue("src", string.Empty);

            // download img 
            var fileName = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string fullPath = environment.WebRootPath + "/pictures/" + fileName + ".jpg";

            using (WebClient client = new WebClient())
            {
                client.DownloadFile(new Uri(src), fullPath);

            }



            data.Add(name.InnerText.Trim());
            data.Add(brand.InnerText);
            data.Add(price.InnerText.Substring(1, price.InnerText.Length - 1));
            data.Add(fileName + ".jpg");

            return data;
           

            
        }

    }






}
