using Microsoft.AspNetCore.Mvc;
using ShoppingList.Models;
using ShoppingList.Services;
using HtmlAgilityPack;
using System.Net;
using Microsoft.EntityFrameworkCore;
using System.Data;
using ClosedXML.Excel;



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

            // if link is 'null' change to empty string
            if (productDto.Link == null)
            {
                productDto.Link = productDto.Link ?? string.Empty;
                ModelState.Remove("Link");
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
                Link = productDto.Link,
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
                Description = product.Description,
                Link = product.Link
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
            product.Link = productDto.Link;

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
            bool result = WebScrapeAmazon(lnk);

            if (!result)
            {
                return View(lnk);
            }
            return RedirectToAction("Index", "Products");
        }
        public bool WebScrapeAmazon(LinkToProduct linkProduct)
        {
            if (!linkProduct.Link.StartsWith(Constants.AMAZON_LINK))
            {
                ModelState.AddModelError("Link", "Invalid Link");
                return false;
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

            // insert to db
            Product product = new Product()
            {
                Name = parsedData[0],
                Brand = parsedData[1],
                Category = linkProduct.Category,
                Price = priceConverted,
                Description = linkProduct.Description,
                ImageFileName = parsedData[3],
                Link = linkProduct.Link,
                Created = DateTime.Now
            };
            context.Products.Add(product);
            context.SaveChanges();

            return true;

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
            var brand = htmlDoc.DocumentNode.SelectSingleNode("//tr[@class='a-spacing-small po-brand']//span[@class='a-size-base po-break-word']");
            var priceWhole = htmlDoc.DocumentNode.SelectSingleNode("//span[@class='a-price aok-align-center reinventPricePriceToPayMargin priceToPay']//span[@aria-hidden='true']/span[contains(concat(' ', normalize-space(@class), ' '), 'a-price-whole')]");
            var priceFraction = htmlDoc.DocumentNode.SelectSingleNode("//span[@class='a-price aok-align-center reinventPricePriceToPayMargin priceToPay']//span[@aria-hidden='true']/span[contains(concat(' ', normalize-space(@class), ' '), 'a-price-fraction')]");
            var imgLink = htmlDoc.DocumentNode.SelectSingleNode("//img[contains(concat(' ', normalize-space(@id), ' '), 'landingImage')]");
            string src = imgLink.GetAttributeValue("src", string.Empty);

            var fullPrice = (priceWhole.InnerText + priceFraction.InnerText).Trim();

            // download img 
            var fileName = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string fullPath = environment.WebRootPath + "/pictures/" + fileName + ".jpg";

            using (WebClient client = new WebClient())
            {
                client.DownloadFile(new Uri(src), fullPath);

            }



            data.Add(name.InnerText.Trim());
            data.Add(brand.InnerText);
            data.Add(fullPrice);
            data.Add(fileName + ".jpg");

            return data;



        }

        // database to excel sheet
        [HttpGet]
        public async Task<FileResult> ToExcel()
        {
            var toList = await context.Products.ToListAsync();
            var fileName = "ShoppingList.xlsx";

            DataTable table = new DataTable("Products");
            // add excel column names using Product DB columns
            table.Columns.AddRange(new DataColumn[]
            {
                new DataColumn("Id"),
                new DataColumn("Name"),
                new DataColumn("Brand"),
                new DataColumn("Category"),
                new DataColumn("Price"),
                new DataColumn("Image"),
                new DataColumn("Description"),
                new DataColumn("Link"),
                new DataColumn("Created")
            });


            // add each item to table
            foreach (var item in toList)
            {
                // note image file name is just a placeholder, as it will turn into a picture using ClosedXML later
                table.Rows.Add(item.Id, item.Name, item.Brand, item.Category, item.Price, "", item.Description, item.Link, item.Created);

            }

            // export to excel
            using (XLWorkbook wb = new XLWorkbook())
            {
                var worksheet = wb.Worksheets.Add(table, "Sheet1");

                // expand columns to fit text
                worksheet.Columns().AdjustToContents();

                // first row is reversed for column name
                int startRow = 2;
                int startColumn = 6;
                foreach (var item in toList)
                {
                    var currentCell = worksheet.Cell(startRow, startColumn);
                    string fullImagePath = environment.WebRootPath + "/pictures/" + item.ImageFileName;
                    worksheet.Column(startColumn).Width = Constants.EXCEL_IMAGE_WIDTH;
                    worksheet.Row(startRow).Height = Constants.EXCEL_IMAGE_HEIGHT;
                    var pic = worksheet.AddPicture(fullImagePath).MoveTo(currentCell, currentCell.CellBelow().CellRight());
                    pic.Placement = ClosedXML.Excel.Drawings.XLPicturePlacement.MoveAndSize;

                    ++startRow;
                }



                // download via web
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);

                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileName);



                }


            }


        }




    }

}
