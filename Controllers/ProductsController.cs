using Microsoft.AspNetCore.Mvc;
using ShoppingList.Models;
using ShoppingList.Services;
using HtmlAgilityPack;
using System.Net;
using Microsoft.EntityFrameworkCore;
using System.Data;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Irony.Parsing;
using Microsoft.IdentityModel.Tokens;
using System.Web;



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
        /// <summary>
        /// Returns the database to a list type, so it is viewable in Index.cshtml.
        /// </summary>
        /// <returns></returns>
        public IActionResult Index()
        {
            var products = context.Products.ToList();
            return View(products);
        }

        public IActionResult Create()
        {
            return View();
        }


        /// <summary>
        /// Inserts an entry to the Products database through the inputs provided in Create.cshtml
        /// Invalid inputs will refresh the page with error messages below affected input box.
        /// This function will run when the Submit button is pressed in Create.cshtml
        /// </summary>
        /// <param name="productDto"> The inputs provided by the user: name, price, desc, image file etc. </param>
        /// <returns>IActionResult which is used to change/update pages depending if inputs are valid </returns>
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
                Created = DateTime.Now,
                PriceDifferenceSymbol = "•"

            };

            context.Products.Add(product);
            context.SaveChanges();

            return RedirectToAction("Index", "Products");




        }




        /// <summary>
        /// Provides product data to Edit.cshtml.
        /// </summary>
        /// <param name="id"> The ID of the product (primary key) </param>
        /// <returns> Refreshes Index.cshtml if invalid, else adds relevant data to ViewData and goes to Edit.cshtml </returns>
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

        /// <summary>
        /// Updates database from data provided by productDto
        /// Also replaces image, so index.cshtml displays updated image.
        /// </summary>
        /// <param name="id">The primary key of the product, provided by the other Edit function </param>
        /// <param name="productDto">The product information, provided by the other Edit function.</param>
        /// <returns> Goes back to Index.cshtml, and doesn't change info if there's an error. </returns>
        [HttpPost]
        public IActionResult Edit(int id, ProductDto productDto)
        {
            var product = context.Products.Find(id);

            if (product == null)
            {
                return RedirectToAction("Index", "Products");
            }
            // if link is 'null' change to empty string
            if (productDto.Link == null)
            {
                productDto.Link = productDto.Link ?? string.Empty;
                ModelState.Remove("Link");
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
            product.PriceDifferenceSymbol = GetPriceDifferenceSymbol(product.Price, productDto.Price);
            product.Price = productDto.Price;
            product.Description = productDto.Description;
            product.ImageFileName = newFileName;
            product.Link = productDto.Link;


            context.SaveChanges();

            return RedirectToAction("Index", "Products");

        }


        /// <summary>
        /// Delete a database entry, including the linked image in wwwroot/pictures
        /// </summary>
        /// <param name="id">ID of product (primary key) </param>
        /// <returns>
        /// Refreshes Index.cshtml and can see if it is successful
        /// if it is deleted, else ID wasn't found in DB </returns>
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

        /// <summary>
        /// Creates a database entry automatically thru an Amazon link.
        /// </summary>
        /// <param name="lnk"> Data from the LinkToProduct class provided via user input that cannot be automatically filled </param>
        /// <returns> Back to Index.cshtml if successful, else refreshes page with errors under the affected input box </returns>
        [HttpPost]
        public IActionResult LinkToProduct(LinkToProduct lnk)
        {
            if (!ModelState.IsValid)
            {
                return View(lnk);
            }
            if (!lnk.Link.StartsWith(Constants.AMAZON_LINK) && lnk.Link.Length > Constants.AMAZON_LINK_LENGTH_REQUIREMENT)
            {
                ModelState.AddModelError("Link", "Invalid Link");
                return View(lnk);
            }

            // if amazon link provided, scrape and attempt to auto fill.
            List<string> result = WebScrapeAmazon(lnk.Link);
            if (result.IsNullOrEmpty())
            {
                return View(lnk);
            }

            decimal priceConverted = Convert.ToDecimal(result[2]);


            // save to db
            ParsedAmazonToDB(lnk, result, priceConverted);;

            return RedirectToAction("Index", "Products");
        }

        /// <summary>
        /// Web scrapes Amazon to get name, price, image and brand.
        /// </summary>
        /// <param name="linkProduct"> The provided user input data </param>
        /// <returns> List of parsed data if link is valid, else empty list </returns>
        public List<string> WebScrapeAmazon(string link, ACTION_TYPE dataToReturn = ACTION_TYPE.ALL)
        {
            List<string> parsedData;

            HttpClientHandler handler = new HttpClientHandler()
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };
            var client = new HttpClient(handler);
            var response = CallUrl(client, link);

            // parse relevant data to list (price, name, etc)
            parsedData = ParseHtmlElem(response.Result, dataToReturn);


            return parsedData;

        }

        /// <summary>
        /// Save the parsed Amazon data to DB
        /// </summary>
        /// <param name="linkProduct"> The provided user input data from LinkToProduct.cshtml </param>
        /// <param name="parsedData"> The parsed Amazon data, result from ParsedHtmlElem() </param>
        /// <param name="priceConverted"> Parsed amazon data price string -> decimal </param>
        /// <returns> True, else exception will occur automatically </returns>
        private bool ParsedAmazonToDB(LinkToProduct linkProduct, List<string> parsedData, decimal priceConverted)
        {
            // insert to db
            Product product = new Product()
            {
                // decodes symbols like &amp; to &
                Name = HttpUtility.HtmlDecode(parsedData[0]),
                Brand = parsedData[1],
                Category = linkProduct.Category,
                Price = priceConverted,
                Description = linkProduct.Description,
                ImageFileName = parsedData[3],
                Link = linkProduct.Link,
                Created = DateTime.Now,
                PriceDifferenceSymbol = "•"
            };
            context.Products.Add(product);
            context.SaveChanges();

            return true;
        }


        /// <summary>
        /// Attempts send a GET request to link, and returns result.
        /// </summary>
        /// <param name="client"> HTTPClient which decompresses gzip, which Amazon has (else returns gibberish HTML string) </param>
        /// <param name="fullUrl"> Amazon URL </param>
        /// <returns> Should return the HTML data of the link </returns>
        /// <exception cref="HttpRequestException">If response fails either thru invalid url, network issues or server error.</exception>
        private static async Task<string> CallUrl(HttpClient client, string fullUrl)
        {
            var response = await client.GetStringAsync(fullUrl);
            return response;
        }

        /// <summary>
        /// Extracts name, brand, price and image link from Amazon product via XPath queries.
        /// Also downloads the image to wwwroot/pictures, with fileName supplied to the list for image link.
        /// </summary>
        /// <param name="html"> A string with the HTML of the Amazon product </param>
        /// <returns> A string list of name, brand, price and image link </returns>
        private List<string> ParseHtmlElem(string html, ACTION_TYPE dataToReturn = ACTION_TYPE.ALL)
        {
            // scrape relevant data
            List<string> data = new List<string>();
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);

            if (dataToReturn == ACTION_TYPE.ALL)
            {
                // hand pick the relevant elements using XPath
                var name = htmlDoc.DocumentNode.SelectSingleNode("//span[contains(concat(' ', normalize-space(@id), ' '), 'productTitle')]");
                var brand = htmlDoc.DocumentNode.SelectSingleNode("//tr[@class='a-spacing-small po-brand']//span[@class='a-size-base po-break-word']");
                var priceWhole = htmlDoc.DocumentNode.SelectSingleNode("//span[@class='a-price aok-align-center reinventPricePriceToPayMargin priceToPay']//span[@aria-hidden='true']/span[contains(concat(' ', normalize-space(@class), ' '), 'a-price-whole')]");
                var priceFraction = htmlDoc.DocumentNode.SelectSingleNode("//span[@class='a-price aok-align-center reinventPricePriceToPayMargin priceToPay']//span[@aria-hidden='true']/span[contains(concat(' ', normalize-space(@class), ' '), 'a-price-fraction')]");
                var imgLink = htmlDoc.DocumentNode.SelectSingleNode("//img[contains(concat(' ', normalize-space(@id), ' '), 'landingImage')]");
                string src = imgLink.GetAttributeValue("src", string.Empty);

                // combine the price together, as Amazon splits these HTML elements.
                string fullPrice = (priceWhole.InnerText + priceFraction.InnerText).Trim();

                // download img with a filename related to time.
                var fileName = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                string fullPath = environment.WebRootPath + "/pictures/" + fileName + ".jpg";

                using (WebClient client = new WebClient())
                {
                    client.DownloadFile(new Uri(src), fullPath);

                }


                // add parsed HTML elements to List<string>, ensuring no whitespace.
                data.Add(name.InnerText.Trim());
                data.Add(brand.InnerText);
                data.Add(fullPrice);
                data.Add(fileName + ".jpg");
            }
            else if (dataToReturn == ACTION_TYPE.PRICE)
            {
                var priceWhole = htmlDoc.DocumentNode.SelectSingleNode("//span[@class='a-price aok-align-center reinventPricePriceToPayMargin priceToPay']//span[@aria-hidden='true']/span[contains(concat(' ', normalize-space(@class), ' '), 'a-price-whole')]");
                var priceFraction = htmlDoc.DocumentNode.SelectSingleNode("//span[@class='a-price aok-align-center reinventPricePriceToPayMargin priceToPay']//span[@aria-hidden='true']/span[contains(concat(' ', normalize-space(@class), ' '), 'a-price-fraction')]");
                string fullPrice = (priceWhole.InnerText + priceFraction.InnerText).Trim();
                data.Add(fullPrice);
            }
            

            return data;



        }

        /// <summary>
        /// Converts Products SQL Server database to Excel.
        /// This function runs when the Excel button is pressed.
        /// This should download the Excel file in your browser.
        /// </summary>
        /// <returns>Result of attempting to download the Excel file. </returns>
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
                // unused cells will be added after such as link and image
                table.Rows.Add(item.Id, item.Name, item.Brand, item.Category, item.Price, "", item.Description, "", item.Created);

            }

            // export to excel
            using (XLWorkbook wb = new XLWorkbook())
            {
                var worksheet = wb.Worksheets.Add(table, "Sheet1");
                

                // expand columns to fit text
                worksheet.Columns().AdjustToContents();


                ImageToCell(toList,  worksheet);
                HyperLinkToCell(toList, worksheet);


                worksheet.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                worksheet.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
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




        /// <summary>
        /// Places images into the images cell on the Excel worksheet, resizing to fit the cell.
        /// </summary>
        /// <param name="list"> The list of products (database) </param>
        /// <param name="worksheet"> The Excel worksheet created in ToExcel() </param>
        /// <param name="startRow"> Should be 2 by default, as first is reserved for column names </param>
        /// <param name="startColumn"> The column number of the specified category. </param>
        void ImageToCell(List<Product> list, IXLWorksheet worksheet, int startRow = 2, int startColumn = 6)
        {
            // add images to cell
            foreach (var item in list)
            {
                var currentCell = worksheet.Cell(startRow, startColumn);
                string fullImagePath = environment.WebRootPath + "/pictures/" + item.ImageFileName;
                worksheet.Column(startColumn).Width = Constants.EXCEL_IMAGE_WIDTH;
                worksheet.Row(startRow).Height = Constants.EXCEL_IMAGE_HEIGHT;
                var pic = worksheet.AddPicture(fullImagePath).MoveTo(currentCell, currentCell.CellBelow().CellRight());
                pic.Placement = ClosedXML.Excel.Drawings.XLPicturePlacement.MoveAndSize;


                ++startRow;
            }
        }
        /// <summary>
        /// Places Hyperlinks on the link column in the excel worksheet, if applicable.
        /// </summary>
        /// <param name="list"> The list of products (database) </param>
        /// <param name="worksheet"> The Excel worksheet created in ToExcel() </param>
        /// <param name="startRow"> Should be 2 by default, as first is reserved for column names </param>
        /// <param name="startColumn"> The column number of the specified category. </param>
        void HyperLinkToCell(List<Product> list, IXLWorksheet worksheet, int startRow = 2, int startColumn = 8)
        {
            foreach (var item in list)
            {
                // if no link put No Link in cell instead
                if (String.IsNullOrEmpty(item.Link))
                {
                    worksheet.Cell(startRow, startColumn).Value = "No Link";
                    ++startRow;
                    continue;
                }

                // create hyperlink
                worksheet.Cell(startRow, startColumn).Value = "Link";
                worksheet.Cell(startRow, startColumn).SetHyperlink(new XLHyperlink(@item.Link));

                
                ++startRow;
            }
        }

        /// <summary>
        /// Updates price by rechecking the Amazon link, updates price difference symbol then saves to db.
        /// Arrow up if price increased, down if decreased, else stick with the dot.
        /// </summary>
        /// <param name="id"> Product id, supplied in the Index.cshtml route id </param>
        /// <returns> Refreshes index page. </returns>
        public IActionResult UpdatePrice(int id)
        {

            var product = context.Products.Find(id);

            // if not found in DB, just refresh
            if (product == null)
            {
                return RedirectToAction("Index", "Products");
            }

            string link = product.Link;
            decimal price = product.Price;

            List<string> result = WebScrapeAmazon(product.Link, ACTION_TYPE.PRICE);

            decimal priceConverted = Convert.ToDecimal(result[0]);

            product.PriceDifferenceSymbol = GetPriceDifferenceSymbol(product.Price, priceConverted);

            product.Price = priceConverted;

            context.SaveChanges();


            return RedirectToAction("Index", "Products");


        }
        /// <summary>
        /// Compares two decimal parameters and returns a unicode symbol depending on the result.
        /// </summary>
        /// <param name="oldPrice"> Old product price before updating </param>
        /// <param name="newPrice"> New product price to compare with </param>
        /// <returns> Up arrow if new price is higher, down if lower, dot if same. </returns>
        string GetPriceDifferenceSymbol(decimal oldPrice, decimal newPrice)
        {
            string symbol = "•";
            if (oldPrice > newPrice)
            {
                symbol = "↓";
            }
            else if (oldPrice < newPrice)
            {
                symbol = "↑";
            }

            return symbol;
        }




    }

}
