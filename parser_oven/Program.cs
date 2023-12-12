using Aspose.Cells;
using Newtonsoft.Json;
using parser_oven.DB;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Mail;
using static XlsFile;

class Program
{
    static async Task Main()
    {
        //app.config скрыт
        Log log = new Log();
        Context context = new Context();
        await context.GetListProductDB();
        log.Logger("Продукты из БД взяты", DateTime.Now);
        XlsFile xlsFile = new XlsFile();
        await xlsFile.DownloadFileXls();
        log.Logger("Файл с ценами скачан", DateTime.Now);
        xlsFile.ReadFileXls();
        log.Logger("Файл с ценами спаршен", DateTime.Now);
        Sorting sorting = new Sorting(context, xlsFile);
        ProductApi product = sorting.SortingProduct();
        log.Logger("Продукты отсортированы", DateTime.Now);
        Api api = new Api(product);
        await api.SendProductApi();
        log.Logger("Продукты отправлены по API", DateTime.Now);
        SendMessage sendMessage = new SendMessage(product);
        sendMessage.SendEmail();
        log.Logger("Отчет на почту отправлен, конец", DateTime.Now);
    }
}
interface ILog
{
    public void Logger(string message, DateTime date);
}
class Log : ILog
{
    public void Logger(string message, DateTime date) => File.AppendAllText("Info/log.txt", $"{message} {date.ToString("dd.MM.yyyy")}" + Environment.NewLine);
}
interface ISendMessage
{
    public void SendEmail();
}
class SendMessage : ISendMessage
{
    private List<string> _to = new List<string>(ConfigurationManager.AppSettings["SendTo"].Split(new char[] { ';' }));
    private string _from = ConfigurationManager.AppSettings["SendFrom"];
    private string _password = ConfigurationManager.AppSettings["SendPassword"];
    private ProductApi _productApi;
    private int _countProduct;
    private int _countProductNull = 0;
    private List<int> _idproductNull = new List<int>();
    public SendMessage(ProductApi productApi)
    {
        _productApi = productApi;
    }
    public void SendEmail()
    {
        _countProduct = _productApi.products.Count;
        MailMessage msg = new MailMessage();
        msg.IsBodyHtml = true;
        msg.From = new MailAddress(_from);
        msg.Subject = "Отчет Парсера ОВЕН от " + DateTime.Now.ToString("dd.MM.yyyy");
        foreach (var email in _to)
            msg.To.Add(new MailAddress(email));
        if (_countProduct != 0)
        {
            foreach (var product in _productApi.products)
            {
                if (product.price == "0")
                {
                    _countProductNull++;
                    _idproductNull.Add(product.product_id);
                }
            }
            msg.Body += $"Количество обновленных продуктов: {_countProduct}<br>";
            msg.Body += $"Из них цена обнулилась у {_countProductNull} продуктов<br>";
            msg.Body += $"<br>";
            if (_countProductNull > 0)
            {
                msg.Body += $"<b>Продукты которые обнулились</b><br>";
                foreach (var productId in _idproductNull)
                {
                    msg.Body += $"{productId}: http://сайт&product_id={productId}<br>";
                }
            }
        }
        else
        {
            msg.Body += $"Парсер не обнаружил изменений в цене";
        }
        using (SmtpClient client = new SmtpClient())
        {
            client.EnableSsl = true;
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential(_from, _password);
            client.Host = "smtp.yandex.ru";
            client.Port = 587;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.Send(msg);
        }
    }
}
class Api
{
    private string _authorization = ConfigurationManager.AppSettings["Authorization"];
    private ProductApi _product;
    public Api(ProductApi sorting)
    {
        _product = sorting;
    }
    public async Task SendProductApi()
    {
        string jsonProduct = JsonConvert.SerializeObject(_product);
        var client = new HttpClient();
        var request = new HttpRequestMessage()
        {
            Method = HttpMethod.Post,
            RequestUri = new Uri("urlAPI"), // скрыл
            Headers =
            {
                {"Authorization", _authorization},
            },
            Content = new StringContent(jsonProduct)
            {
                Headers =
                {
                   ContentType = new MediaTypeHeaderValue("application/json")
                }
            }
        };
        using (var response = await client.SendAsync(request))
        {
            var body = response.Content.ReadAsStringAsync();
        }
    }

}
class Sorting
{
    private Context _context;
    private XlsFile _xlsFile;
    private ProductApi _productApi = new ProductApi();
    public Sorting(Context context, XlsFile xlsFile)
    {
        _context = context;
        _xlsFile = xlsFile;
    }
    public ProductApi SortingProduct()
    {
        int i = 1;
        foreach (var product in _xlsFile.products)
        {
            var productDB = _context.products.Where(p => p.model == product.Article || p.model == product.sku).FirstOrDefault();
           
            if (productDB != null)
            {
                ProductFind productReady = new ProductFind();
                string name_xls = product.sku.Replace(".", "-").Replace(",", "-").Replace("/", "-").Trim();

                string name_DB = productDB.oc_Product_Description.name.Split("ОВЕН").Last().Replace(".", "-").Replace(",", "-").Replace("/", "-").Trim();

                if (product.Article == productDB.model && name_xls != name_DB)
                {
                    productReady.product_id = productDB.product_id;
                    productReady.price = "0";
                    _productApi.products.Add(productReady);
                    continue;
                }
                string priceDB = productDB.price.ToString("0.0000").Trim();
                string priceXmls = product.Price.ToString("0.0000").Trim();

                if (priceDB != priceXmls)
                {
                    try
                    {
                        productReady.price = priceXmls;
                        Console.WriteLine($" {i} - {productDB.product_id} - цена = {productReady.price}");
                    }
                    catch
                    {
                        productReady.price = priceXmls;
                        Console.WriteLine($" {i} - {productDB.product_id} - цена = {productReady.price}");
                    }
                    productReady.product_id = productDB.product_id;

                    _productApi.products.Add(productReady);
                }
                i++;

            }
        }
        return _productApi;

    }
}
class XlsFile
{
    public class Product
    {
        public string? Article;
        private double _price;
        public string sku;
        public double Price
        {
            get
            {
                return _price;
            }
            set
            {
                _price = value;
            }
        }
    }
    public List<Product> products = new List<Product>();
    private string _url = ConfigurationManager.AppSettings["Url"];
    public async Task DownloadFileXls()
    {
        WebClient web = new WebClient();
        web.DownloadFile(_url, "Info/price_oven.xlsx");
    }
    public List<Product> ReadFileXls()
    {
        Workbook wb = new Workbook($"Info/price_oven.xlsx");
        WorksheetCollection collection = wb.Worksheets;
        for (int index = 0; index < collection.Count; index++)
        {
            Worksheet worksheet = collection[index];

            Console.WriteLine($"Worksgeet:" + worksheet.Name);

            int rows = worksheet.Cells.MaxDataRow;

            int cols = worksheet.Cells.MaxDataColumn;

            for (int i = 4; i <= rows; i++)
            {
                Product product = new Product();
                try
                {
                    product.Article = worksheet.Cells[i, 1].Value.ToString();
                    
                    string price = "";
                    try
                    {
                        price = worksheet.Cells[i, 5].Value.ToString();

                        product.Price = Convert.ToDouble(price.Replace(",", "."));
                    }
                    catch
                    {
                        price = worksheet.Cells[i, 5].Value.ToString();

                        product.Price = Convert.ToDouble(price.Replace(".",","));
                    }
                    product.sku = worksheet.Cells[i, 2].Value.ToString().Trim();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                products.Add(product);
            }
        }
        return products;
    }
}
