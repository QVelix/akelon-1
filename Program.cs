// See https://aka.ms/new-console-template for more information

using akelon_1.Classes;
using System.Linq;

string filepath;
ExcelReader excelReader;

if (args.Contains("help"))
{
    Console.WriteLine("Список команд приложения:\n" +
                      "* help- Вывести список команд\n" +
                      "* filepath - Ввести путь до файла. Пример: filepath=/home/user/download/myExcel.xlsx\n" +
                      "* command - Выполнить команду. Пример: command=GetGoldenClient\n" +
                      "* command-help - Получить список доступных команд");
    return;
}
if (args.Contains("command-help"))
{
    Console.WriteLine("Список доступных команд для выполнения:\n" +
                      "* GetGoldenClinet - получить золотого клиента\n" +
                      "* ChangeClientName - сменить имя у клиента\n" +
                      "* GetRequestByProduct - получить все заказы по товару");
    return;
}
// ExcelReader er = new ExcelReader("C:\\Users\\Austrolia\\Downloads\\Практическое задание для кандидата.xlsx");
// er.GetCleints();

if (args.Where(s=>s.IndexOf("filepath") != -1).Count()==0)
{
    Console.WriteLine("Введите путь к файлу: ");
    filepath = Console.ReadLine();
}
else
{
    filepath = args.Where(s => s.IndexOf("filepath") != -1).First();
    filepath = filepath.Replace("filepath=", "");
}

if (args.Where(s=> (s.IndexOf("command") != -1 && s.IndexOf("help") == -1)).Count()>0)
{
    string command = args.Where(s => s.IndexOf("command")!=-1).First();
    command = command.Replace("command=", "");
    switch (command)
    {
        case "GetGoldenClient":
        {
            Console.WriteLine("Введите год: ");
            int year = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Введите месяц: ");
            int month = Convert.ToInt32(Console.ReadLine());
            excelReader = new ExcelReader(filepath);
            List<Client> clients = excelReader.GetCleints();
            List<Request> requests = excelReader.GetRequests().Where(r=>r.PlacementDate.Year==year&&r.PlacementDate.Month==month).ToList();
            string GoldenClient = "";
            int MaxRequestCount = 0;
            foreach (var client in clients)
            {
                List<Request> clientRequest = requests.FindAll(r => r.ClientCode == client.Code);
                int clientRequestsSum = 0;
                foreach (var cr in clientRequest)
                {
                    clientRequestsSum += cr.Count;
                }

                if (MaxRequestCount < clientRequestsSum)
                {
                    MaxRequestCount = clientRequestsSum;
                    Console.WriteLine(client.CompanyName);
                    GoldenClient = client.CompanyName;
                }
            }
            
            Console.WriteLine("Золотой клиент - {0}", GoldenClient);
            break;
        }
        case "ChangeClientName":
        {
            break;
        }
        case "GetRequestByProduct":
        {
            Console.WriteLine("Введите название товара: ");
            string productName = Console.ReadLine().Trim();
            excelReader = new ExcelReader(filepath);
            List<Product> products = excelReader.GetProducts();
            var product = new Product();
            foreach (var p in products)
            {
                if (p.Name.IndexOf(productName, 0)!=-1)
                {
                    product = p;
                    Console.WriteLine(p.Name);
                }
            }
            // Console.WriteLine(products.Count);
            // var product = products.Find(p => p.Name == productName);
            Console.WriteLine(product.Name);
            List<Request> requests = excelReader.GetRequests().Where(r => r.ProductCode == product.Code).ToList();
            List<Client> clients = excelReader.GetCleints();
            Console.WriteLine("Заказали: ");
            foreach (var request in requests)
            {
                Client clientWhoOrdered = clients.Find(c => c.Code == request.ClientCode);
                Console.WriteLine($@"{clientWhoOrdered.CompanyName} заказал {request.PlacementDate} {request.Count} {product.MeasureUnit} на сумму {request.Count*product.Price} по цене {product.Price}");
            }
            break;
        }
    }
}