using ClosedXML.Excel;

namespace akelon_1.Classes;

public class ExcelReader
{
    private XLWorkbook excelFile;

    public ExcelReader(string filepath)
    {
        excelFile = new XLWorkbook(filepath);
    }

    public List<Product> GetProducts()
    {
        List<Product> products = new List<Product>();
        IXLWorksheet worksheet = excelFile.Worksheet(1);
        int worksheetRows = worksheet.Rows().Count();
        for (int i = 2; i < worksheetRows; i++)
        {
            if (CheckEndRow(worksheet, i))
            {
                break;
            }
            
            Product product = new Product();
            product.Code = Convert.ToInt32(worksheet.Cell(i, 1).Value.ToString());
            product.Name = worksheet.Cell(i, 2).Value.ToString();
            product.MeasureUnit = worksheet.Cell(i, 3).Value.ToString();
            product.Price = Convert.ToDecimal(worksheet.Cell(i, 4).Value.ToString());
            products.Add(product);
        }
        return products;
    }

    public List<Client> GetCleints()
    {
        List<Client> clients = new List<Client>();
        IXLWorksheet worksheet = excelFile.Worksheet(2);
        int worksheetRows = worksheet.Rows().Count();
        Console.WriteLine(worksheetRows);
        for (int i = 2; i < worksheetRows; i++)
        {
            if (CheckEndRow(worksheet, i))
            {
                break;
            }

            Client client = new Client();
            client.Code = Convert.ToInt32(worksheet.Cell(i, 1).Value.ToString());
            client.CompanyName = worksheet.Cell(i, 2).Value.ToString();
            client.Address = worksheet.Cell(i, 3).Value.ToString();
            client.ContactPerson = worksheet.Cell(i, 4).Value.ToString();
            clients.Add(client);
        }
        return clients;
    }

    public List<Request> GetRequests()
    {
        List<Request> requests = new List<Request>();
        IXLWorksheet worksheet = excelFile.Worksheet(3);
        int worksheetRows = worksheet.Rows().Count();
        for (int i = 2; i < worksheetRows; i++)
        {
            if (CheckEndRow(worksheet, i))
            {
                break;
            }

            Request request = new Request();
            request.Code = Convert.ToInt32(worksheet.Cell(i, 1).Value.ToString());
            request.ProductCode = Convert.ToInt32(worksheet.Cell(i, 2).Value.ToString());
            request.ClientCode = Convert.ToInt32(worksheet.Cell(i, 3).Value.ToString());
            request.Number = Convert.ToInt32(worksheet.Cell(i, 4).Value.ToString());
            request.Count = Convert.ToInt32(worksheet.Cell(i, 5).Value.ToString());
            request.PlacementDate = DateOnly.FromDateTime(DateTime.Parse(worksheet.Cell(i, 6).Value.ToString()));
            requests.Add(request);
        }
        return requests;
    }

    private bool CheckEndRow(IXLWorksheet worksheet, int row)
    {
        return String.IsNullOrEmpty(worksheet.Cell(row, 1).Value.ToString());
    }
}