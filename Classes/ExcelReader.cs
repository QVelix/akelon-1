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
        for (int i = 1; i < worksheetRows; i++)
        {
            if (CheckEndRow(worksheet))
            {
                break;
            }
            
            Product product = new Product();
            product.Code = Convert.ToInt32(worksheet.Cell(i, 1).Value);
            product.Name = Convert.ToString(worksheet.Cell(i, 2).Value);
            product.MeasureUnit = Convert.ToString(worksheet.Cell(i, 3).Value);
            product.Price = Convert.ToDecimal(worksheet.Cell(i, 4).Value);
            products.Add(product);
        }
        return products;
    }

    public List<Client> GetCleints()
    {
        List<Client> clients = new List<Client>();
        IXLWorksheet worksheet = excelFile.Worksheet(2);
        int worksheetRows = worksheet.Rows().Count();
        for (int i = 1; i < worksheetRows; i++)
        {
            if (CheckEndRow(worksheet))
            {
                break;
            }

            Client client = new Client();
            client.Code = Convert.ToInt32(worksheet.Cell(i, 1).Value);;
            client.CompanyName = Convert.ToString(worksheet.Cell(i, 2).Value);;
            client.Address = Convert.ToString(worksheet.Cell(i, 3).Value);;
            client.ContactPerson = Convert.ToString(worksheet.Cell(i, 4).Value);;
            clients.Add(client);
        }
        return clients;
    }

    public List<Request> GetRequests()
    {
        List<Request> requests = new List<Request>();
        IXLWorksheet worksheet = excelFile.Worksheet(3);
        int worksheetRows = worksheet.Rows().Count();
        for (int i = 1; i < worksheetRows; i++)
        {
            if (CheckEndRow(worksheet))
            {
                break;
            }

            Request request = new Request();
            request.Code = Convert.ToInt32(worksheet.Cell(i, 1).Value);
            request.ProductCode = Convert.ToInt32(worksheet.Cell(i, 2).Value);
            request.ClientCode = Convert.ToInt32(worksheet.Cell(i, 3).Value);
            request.Number = Convert.ToInt32(worksheet.Cell(i, 4).Value);
            request.Count = Convert.ToInt32(worksheet.Cell(i, 5).Value);
            request.PlacementDate = DateOnly.Parse(Convert.ToString(worksheet.Cell(i, 6).Value));
            requests.Add(request);
        }
        return requests;
    }

    private bool CheckEndRow(IXLWorksheet worksheet)
    {
        return String.IsNullOrEmpty(Convert.ToString(worksheet.Cell(1, 0)));
    }
}