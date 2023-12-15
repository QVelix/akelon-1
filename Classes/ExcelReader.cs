﻿using ClosedXML.Excel;

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
            if (String.IsNullOrEmpty(Convert.ToString(worksheet.Cell(i, 0).Value)))
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
}