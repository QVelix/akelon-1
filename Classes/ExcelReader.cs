using ClosedXML.Excel;

namespace akelon_1.Classes;

public class ExcelReader
{
    private XLWorkbook excelFile;

    public ExcelReader(string filepath)
    {
        excelFile = new XLWorkbook(filepath);
    }

    
}