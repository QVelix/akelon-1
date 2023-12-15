// See https://aka.ms/new-console-template for more information

using akelon_1.Classes;

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
            excelReader = new ExcelReader(filepath);
            
            break;
        }
        case "ChangeClientName":
        {
            break;
        }
        case "GetRequestyByProduct":
        {
            break;
        }
    }
}