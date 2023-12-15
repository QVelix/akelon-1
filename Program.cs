// See https://aka.ms/new-console-template for more information
if ((args.Contains("help") || args.Contains("h"))&&(!args.Contains("command")||!args.Contains("c")))
{
    Console.WriteLine("Список команд приложения:\n" +
                      "* help/h - Вывести список команд\n" +
                      "* filepath/f - Ввести путь до файла. Пример: filepath=/home/user/download/myExcel.xlsx\n" +
                      "* command/c - Выполнить команду. Пример: command=GetGoldenClient\n" +
                      "* command-help/ch - Получить список доступных команд");
    return;
}

if (args.Contains("filepath") || args.Contains("f"))
{
    
}

if ((args.Contains("command") || args.Contains("c")) && (!args.Contains("help") || !args.Contains("h")))
{
    
}
else if (args.Contains("command") || args.Contains("ch"))
{
    Console.WriteLine("Список доступных команд для выполнения:\n" +
                      "* GetGoldenClinet - получить золотого клиента\n" +
                      "* ChangeClientName - сменить имя у клиента\n" +
                      "* GetRequestByProduct - получить все заказы по товару");
    return;
}