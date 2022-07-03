// See https://aka.ms/new-console-template for more information
using PlaygroundApp;

Parallel.For(0, 500, i =>
{
    var handler = ExcelComHandler.GetInstance();
    var api = new ExcelApi(handler);
    api.Run();
});
Console.ReadLine();
