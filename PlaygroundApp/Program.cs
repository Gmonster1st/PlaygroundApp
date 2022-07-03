// See https://aka.ms/new-console-template for more information
using PlaygroundApp;
using System.Diagnostics;

Stopwatch stopwatch = Stopwatch.StartNew();
stopwatch.Start();
Parallel.For(0, 500, i =>
{
    var handler = ExcelComHandler.GetInstance();
    var api = new ExcelApi(handler);
    api.Run();
});
Console.ReadLine();
stopwatch.Stop();
Console.WriteLine();
Console.WriteLine($"Elapsed time: {stopwatch.Elapsed}");
Console.ReadLine();
