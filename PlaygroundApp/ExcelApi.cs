namespace PlaygroundApp
{
    public class ExcelApi
    {
        public event EventHandler<EventArgs>? OnExcelApiEvent;

        private ExcelComHandler _comHandler;
        public ExcelApi(ExcelComHandler comHandler)
        {
            _comHandler = comHandler;
            OnExcelApiEvent += _comHandler.OnProcessComplete;
        }

        public void Run()
        {
            Console.ForegroundColor = (ConsoleColor)(_comHandler.GetHashCode() % 10);
            Console.WriteLine($"Runing on thread {Thread.CurrentThread.ManagedThreadId}");
            Console.WriteLine($"Using App {_comHandler.GetHashCode()} that has {_comHandler.GetUserCount()} users");
            for (int i = 0; i < 1000000000; i++)
            {
                var x = ((i * 10) / 5) + 5;
            }
            processCompleted();
        }

        public void processCompleted()
        {
            if (OnExcelApiEvent != null)
            {
                OnExcelApiEvent(this, EventArgs.Empty);
                OnExcelApiEvent -= _comHandler.OnProcessComplete;
            }
        }
    }
}
