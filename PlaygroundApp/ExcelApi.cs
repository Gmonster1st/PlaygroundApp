namespace PlaygroundApp
{
    public class ExcelApi
    {
        public event EventHandler<ECHEventArgs>? OnExcelApiEvent;

        private ExcelComHandler _comHandler;
        private int procColor;

        public ExcelApi(ExcelComHandler comHandler)
        {
            _comHandler = comHandler;
            OnExcelApiEvent += _comHandler.OnProcessComplete;
        }

        public void Run()
        {
            procColor = _comHandler.GetHashCode() % 10;
            Console.ForegroundColor = (ConsoleColor)(procColor);
            Console.WriteLine($"Runing on thread {Thread.CurrentThread.ManagedThreadId} => Using App {_comHandler.GetHashCode()} that has {_comHandler.GetUserCount()} users");
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
                OnExcelApiEvent(this, new ECHEventArgs() { ConsoleColor = procColor});
                OnExcelApiEvent -= _comHandler.OnProcessComplete;
            }
        }
    }
}
