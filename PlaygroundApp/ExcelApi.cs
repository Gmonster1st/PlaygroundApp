namespace PlaygroundApp
{
    public class ExcelApi
    {
        /// <summary>
        /// Event handler that is called when a process is finished.
        /// </summary>
        public event EventHandler<ECHEventArgs>? OnExcelApiEvent;

        /// <summary>
        /// Reference of the Excel application
        /// </summary>
        private ExcelComHandler _comHandler;
        private int procColor;

        /// <summary>
        /// Default Constructor that initializes and configures the class object.
        /// </summary>
        /// <param name="comHandler">The Excel application instance.</param>
        public ExcelApi(ExcelComHandler comHandler)
        {
            _comHandler = comHandler;
            // Subscribe to the event using the Application object.
            OnExcelApiEvent += _comHandler.OnProcessComplete;
        }

        /// <summary>
        /// Run the process.
        /// </summary>
        public void Run()
        {
            procColor = _comHandler.GetHashCode() % 10;
            Console.ForegroundColor = (ConsoleColor)(procColor);
            Console.WriteLine($"Runing on thread {Thread.CurrentThread.ManagedThreadId} => Using App {_comHandler.GetHashCode()} that has {_comHandler.GetUserCount()} users");
            for (int i = 0; i < 1000000000; i++)
            {
                var x = ((i * 10) / 5) + 5;
            }
            // Fire the event to release the Application
            processCompleted();
        }

        /// <summary>
        /// Fire the event that releases the Application object
        /// </summary>
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
