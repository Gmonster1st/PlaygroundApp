using System.Runtime.InteropServices;
using Timer = System.Timers.Timer;

namespace PlaygroundApp
{
    public class ExcelComHandler : IDisposable
    {
        private static List<ExcelComHandler> _pool = new List<ExcelComHandler>();
        private static readonly Object _lock = new object();

        private int _userCount;
        private Timer? _timer;
        private int _procColor;

        private ExcelComHandler() { }

        public int GetUserCount()
        {
            return _userCount;
        }

        public static ExcelComHandler GetInstance()
        {
            lock (_lock)
            {
                var instance = _pool.FirstOrDefault(x => x._userCount >= 0 && x._userCount < 5);
                if (instance == null)
                {

                    instance = new ExcelComHandler();
                    _pool.Add(instance);
                    instance._userCount++;
                }
                else
                {
                    instance._userCount++;
                    instance.Timeout();
                }
                return instance;
            }
        }

        public void OnProcessComplete(object? sender, ECHEventArgs e)
        {
            this._userCount--;
            if (this._userCount == 0)
            {
                this.Timeout();
            }
            _procColor = e.ConsoleColor;
        }

        private void Timeout()
        {
            if (_timer == null)
            {
                _timer = new Timer(2000);
                _timer.Elapsed += OnTimerElapsed;
                _timer.AutoReset = false;
                _timer.Start();
            }
            else
            {
                if (_timer.Enabled)
                {
                    _timer.Stop();
                    if (this._userCount == 0)
                    {
                        _timer.Start();
                    }
                    else
                    {
                        _timer.Close();
                        _timer = null;
                    }
                }
            }
        }

        private void OnTimerElapsed(object? sender, System.Timers.ElapsedEventArgs e)
        {
            if (_timer != null)
            {
                _timer.Stop();
                _timer.Close();
            }
            Console.ForegroundColor = (ConsoleColor)_procColor;
            Console.WriteLine($"\n<<kill App {this.GetHashCode()}>> => {_pool.Count}\n");
            _pool.Remove(this);
            _ = Marshal.FinalReleaseComObject(this);
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}
