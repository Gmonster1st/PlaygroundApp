using System.Runtime.InteropServices;
using Timer = System.Timers.Timer;

namespace PlaygroundApp
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelComHandler : IDisposable
    {
        /// <summary>
        /// This is the application pool that holds the references of the Excel applications in memory.
        /// </summary>
        /// <typeparam name="ExcelComHandler"></typeparam>
        private static List<ExcelComHandler> _pool = new List<ExcelComHandler>();

        /// <summary>
        /// Object used as lock for thread safe calls to the <see cref="ExcelComHandler"/>
        /// </summary>
        private static readonly Object _lock = new object();

        /// <summary>
        /// Private field that holds the number of processes currently using the instance of the Application.
        /// </summary>
        private int _userCount;

        /// <summary>
        /// <see cref="System.Timers.Timer"/> Object used for the timeout sequence of the application
        /// </summary>
        private Timer? _timer;
        private int _procColor;// TODO: remove this

        /// <summary>
        /// Default Constructor.
        /// </summary>
        private ExcelComHandler() { }

        /// <summary>
        /// Gets the number of users using this instance of the application.
        /// </summary>
        /// <returns>The number of user as <see cref="int"/>.</returns>
        public int GetUserCount()
        {
            return _userCount;
        }

        /// <summary>
        /// This method returns an instance of the <see cref="ExcelComHandler"/>.<br><br/> 
        /// It will spin up a new instance if none exist, <br><br/>
        /// or if the ones that exist are occupied with the maximum number of users.
        /// </summary>
        /// <returns>An Instance of the <see cref="ExcelComHandler"/>.</returns>
        public static ExcelComHandler GetInstance()
        {
            // Lock the instance before trying to instanciate the class.
            lock (_lock)
            {
                // Find an instance that doesn't have a full user count and return it else return null
                var instance = _pool.FirstOrDefault(x => x._userCount >= 0 && x._userCount < 5);
                if (instance == null)
                {
                    // Create a new Object of the class, add it to the pool and adjust the user count.
                    instance = new ExcelComHandler();
                    _pool.Add(instance);
                    instance._userCount++;
                }
                else
                {
                    // In an already existing instance adjust the user count an call for timeout toadjust as well
                    instance._userCount++;
                    instance.Timeout();
                }
                return instance;
            }
        }

        /// <summary>
        /// This method is used to subscribe to the events of the classes using <see cref="ExcelComHandler"/>.<br><br/>
        /// This method will call the timeout sequence for releasing resources.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void OnProcessComplete(object? sender, ECHEventArgs e)
        {
            // Decrease the user count and if the user count 0 start the timeout sequence
            this._userCount--;
            if (this._userCount == 0)
            {
                this.Timeout();
            }
            _procColor = e.ConsoleColor; // TODO: remove this
        }

        /// <summary>
        /// Checks for the timeout state of the object and adjusts it accordingly.
        /// </summary>
        private void Timeout()
        {
            if (_timer == null)
            {
                // If no timer is available create a new timer and set the interval to the value set.
                // Then start the timer.
                _timer = new Timer(2000);
                _timer.Elapsed += OnTimerElapsed;
                _timer.AutoReset = false;
                _timer.Start();
            }
            else
            {
                if (_timer.Enabled)
                {
                    // If timer exist check if it is running and then adjust it depending on the user Count
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

        /// <summary>
        /// EventHandler method used when the timer expires.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnTimerElapsed(object? sender, System.Timers.ElapsedEventArgs e)
        {
            // Close the timer and release the resources of this instance
            if (_timer != null)
            {
                _timer.Stop();
                _timer.Close();
            }
            // TODO: Remove this region of Writelines
            Console.ForegroundColor = (ConsoleColor)_procColor;
            Console.WriteLine($"\n<<kill App {this.GetHashCode()}>> => {_pool.Count}\n");
            _pool.Remove(this);
            _ = Marshal.FinalReleaseComObject(this);
        }

        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}
