namespace PlaygroundApp
{
    public class DemoWithEvent
    {
        public event EventHandler<EventArgs>? MyEvent;
        public void DoSomething()
        {
            if (MyEvent != null)
            {
                MyEvent(this, EventArgs.Empty);
            }
        }
    }
}
