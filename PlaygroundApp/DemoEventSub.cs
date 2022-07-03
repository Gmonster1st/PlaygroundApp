namespace PlaygroundApp
{
    public class DemoEventSub
    {
        DemoWithEvent test;
        public DemoEventSub(DemoWithEvent handler)
        {
            test = handler;
            test.MyEvent += OnEvent;
        }
        
        public void OnEvent(object? sender, EventArgs e)
        {
            Console.WriteLine("Event Fired!");
        }
    }
}
