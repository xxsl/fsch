package connection
{
    import flash.events.Event;

    public class ConnectionEvent extends Event
    {
        public static const MESSAGE_INPUT:String = "MESSAGE_IN";

        public var message:ITraceMessage;


        public function ConnectionEvent(type:String)
        {
            super(type, true, true);
        }
    }
}