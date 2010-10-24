package connection
{
	[RemoteClass(alias="connection.TraceMessage")]
    public class TraceMessage implements ITraceMessage
    {
        private var _time:Date;
        private var _sender:String;
        private var _messageAsString:String;
        private var _message:*;
        private var _version:int;


        public function TraceMessage( message:* = null, sender:Object = null)
        {
            _time = new Date();
            _sender = Utils.toString(sender);
            _message = message;
            _version = ClassesVersion.VERSION;
            _messageAsString = message ? message.toString(): "null";
        }

        public function get time():Date
        {
            return _time;
        }

        public function get sender():String
        {
            return _sender;
        }

        public function get message():*
        {
            return _message;
        }

        public function get version():int
        {
            return _version;
        }

        public function get messageAsString():String
        {
            return _messageAsString;
        }

        public function set messageAsString(value:String):void
        {
            _messageAsString = value;
        }

        public function set time(value:Date):void
        {
            _time = value;
        }

        public function set sender(value:String):void
        {
            _sender = value;
        }

        public function set message(value:*):void
        {
            _message = value;
        }


        public function set version(value:int):void
        {
            _version = value;
        }

        public function toString():String
        {
            return _messageAsString;
        }
    }
}