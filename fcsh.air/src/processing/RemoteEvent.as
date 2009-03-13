package processing {
    import amf.vo.BaloonVO;
    import amf.vo.CommandVO;
    import amf.vo.DataVO;
    import amf.vo.ErrorVO;

    import flash.events.Event;

    public class RemoteEvent extends Event
    {
        public static const DATA_EVENT:String = "DATA_EVENT";
        public static const ERROR_EVENT:String = "ERROR_EVENT";
        public static const COMMAND_EVENT:String = "COMMAND_EVENT";

        public var data:DataVO;
        public var command:CommandVO;
        public var error:ErrorVO;


        public function RemoteEvent(type:String, data:DataVO, command:CommandVO, error:ErrorVO, bubbles:Boolean = false, cancelable:Boolean = false)
        {
            super(type, bubbles, cancelable);
            this.data = data;
            this.command = command;
            this.error = error;
        }
    }
}