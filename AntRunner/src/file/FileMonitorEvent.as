package file
{
    import flash.events.Event;
    import flash.filesystem.File;

    public class FileMonitorEvent extends Event
    {
        public static const CREATE:String = "CREATE";
        public static const MOVE:String = "MOVE";
        public static const CHANGE:String = "CHANGE";

        public var fileProperty:File = null;

        public function FileMonitorEvent(type:String, bubbles:Boolean = false, cancelable:Boolean = true)
        {
            super(type, bubbles, cancelable);
        }
    }
}