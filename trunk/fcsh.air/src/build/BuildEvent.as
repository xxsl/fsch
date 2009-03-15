package build {
    import flash.events.Event;

    public class BuildEvent extends Event
    {
        public static const BUILD_SUCCESSFULL:String = "BUILD_SUCCESSFULL";
        public static const BUILD_ERROR:String = "BUILD_ERROR";
        public static const BUILD_WARNING:String = "BUILD_WARNING";

        public var info:Array;
        public var errors:Array;
        public var warnings:Array;

        public function BuildEvent(type:String, info:Array = null, warnings:Array = null, errors:Array = null, bubbles:Boolean = false, cancelable:Boolean = false)
        {
            super(type, bubbles, cancelable);
            this.info = info;
            this.warnings = warnings;
            this.errors = errors;
        }
    }
}