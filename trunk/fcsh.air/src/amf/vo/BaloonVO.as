package amf.vo {

    public class BaloonVO
    {
        public static const INFO:int = 1;
        public static const WARNING:int = 2;
        public static const ERROR:int = 3;

        [Serializable]
        public var title:String;
        [Serializable]
        public var message:String;
        [Serializable]
        public var type:int;

        public function toString():String
        {
            return "BaloonVO: title = " + title + "; message = " + message + "; type = " + type;
        }
    }
}