package amf.vo {
    public class DataVO
    {
        [Serializable]
        public var data:String;
        [Serializable]
        public var target:String;

        public function toString():String
        {
            return "DataVO: target = " + target + "; data = " + data + ";";
        }
    }
}