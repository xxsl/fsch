package amf.vo {
    public class ErrorVO
    {
        [Serializable]
        public var description:String;
        [Serializable]
        public var id:int;

        public function toString():String
        {
            return "ErrorVO: id = " + id + "; description = " + description + ";";
        }
    }
}