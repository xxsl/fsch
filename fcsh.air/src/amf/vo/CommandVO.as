package amf.vo {

    public class CommandVO
    {
        [Serializable]
        public var command:String;
        [Serializable]
        public var target:String;

        public function toString():String
        {
            return "CommandVO: target = " + target + "; command = " + command + ";";
        }
    }
}