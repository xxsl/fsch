package amf
{
	public class ProperyObject
	{
		public var name:String;
		public var type:String;


        public function ProperyObject(name:String, type:String)
        {
            this.name = name;
            this.type = type;
        }

        public function toString():String
        {
            return "Property " + name + " of type " + type;
        }
    }
}