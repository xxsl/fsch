package amf {
    import flash.utils.IDataInput;
    import flash.utils.getDefinitionByName;

    public class VB6Deserializer  extends FlexTypeDescriber
    {
        public static function deserialize(input:IDataInput):*
        {
            trace("Begin deserialization");
            var className:String = readUnicodeString(input);
            trace(" Class name is " + className);
            try
            {
            	var classObject:Class = getDefinitionByName(className) as Class;
            }
            catch(e:Error)
            {
            	trace("		[ERROR]Object definition was not found: " + className);
            	return null;
            }
            var result:Object = new classObject();

            if (result)
            {
                trace("     Class instantiated successfully");
            }
            else
            {
                trace("     Class instantiation failed");
            }

            for each(var property:ProperyObject in getProperties(result))
            {
                trace("         Process property: " + property);
                if (isSerializable(result, property.name))
                {
                    trace("         Property: " + property.toString() + " is serializable");
                    if (property.type == "int")
                    {
                        var integer:int = readInt(input);
                        trace("             Property " + property.name + " has value " + integer);
                        result[property.name] = integer;
                    }
                    else if (property.type == "Number")
                    {
                        var number:Number = readNumber(input);
                        trace("             Property " + property.name + " has value " + number);
                        result[property.name] = number;
                    }
                    else if (property.type == "String")
                        {
                            var string:String = readUnicodeString(input);
                            trace("             Property " + property.name + " has value " + string);
                            result[property.name] = string;
                        }
                        else if (property.type == "Boolean")
                            {
                                var boolean:Boolean = readBoolean(input);
                                trace("             Property " + property.name + " has value " + boolean);
                                result[property.name] = boolean;
                            }
                            else
                            {
                                throw new Error("Unsupported type for deserialization: " + property.type);
                            }
                }
                else
                {
                    trace("		Property: " + property + " is not serializable. skip");
                }
            }
            trace("Deserialization finished");
            return result;
        }

        private static function readBoolean(input:IDataInput):Boolean
        {
            var result:int;
            result = input.readByte();
            return result == 1;
        }

        private static function readNumber(input:IDataInput):Number
        {
            var result:Number;
            result = input.readDouble();
            return result;
        }

        private static function readInt(input:IDataInput):int
        {
            var result:int;
            result = input.readInt();
            return result;
        }

        private static function readUnicodeString(input:IDataInput):String
        {
            var result:String;
            var len:int = input.readInt();
            result = input.readMultiByte(len, "utf-16");
            return result;
        }
    }
}