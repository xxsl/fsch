package amf {
    import flash.utils.IDataInput;
    import flash.utils.getDefinitionByName;

    import mx.logging.ILogger;
    import mx.logging.Log;

    public class VB6Deserializer  extends FlexTypeDescriber
    {
        private static var log:ILogger = Log.getLogger("amf.VB6Deserializer");

        public static function deserialize(input:IDataInput):*
        {
            log.debug("Begin deserialization");
            var className:String = readUnicodeString(input);
            log.debug(" Class name is " + className);
            try
            {
                var classObject:Class = getDefinitionByName(className) as Class;
            }
            catch(e:Error)
            {
                throw new Error("Object definition was not found: " + className);
            }
            var result:Object = new classObject();

            if (result)
            {
                log.debug("     Class instantiated successfully");
            }
            else
            {
                log.debug("     Class instantiation failed");
            }

            for each(var property:ProperyObject in getProperties(result))
            {
                log.debug("         Process property: " + property);
                if (isSerializable(result, property.name))
                {
                    log.debug("         Property: " + property.toString() + " is serializable");
                    if (property.type == "int")
                    {
                        var integer:int = readInt(input);
                        log.debug("             Property " + property.name + " has value " + integer);
                        result[property.name] = integer;
                    }
                    else if (property.type == "Number")
                    {
                        var number:Number = readNumber(input);
                        log.debug("             Property " + property.name + " has value " + number);
                        result[property.name] = number;
                    }
                    else if (property.type == "String")
                        {
                            var string:String = readUnicodeString(input);
                            log.debug("             Property " + property.name + " has value " + string);
                            result[property.name] = string;
                        }
                        else if (property.type == "Boolean")
                            {
                                var boolean:Boolean = readBoolean(input);
                                log.debug("             Property " + property.name + " has value " + boolean);
                                result[property.name] = boolean;
                            }
                            else
                            {
                                throw new Error("Unsupported type for deserialization: " + property.type);
                            }
                }
                else
                {
                    log.debug("		Property: " + property + " is not serializable. skip");
                }
            }
            log.debug("Deserialization finished");
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