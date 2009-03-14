package amf
{
    import flash.utils.ByteArray;
    import flash.utils.IDataOutput;
    import flash.utils.getQualifiedClassName;

    import mx.logging.ILogger;
    import mx.logging.Log;

    public class VB6Serializer extends FlexTypeDescriber
    {
        private static var log:ILogger = Log.getLogger("amf.VB6Serializer");


        public static function serialize(object:Object, output:IDataOutput):void
        {
            log.debug("Begin serialization");
            var className:String = getQualifiedClassName(object);
            log.debug(" Class name is " + className);
            writeUnicodeString(className, output);
            for each(var property:ProperyObject in getProperties(object))
            {
                log.debug("     Process property: " + property);
                if (isSerializable(object, property.name))
                {
                    log.debug("         Property: " + property.toString() + " is serializable");
                    if (property.type == INT)
                    {
                        var integer:int = object[property.name];
                        log.debug("         Property " + property.name + " has value " + integer);
                        writeInt(integer, output);
                    }
                    else if (property.type == NUMBER)
                    {
                        var number:Number = object[property.name];
                        log.debug("         Property " + property.name + " has value " + number);
                        writeNumber(number, output);
                    }
                    else if (property.type == STRING)
                        {
                            var string:String = object[property.name];
                            log.debug("         Property " + property.name + " has value " + string);
                            writeUnicodeString(string, output);
                        }
                        else if (property.type == BOOLEAN)
                            {
                                var boolean:Boolean = object[property.name];
                                log.debug("         Property " + property.name + " has value " + boolean);
                                writeBoolean(boolean, output);
                            }
                            else
                            {
                                throw new Error("Unsupported type for serialization: " + property.type);
                            }
                }
                else
                {
                    log.debug("         Property: " + property + " is not serializable. skip");
                }
            }
            log.debug("Serialization finished");
        }

        private static function writeBoolean(boolean:Boolean, output:IDataOutput):void
        {
            output.writeByte(boolean ? 1 : 0);
        }

        private static function writeNumber(number:Number, output:IDataOutput):void
        {
            output.writeDouble(number);
        }

        private static function writeInt(integer:int, output:IDataOutput):void
        {
            output.writeInt(integer);
        }

        private static function writeUnicodeString(str:String, output:IDataOutput):void
        {
            var byteArr:ByteArray = new ByteArray();
            byteArr.writeMultiByte(str, "utf-16");
            var len:int = byteArr.length;
            output.writeInt(len);
            output.writeMultiByte(str, "unicode");
        }
    }
}