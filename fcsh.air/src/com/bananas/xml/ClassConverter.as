package com.bananas.xml {
    import com.bananas.generated.FlexConfig;

    import flash.utils.getDefinitionByName;

    import mx.logging.ILogger;
    import mx.logging.Log;
    import mx.utils.DescribeTypeCache;

    public class ClassConverter
    {
        private static var log:ILogger = Log.getLogger("com.bananas.xml.ClassConverter");
        private var config:FlexConfig;

        public function ClassConverter(config:FlexConfig)
        {
            this.config = config;
        }

        public function convert():Object
        {
            return convert1(config);
        }

        private function convert1(config:Object):Object
        {
            var result:Object = new Object;
            var classInfo:XML = DescribeTypeCache.describeType(config).typeDescription;
            var elements:XMLList = classInfo.variable;
            var prpertyName:String;
            var nodeName:String;
            for each(var element:XML in elements)
            {
                prpertyName = String(element.@name);
                if (config[prpertyName] != null &&
                    (!isNaN(config[prpertyName]) || !(config[prpertyName] is Number)))
                {
                    nodeName = String(element.metadata.arg.(@key == "name").@value);
                    var className:String = element.metadata.arg.(@key == "object").@value;
                    var cls:Class = getDefinition(className);
                    if (isSimple(className))
                    {
                        result[nodeName] = (config[prpertyName]);
                        result.setPropertyIsEnumerable(nodeName, true);
                    }
                    else
                    {
                        if (String(element.metadata.arg.(@key == "array").@value) == "true")
                        {
                            for each(var i:* in config[prpertyName])
                            {
                                if (!result[nodeName]) {
                                    result[nodeName] = [];
                                    result.setPropertyIsEnumerable(nodeName, true);
                                }
                                result[nodeName].push(convert1(i));
                            }
                        }
                        else
                        {
                            result[nodeName] = convert1(config[prpertyName]);
                            result.setPropertyIsEnumerable(nodeName, true);
                        }
                    }

                }
                else
                {
                    log.debug("ignore property: " + prpertyName + " value: " + config[prpertyName]);
                }
            }
            for (var prop:String in result)
            {
                log.debug("prop: " + result[prop]);
            }
            return result;
        }

        private function isSimple(type:String):Boolean
        {
            var result:Boolean = false;
            switch (type)
                    {
                case "Boolean":
                    result = true;
                    break;
                case "String":
                    result = true;
                    break;
                case "Number":
                    result = true;
                    break;
                case "int":
                    result = true;
                    break;
                case "uint":
                    result = true;
                    break;
                default :
                    result = false;
            }
            return result;
        }

        private function getDefinition(type:String):Class
        {
            var result:Class;

            try
            {
                result = getDefinitionByName(type) as Class;
            }
            catch(e:Error)
            {
                switch (type)
                        {
                    case "Boolean":
                        result = Boolean;
                        break;
                    case "String":
                        result = String;
                        break;
                    case "Number":
                        result = Number;
                        break;
                    case "int":
                        result = Number;
                        break;
                    case "uint":
                        result = Number;
                        break;
                }
            }
            return result;
        }
    }
}