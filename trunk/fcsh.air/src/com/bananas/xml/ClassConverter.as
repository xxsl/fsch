package com.bananas.xml {
    import mx.utils.DescribeTypeCache;

    public class ClassConverter
    {
        private var config:*;

        public function ClassConverter(config:*)
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
                if (config[prpertyName] != null && (!isNaN(config[prpertyName]) || !(config[prpertyName] is Number)))
                {
                    nodeName = String(element.metadata.arg.(@key == "name").@value);
                    var className:String = element.metadata.arg.(@key == "object").@value;
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
                    //ignore property
                }
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
    }
}