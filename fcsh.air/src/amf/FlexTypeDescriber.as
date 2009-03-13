package amf
{
    import flash.utils.describeType;

    import mx.collections.ArrayCollection;
    import mx.collections.Sort;
    import mx.collections.SortField;

    public class FlexTypeDescriber
    {
        protected static const INT:String = "int";
        protected static const NUMBER:String = "Number";
        protected static const STRING:String = "String";
        protected static const BOOLEAN:String = "Boolean";

        public function FlexTypeDescriber()
        {
        }

        protected static function getProperties(object:Object):ArrayCollection
        {
            var classInfo:XML = describeType(object);
            var serializable:XMLList = classInfo.variable.@name;
            var result:ArrayCollection = new ArrayCollection();
            for each (var item:* in serializable)
            {
                var name:String = item.toString();
                var property:ProperyObject = new ProperyObject(name, classInfo.variable.(@name == name).@type.toString());
                result.addItem(property);
            }
            sortProperties(result);
            return result;
        }

        private static function sortProperties(col:ArrayCollection):void
        {
            var sort:Sort = new Sort();
            sort.fields = [new SortField("name", false, false)];
            col.sort = sort;
            col.refresh();
        }

        protected static function isSerializable(object:Object, property:String):Boolean
        {
            var classInfo:XML = describeType(object);
            var serializable:String = classInfo.variable.(@name == property).metadata.(@name == "Serializable").@name;
            return serializable != "";
        }
    }
}