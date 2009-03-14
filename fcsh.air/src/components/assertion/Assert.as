package components.assertion {
    import mx.utils.ObjectUtil;

    public class Assert
    {
        public static function equals(obj1:Object, obj2:Object):void
        {
            if (ObjectUtil.compare(obj1, obj2) != 0)
            {
                throwError();
            }
        }

        public static function assertTrue(result:Boolean):void
        {
            if (!result)
            {
                throwError();
            }
        }

        private static function throwError():void
        {
            throw new Error("Assert failed!");
        }
    }
}