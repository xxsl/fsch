package log
{
    import avmplus.getQualifiedClassName;
    import avmplus.getQualifiedSuperclassName;

    import connection.ITraceMessage;

    import mx.controls.Label;
    import mx.core.IFactory;

    /**
     *@author aturtsevitch
     *@date   Oct 20, 2010
     *@time   2:51:47 PM
     *@langversion ActionScript 3.0
     */
    public class MessageRenderer extends Label implements IFactory
    {
        public function MessageRenderer()
        {
            truncateToFit = true;
            minWidth = 0;
        }

        override public function set data(value:Object):void
        {
            super.data = value;
            if (value)
            {
                var message:ITraceMessage = ITraceMessage(value);
                if(getQualifiedSuperclassName(message.message) == "Object")
                {
                    text = message.message;
                }
                else
                {
                    text = message.messageAsString;
                }
            }
            else
            {
                text = "";
            }
        }

        public function newInstance():*
        {
            return new MessageRenderer();
        }
    }
}