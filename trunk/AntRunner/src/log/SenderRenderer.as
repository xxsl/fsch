package log
{
    import connection.ITraceMessage;

    import mx.controls.Label;
    import mx.core.IFactory;
    import mx.formatters.DateFormatter;

    /**
     *@author aturtsevitch
     *@date   Oct 20, 2010
     *@time   2:51:47 PM
     *@langversion ActionScript 3.0
     */
    public class SenderRenderer extends Label implements IFactory
    {
        public function SenderRenderer()
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
                text = message.sender;
            }
            else
            {
                text = "";
            }
        }

        public function newInstance():*
        {
            return new SenderRenderer();
        }
    }
}