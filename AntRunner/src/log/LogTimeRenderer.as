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
    public class LogTimeRenderer extends Label implements IFactory
    {
        private static var _formatter:DateFormatter;

        public function LogTimeRenderer()
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
                text = formatter.format(message.time) + '.' + message.time.getMilliseconds();
            }
            else
            {
                text = "";
            }
        }

        public function get formatter():DateFormatter
        {
            if (!_formatter)
            {
                _formatter = new DateFormatter();
                _formatter.formatString = "HH:JJ:SS";
            }
            return _formatter;
        }


        public function newInstance():*
        {
            return new LogTimeRenderer();
        }
    }
}