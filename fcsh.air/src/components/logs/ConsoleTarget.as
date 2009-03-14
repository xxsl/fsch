package components.logs {
    import components.assertion.Assert;

    import mx.controls.TextArea;
    import mx.formatters.DateFormatter;
    import mx.logging.AbstractTarget;
    import mx.logging.ILogger;
    import mx.logging.ILoggingTarget;
    import mx.logging.LogEvent;

    public class ConsoleTarget extends AbstractTarget implements ILoggingTarget
    {
        private var formatter:DateFormatter = new DateFormatter();
        private var window:TextArea;

        public function ConsoleTarget(window:TextArea)
        {
            super();
            Assert.assertTrue(window != null);
            this.window = window;
            formatter.formatString = "JJ:NN:SS";
        }

        override public function logEvent(event:LogEvent):void
        {
            super.logEvent(event);

            output(LogEvent.getLevelString(event.level), ILogger(event.target).category, event.message, getColor(event.level));
        }

        private function getColor(level:int):String
        {
            var result:String = "#";
            switch (level)
                    {
                case 0:
                    result += "000000";
                    break;
                case 2:
                    result += "808080";
                    break;
                case 8:
                    result += "C9433C";
                    break;
                case 4:
                    result += "3333CC";
                    break;
                case 1000:
                    result += "FF0000";
                    break;
                case 6:
                    result += "FF9E49";
                    break;
                default:
                    throw new Error("Unknown log level:" + level);
            }
            return result;
        }

        private function output(prefix:String, category:String, msg:String, color:String):void
        {
            var result:String = "[" + prefix + "] " + formatter.format(new Date()) + " [" + category + "] " + msg;
            trace(result);
            var html:String = "<FONT COLOR='" + color + "'>" + result + "</FONT>" + "<br>";
            window.htmlText += html;
            window.validateNow();
        }
    }
}