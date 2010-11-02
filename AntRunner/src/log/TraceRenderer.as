package log
{
    import mx.controls.Label;
    import mx.core.IFactory;

    public class TraceRenderer extends Label implements IFactory
    {
        public function TraceRenderer()
        {
            percentHeight = 100;
            percentWidth = 100;
            setStyle("paddingTop", 0);
            setStyle("paddingLeft", 0);
            setStyle("paaddingRight", 0);
            setStyle("paddingBottom", 0);
            truncateToFit = true;
        }

        override public function set data(value:Object):void
        {
            super.data = value;
            if (value)
            {
                var line:String = TraceLine(value).text;
                if(line.toLocaleLowerCase().indexOf("warning") >= 0)
                {
                    setStyle("color", 0xEE5400);
                }
                else if(line.toLocaleLowerCase().indexOf("error") >= 0)
                {
                    setStyle("color", 0xFF0000);
                }
                else
                {
                    setStyle("color", 0x000000);
                }
                text = line;
            }
            else
            {
                setStyle("color", 0x000000);
                text = "";
            }
        }

        public function newInstance():*
        {
            return new TraceRenderer();
        }
    }
}