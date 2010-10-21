package log
{
    import mx.controls.Text;

    public class TraceRenderer extends Text
    {
        public function TraceRenderer()
        {
            percentHeight = 100;
            percentWidth = 100;
        }

        override public function set data(value:Object):void
        {
            super.data = value;
            if (value)
            {
                text = value as String;
            }
            else
            {
                text = "";
            }
        }
    }
}