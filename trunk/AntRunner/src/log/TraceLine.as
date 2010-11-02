package log
{
    /**
     *@author aturtsevitch
     *@date   Nov 2, 2010
     *@time   5:06:29 PM
     *@langversion ActionScript 3.0
     */
    public class TraceLine
    {
        public var text:String;


        public function TraceLine(text:String)
        {
            this.text = text;
        }

        public function toString():String
        {
            return text;
        }
    }
}