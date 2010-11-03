package ui
{
    import mx.collections.ArrayCollection;

    /**
     *@author aturtsevitch
     *@date   Nov 3, 2010
     *@time   10:55:41 AM
     *@langversion ActionScript 3.0
     */
    public class ViewComposition implements ILogConsole
    {
        private var views:Array = [];

        private var _dataPrivider:ArrayCollection;
        private var _autoScroll:Boolean;

        public function ViewComposition(...rest)
        {
            if(rest)
                this.views = rest;
        }

        public function set dataProvider(value:ArrayCollection):void
        {
            //ignored
            /*_dataPrivider = value;
            for each (var logConsole:ILogConsole in views)
            {
                logConsole.dataProvider = _dataPrivider;
            }*/
        }

        public function set autoScroll(value:Boolean):void
        {
            _autoScroll = value;
            for each (var logConsole:ILogConsole in views)
            {
                logConsole.autoScroll = _autoScroll;
            }
        }

        public function get autoScroll():Boolean
        {
            return _autoScroll;
        }
    }
}