package ui
{
    import mx.collections.ArrayCollection;

    /**
     *@author aturtsevitch
     *@date   Nov 2, 2010
     *@time   5:55:54 PM
     *@langversion ActionScript 3.0
     */
    public interface ILogConsole
    {
        function set dataProvider(value:ArrayCollection):void;

        function set autoScroll(value:Boolean):void;

        function get autoScroll():Boolean;
    }
}