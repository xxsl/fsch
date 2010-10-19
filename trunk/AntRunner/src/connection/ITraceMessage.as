package connection
{
    public interface ITraceMessage extends ISerializable
    {
        function get time():Date;

        function get sender():String;

        function get message():*;
        
        function set time(value:Date):void;

        function set sender(value:String):void;

        function set message(value:*):void;
    }
}