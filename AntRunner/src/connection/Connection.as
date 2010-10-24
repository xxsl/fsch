package connection
{
    import flash.events.AsyncErrorEvent;
    import flash.events.EventDispatcher;
    import flash.events.SecurityErrorEvent;
    import flash.events.StatusEvent;
    import flash.net.LocalConnection;
    import flash.utils.ByteArray;

    /**
     *@author aturtsevitch
     *@date   Oct 20, 2010
     *@time   3:45:28 PM
     *@langversion ActionScript 3.0
     */
    public class Connection extends EventDispatcher
    {
        private static var _instance:Connection;

        private var counter:Number = 0;

        // Connections
        private var lineOut:LocalConnection;
        private var lineIn:LocalConnection;


        // Connection names
        private const LINE_OUT:String = "_debugger_in";
        private const LINE_IN:String = "_debugger_out";

        // The allow domain for the local connection
        // * = Allow communication with all domains
        private const ALLOWED_DOMAIN:String = "*";

        private var isConnected:Boolean = false;


        public function Connection(seal:Seal)
        {
            if (!seal)
            {
                throw new Error("private");
            }

            // Setup line out
            lineOut = new LocalConnection();
            lineOut.addEventListener(AsyncErrorEvent.ASYNC_ERROR, asyncErrorHandler, false, 0, true);
            lineOut.addEventListener(SecurityErrorEvent.SECURITY_ERROR, securityErrorHandler, false, 0, true);
            lineOut.addEventListener(StatusEvent.STATUS, statusHandler, false, 0, true);

            // Setup line in
            lineIn = new LocalConnection();
            lineIn.addEventListener(AsyncErrorEvent.ASYNC_ERROR, asyncErrorHandler, false, 0, true);
            lineIn.addEventListener(SecurityErrorEvent.SECURITY_ERROR, securityErrorHandler, false, 0, true);
            lineIn.addEventListener(StatusEvent.STATUS, statusHandler, false, 0, true);
            lineIn.allowDomain(ALLOWED_DOMAIN);
            lineIn.client = this;

            try
            {
                lineIn.connect(LINE_IN);
            }
            catch(error:ArgumentError)
            {
                trace(error.getStackTrace());
            }
        }


        public static function get instance():Connection
        {
            if (!_instance)
            {
                _instance = new Connection(new Seal());
            }
            return _instance;
        }

        /**
         * External data input
         * @param data compressed ByteArray
         */
        public function input(data:ByteArray):void
        {
            counter++;
            trace("message length", data.length);
            data.uncompress();
            var readObject:ITraceMessage = ITraceMessage(data.readObject());
            trace(counter, readObject.time,readObject.sender, readObject.message);
            var connectionEvent:ConnectionEvent = new ConnectionEvent(ConnectionEvent.MESSAGE_INPUT);
            connectionEvent.message = readObject;
            dispatchEvent(connectionEvent);
        }

        /**
         * Sends data
         * @param message ITraceMessage impl
         */
        public function output(message:ITraceMessage):void
        {
            var bytes:ByteArray = new ByteArray();
            bytes.writeObject(message);
            bytes.compress();
            try
            {
                lineOut.send(LINE_OUT, "input", bytes);
            }
            catch (error:Error)
            {
                trace(error.getStackTrace());
            }
        }

        /**
         * Event handlers for localconnection
         * Disconnect on error
         */
        private function asyncErrorHandler(event:AsyncErrorEvent):void
        {
            isConnected = false;
        }

        private function securityErrorHandler(event:SecurityErrorEvent):void
        {
            isConnected = false;
        }

        private function statusHandler(event:StatusEvent):void
        {
            if (event.level == "error")
            {
                isConnected = false;
            }
        }
    }
}

class Seal
{

}