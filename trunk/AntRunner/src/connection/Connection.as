package connection
{
    import flash.events.AsyncErrorEvent;
    import flash.events.SecurityErrorEvent;
    import flash.events.StatusEvent;
    import flash.net.LocalConnection;

    /**
     *@author aturtsevitch
     *@date   Oct 20, 2010
     *@time   3:45:28 PM
     *@langversion ActionScript 3.0
     */
    public class Connection
    {
        // Connections
        private var lineOut:LocalConnection;
        private var lineIn:LocalConnection;


        // Connection names
        private const LINE_OUT:String = "_debugger_out";
        private const LINE_IN:String = "_debugger_in";

        // The allow domain for the local connection
        // * = Allow communication with all domains
        private const ALLOWED_DOMAIN:String = "*";

        private var isConnected:Boolean = false;


        public function Connection()
        {
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
                } catch(error:ArgumentError)
                {
                    // Do nothing
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