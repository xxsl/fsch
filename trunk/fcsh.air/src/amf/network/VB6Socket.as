package amf.network {
    import amf.VB6Deserializer;
    import amf.VB6Serializer;

    import flash.events.Event;
    import flash.events.IOErrorEvent;
    import flash.events.ProgressEvent;
    import flash.events.SecurityErrorEvent;
    import flash.net.Socket;
    import flash.utils.ByteArray;
    import flash.utils.getQualifiedClassName;

    import mx.logging.ILogger;
    import mx.logging.Log;

    import processing.RemoteController;

    public class VB6Socket extends Socket
    {
        private static var _instance:VB6Socket;

        private var size:int = -1;
        private var buffer:SocketBuffer = new SocketBuffer();
        private var _isConnected:Boolean;

        private static var log:ILogger = Log.getLogger("amf.network.VB6Socket");


        public static function get instance():VB6Socket
        {
            if (!_instance)
            {
                _instance = new VB6Socket();
            }
            return _instance;
        }

        public function VB6Socket(host:String = null, port:uint = 0)
        {
            super(host, port);
            configureListeners();
        }


        private function configureListeners():void
        {
            addEventListener(Event.CLOSE, closeHandler);
            addEventListener(Event.CONNECT, connectHandler);
            addEventListener(IOErrorEvent.IO_ERROR, ioErrorHandler);
            addEventListener(SecurityErrorEvent.SECURITY_ERROR, securityErrorHandler);
            addEventListener(ProgressEvent.SOCKET_DATA, socketDataHandler);
        }

        private function closeHandler(event:Event):void
        {
            log.info("Socket closed");
        }

        private function connectHandler(event:Event):void
        {
            log.info("Socket connected");
        }

        private function ioErrorHandler(event:IOErrorEvent):void
        {
            log.error("IOError: " + event.text);
            dispatchEvent(new Event(Event.CLOSE));
        }

        private function securityErrorHandler(event:SecurityErrorEvent):void
        {
            log.error("Security Error: " + event.text);
            dispatchEvent(new Event(Event.CLOSE));
        }

        private function socketDataHandler(event:ProgressEvent):void
        {
            log.debug("Socket data arrival, bytes loaded: " + event.bytesLoaded);
            var _buffer:ByteArray = new ByteArray();
            readBytes(_buffer, 0, event.bytesLoaded);
            buffer.writeBytes(_buffer, 0, _buffer.length);

            if (size == -1 && buffer.length >= 4)
            {
                size = buffer.readInt();
                log.debug("Object size: " + size);
            }

            while (buffer.length >= size && (size != -1))
            {
                log.debug("Read object: " + size + " bytes");
                var object:Object = VB6Deserializer.deserialize(buffer);
                log.info("Object is: " + getQualifiedClassName(object));
                RemoteController.instance.process(object);
                log.debug("Read complete");
                size = -1;
                if (buffer.length >= 4)
                {
                    size = buffer.readInt();
                    log.debug("Object size: " + size);
                }
            }
            log.debug("Socket data processed. size " + size + " bytes , buffer.length " + buffer.length);
        }

        public function sendObject(object:Object):void
        {
            var byteArray:ByteArray = new ByteArray();
            VB6Serializer.serialize(object, byteArray);
            writeInt(byteArray.length);
            writeBytes(byteArray);
            flush();
        }
    }
}