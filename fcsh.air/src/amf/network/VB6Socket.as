package amf.network {
    import amf.VB6Deserializer;
    import amf.VB6Serializer;
    import amf.vo.BaloonVO;
    import amf.vo.DataVO;
    import amf.vo.ErrorVO;

    import flash.events.Event;
    import flash.events.IOErrorEvent;
    import flash.events.ProgressEvent;
    import flash.events.SecurityErrorEvent;
    import flash.net.Socket;
    import flash.utils.ByteArray;
    import flash.utils.getQualifiedClassName;

    import processing.RemoteController;

    public class VB6Socket extends Socket
    {
        private var size:int = -1;
        private var buffer:SocketBuffer = new SocketBuffer();

        public function VB6Socket(host:String = null, port:uint = 0) {
            super(host, port);
            configureListeners();
        }

        private function configureListeners():void {
            addEventListener(Event.CLOSE, closeHandler);
            addEventListener(Event.CONNECT, connectHandler);
            addEventListener(IOErrorEvent.IO_ERROR, ioErrorHandler);
            addEventListener(SecurityErrorEvent.SECURITY_ERROR, securityErrorHandler);
            addEventListener(ProgressEvent.SOCKET_DATA, socketDataHandler);
        }

        private function closeHandler(event:Event):void {
            fcsh.instance.log("Socket closed");
            fcsh.instance.setConnected(false);
        }

        private function connectHandler(event:Event):void {
            fcsh.instance.log("Socket connected");
            fcsh.instance.setConnected(true);
        }

        private function ioErrorHandler(event:IOErrorEvent):void {
            fcsh.instance.log(event.text);
            fcsh.instance.setConnected(false);
        }

        private function securityErrorHandler(event:SecurityErrorEvent):void {
            fcsh.instance.log(event.text);
            fcsh.instance.setConnected(false);
        }

        private function socketDataHandler(event:ProgressEvent):void {
            fcsh.instance.log("Socket data arrival, bytes loaded: " + event.bytesLoaded);
            var _buffer:ByteArray = new ByteArray();
            readBytes(_buffer, 0, event.bytesLoaded);
            buffer.writeBytes(_buffer, 0, _buffer.length);

            if (size == -1 && buffer.length >= 4)
            {
                size = buffer.readInt();
                fcsh.instance.log("Object size: " + size);
            }

            while (buffer.length >= size && (size != -1))
            {
                fcsh.instance.log("Read object: " + size + " bytes");
                var object:Object = VB6Deserializer.deserialize(buffer);
                fcsh.instance.log("Object is: " + getQualifiedClassName(object));
                RemoteController.instance.process(object);
                fcsh.instance.log("Read complete");
                size = -1;
                if (buffer.length >= 4)
                {
                    size = buffer.readInt();
                    fcsh.instance.log("Object size: " + size);
                }
            }
            fcsh.instance.log("Socket data processed. size " + size + " bytes , buffer.length " + buffer.length + "\n");
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