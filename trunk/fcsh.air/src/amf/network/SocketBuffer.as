package amf.network {
    import flash.utils.ByteArray;
    import flash.utils.ByteArray;
    import flash.utils.ByteArray;
import flash.utils.IDataInput;
    import flash.utils.IDataOutput;

    /**
     *@author aturtsevitch
     *@date   12.03.2009
     *@time   17:19:48
     *@langversion ActionScript 3.0
     */
    public class SocketBuffer extends ByteArray implements IDataInput, IDataOutput
    {
        private var _buffer:ByteArray = new ByteArray();

        public function SocketBuffer()
        {
            super();
        }


        override public function writeBytes(bytes:ByteArray, offset:uint = 0, length:uint = 0):void
        {
            _buffer.position = _buffer.length;
            _buffer.writeBytes(bytes, offset, length);
            _buffer.position = 0;
        }


        override public function readBytes(bytes:ByteArray, offset:uint = 0, length:uint = 0):void
        {
            _buffer.readBytes(bytes, offset, length);
            replaceBuffer();
        }

        override public function get position():uint
        {
            return _buffer.position;
        }


        override public function set position(offset:uint):void
        {
            _buffer.position = offset;
        }

        override public function get length():uint
        {
            return _buffer.length;
        }


        override public function writeMultiByte(value:String, charSet:String):void
        {
            _buffer.writeMultiByte(value, charSet);
        }

        override public function writeByte(value:int):void
        {
            _buffer.writeByte(value);
        }

        override public function writeBoolean(value:Boolean):void
        {
            _buffer.writeBoolean(value);
        }

        override public function writeInt(value:int):void
        {
            _buffer.writeInt(value);
        }

        override public function writeFloat(value:Number):void
        {
            _buffer.writeFloat(value);
        }

        override public function get bytesAvailable():uint
        {
            return _buffer.bytesAvailable;
        }

        override public function readMultiByte(length:uint, charSet:String):String
        {
            var result:String = _buffer.readMultiByte(length, charSet);
            replaceBuffer();
            return result;
        }

        override public function readFloat():Number
        {
            var result:Number = _buffer.readFloat();
            replaceBuffer();
            return result;
        }

        override public function readBoolean():Boolean
        {
            var result:Boolean = _buffer.readBoolean();
            replaceBuffer();
            return result;
        }

        override public function readInt():int
        {
            var result:int = _buffer.readInt();
            replaceBuffer();
            return result;
        }

        override public function readUTFBytes(length:uint):String
        {
            var result:String = _buffer.readUTFBytes(length);
            replaceBuffer();
            return result;
        }

        override public function readByte():int
        {
            var result:int = _buffer.readByte();
            replaceBuffer();
            return result;
        }

        private function replaceBuffer():void
        {
            var newBuffer:ByteArray = new ByteArray();
            _buffer.readBytes(newBuffer);
            _buffer = newBuffer;
        }
    }
}