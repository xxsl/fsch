package amf.network {
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
            trimBuffer();
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
            trimBuffer();
            return result;
        }

        override public function readFloat():Number
        {
            var result:Number = _buffer.readFloat();
            trimBuffer();
            return result;
        }

        override public function readBoolean():Boolean
        {
            var result:Boolean = _buffer.readBoolean();
            trimBuffer();
            return result;
        }

        override public function readInt():int
        {
            var result:int = _buffer.readInt();
            trimBuffer();
            return result;
        }

        override public function readUTFBytes(length:uint):String
        {
            var result:String = _buffer.readUTFBytes(length);
            trimBuffer();
            return result;
        }

        override public function readByte():int
        {
            var result:int = _buffer.readByte();
            trimBuffer();
            return result;
        }

        private function trimBuffer():void
        {
            var newBuffer:ByteArray = new ByteArray();
            _buffer.readBytes(newBuffer);
            _buffer = newBuffer;
        }

        /*mthods inherited from object*/


        override AS3 function hasOwnProperty(V:* = null):Boolean
        {
            return _buffer.hasOwnProperty(V);
        }

        override AS3 function propertyIsEnumerable(V:* = null):Boolean
        {
            return _buffer.propertyIsEnumerable(V);
        }

        override AS3 function isPrototypeOf(V:* = null):Boolean
        {
            return _buffer.isPrototypeOf(V);
        }

        /* next methods not implemented*/


        override public function writeUTFBytes(value:String):void
        {
            if (true) {
                throw new Error("Not implemented");
            }
            super.writeUTFBytes(value);
        }

        override public function readObject():*
        {
            if (true) {
                throw new Error("Not implemented");
            }
            return super.readObject();
        }

        override public function writeObject(object:*):void
        {
            if (true) {
                throw new Error("Not implemented");
            }
            super.writeObject(object);
        }

        override public function readShort():int
        {
            if (true) {
                throw new Error("Not implemented");
            }
            return super.readShort();
        }

        override public function writeDouble(value:Number):void
        {
            if (true) {
                throw new Error("Not implemented");
            }
            super.writeDouble(value);
        }

        override public function readUnsignedShort():uint
        {
            if (true) {
                throw new Error("Not implemented");
            }
            return super.readUnsignedShort();
        }

        override public function get endian():String
        {
            if (true) {
                throw new Error("Not implemented");
            }
            return super.endian;
        }

        override public function readDouble():Number
        {
            if (true) {
                throw new Error("Not implemented");
            }
            return super.readDouble();
        }

        override public function set endian(type:String):void
        {
            if (true) {
                throw new Error("Not implemented");
            }
            super.endian = type;
        }

        override public function readUTF():String
        {
            if (true) {
                throw new Error("Not implemented");
            }
            return super.readUTF();
        }

        override public function readUnsignedInt():uint
        {
            if (true) {
                throw new Error("Not implemented");
            }
            return super.readUnsignedInt();
        }

        override public function writeUTF(value:String):void
        {
            if (true) {
                throw new Error("Not implemented");
            }
            super.writeUTF(value);
        }

        override public function get objectEncoding():uint
        {
            if (true) {
                throw new Error("Not implemented");
            }
            return super.objectEncoding;
        }

        override public function readUnsignedByte():uint
        {
            if (true) {
                throw new Error("Not implemented");
            }
            return super.readUnsignedByte();
        }

        override public function writeUnsignedInt(value:uint):void
        {
            if (true) {
                throw new Error("Not implemented");
            }
            super.writeUnsignedInt(value);
        }

        override public function writeShort(value:int):void
        {
            if (true) {
                throw new Error("Not implemented");
            }
            super.writeShort(value);
        }

        override public function toString():String
        {
            if (true) {
                throw new Error("Not implemented");
            }
            return super.toString();
        }

        override public function set length(value:uint):void
        {
            if (true) {
                throw new Error("Not implemented");
            }
            super.length = value;
        }

        override public function set objectEncoding(version:uint):void
        {
            if (true) {
                throw new Error("Not implemented");
            }
            super.objectEncoding = version;
        }

    }
}