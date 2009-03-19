package vo;

import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.io.IOException;
import java.nio.ByteBuffer;

/**
 * User: Dookie
 * Date: 18.03.2009
 * Time: 23:14:15
 */
public class DataVO implements IExternalizable {
    public static String AMF_TYPE = "amf.vo::DataVO";
    public String target;
    public String data;

    public DataVO() {
    }

    public DataVO(String target, String data) {
        this.target = target;
        this.data = data;
    }

    public static boolean isClass(String qualifiedName) {
        return AMF_TYPE.equals(qualifiedName);
    }

    public void serialize(DataOutputStream output) throws IOException {
        ByteBuffer buffer = ByteBuffer.allocate(8192);
        buffer.putInt(AMF_TYPE.getBytes(Encodings.out).length);
        buffer.put(AMF_TYPE.getBytes(Encodings.out));
        buffer.putInt(data.getBytes(Encodings.out).length);
        buffer.put(data.getBytes(Encodings.out));
        buffer.putInt(target.getBytes(Encodings.out).length);
        buffer.put(target.getBytes(Encodings.out));
        int size = buffer.limit() - buffer.remaining();
        output.writeInt(size);
        output.flush();
        buffer.flip();
        output.write(buffer.array(), buffer.position(), buffer.remaining());
    }

    public void deSerialize(DataInputStream output) throws IOException {
        int commandSize = output.readInt();
        byte[] commandBuffer = new byte[commandSize];
        output.readFully(commandBuffer);
        data = new String(commandBuffer, Encodings.in);

        int targetSize = output.readInt();
        byte[] targetBuffer = new byte[targetSize];
        output.readFully(targetBuffer);
        target = new String(targetBuffer, Encodings.in);
    }

    @Override
    public String toString() {
        return "[" + AMF_TYPE + "] target=" + target + ", data=" + data;
    }

}
