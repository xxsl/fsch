package vo;

import java.io.DataOutputStream;
import java.io.IOException;
import java.io.DataInputStream;
import java.nio.ByteBuffer;

/**
 * User: Dookie
 * Date: 19.03.2009
 * Time: 0:32:16
 */
public class ErrorVO implements IExternalizable {
    public static String AMF_TYPE = "amf.vo::ErrorVO";
    public String description;
    public int id;

    public ErrorVO() {
    }

    public ErrorVO(String description, int id) {
        this.description = description;
        this.id = id;
    }

    public static boolean isClass(String qualifiedName) {
        return AMF_TYPE.equals(qualifiedName);
    }

    public void serialize(DataOutputStream output) throws IOException {
        ByteBuffer buffer = ByteBuffer.allocate(8192);
        buffer.putInt(AMF_TYPE.getBytes(Encodings.out).length);
        buffer.put(AMF_TYPE.getBytes(Encodings.out));
        buffer.putInt(description.getBytes(Encodings.out).length);
        buffer.put(description.getBytes(Encodings.out));
        buffer.putInt(id);
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
        description = new String(commandBuffer, Encodings.in);

        id = output.readInt();
    }

    @Override
    public String toString() {
        return "[" + AMF_TYPE + "] id=" + id + ", description=" + description;
    }
}
