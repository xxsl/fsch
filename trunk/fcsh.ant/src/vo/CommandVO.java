package vo;

import java.io.DataOutputStream;
import java.io.IOException;
import java.io.DataInputStream;
import java.nio.ByteBuffer;

/**
 * User: Dookie
 * Date: 18.03.2009
 * Time: 23:14:15
 */
public class CommandVO implements IExternalizable {
    public static final String DEFAULT_COMMAND = "empty";
    public static String AMF_TYPE = "amf.vo::CommandVO";
    public String target;
    public String command;

    public CommandVO() {
    }

    public CommandVO(String target, String command) {
        this.target = target;
        this.command = command;
    }

    public static boolean isClass(String qualifiedName) {
        return AMF_TYPE.equals(qualifiedName);
    }

    public void serialize(DataOutputStream output) throws IOException {
        ByteBuffer buffer = ByteBuffer.allocate(8192);
        buffer.putInt(AMF_TYPE.getBytes(Encodings.out).length);
        buffer.put(AMF_TYPE.getBytes(Encodings.out));
        buffer.putInt(command.getBytes(Encodings.out).length);
        buffer.put(command.getBytes(Encodings.out));
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
        command = new String(commandBuffer, Encodings.in);

        int targetSize = output.readInt();
        byte[] targetBuffer = new byte[targetSize];
        output.readFully(targetBuffer);
        target = new String(targetBuffer, Encodings.in);
    }

    @Override
    public String toString() {
        return "[" + AMF_TYPE + "] target=" + target + ", command=" + command;
    }

}
