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
    public static String AMF_TYPE = "amf.vo::CommandVO";
    public static String charset = "UTF-16LE";
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
        buffer.putInt(AMF_TYPE.getBytes(charset).length);
        buffer.put(AMF_TYPE.getBytes(charset));
        buffer.putInt(command.getBytes(charset).length);
        buffer.put(command.getBytes(charset));
        buffer.putInt(target.getBytes(charset).length);
        buffer.put(target.getBytes(charset));
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
        command = new String(commandBuffer, charset);

        int targetSize = output.readInt();
        byte[] targetBuffer = new byte[targetSize];
        output.readFully(targetBuffer);
        target = new String(targetBuffer, charset);
    }

    @Override
    public String toString() {
        return "[" + AMF_TYPE + "] target=" + target + ", command=" + command;
    }

}
