package vo;

import java.io.DataOutputStream;
import java.io.IOException;
import java.io.DataInputStream;

/**
 * User: Dookie
 * Date: 19.03.2009
 * Time: 0:11:39
 */
public interface IExternalizable {
    public void serialize(DataOutputStream output) throws IOException;

    public void deSerialize(DataInputStream output) throws IOException;
}
