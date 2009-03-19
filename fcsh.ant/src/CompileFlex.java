import org.apache.tools.ant.Task;
import org.apache.tools.ant.BuildException;

import java.net.Socket;
import java.net.SocketAddress;
import java.net.InetSocketAddress;
import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import vo.CommandVO;
import vo.ErrorVO;
import vo.DataVO;

/**
 * User: Dookie
 * Date: 18.03.2009
 * Time: 22:48:21
 */

public class CompileFlex extends Task {
    private List<Arg> args = new ArrayList<Arg>();

    public void addARG(Arg arg) {
        args.add(arg);
    }

    public Object createARG(){
        return new Arg();
    }

    public static void main(String args[]) {
        new CompileFlex().execute();
    }

    // The method executing the task
    public void execute() throws BuildException {

        System.out.println(((Arg)args.get(2)).getName());
        DataOutputStream os;
        DataInputStream is;

        try {
            Socket socket = new Socket();
            SocketAddress socketAddress = new InetSocketAddress(40000);
            socket.connect(socketAddress, 5000);
            os = new DataOutputStream(socket.getOutputStream());
            is = new DataInputStream(socket.getInputStream());

            CommandVO startFcshCommand = new CommandVO("fcsh_start", "empty");

            startFcshCommand.serialize(os);

            os.flush();

            Object responce = readResponce(is);
            if ((responce instanceof ErrorVO) && ((ErrorVO) responce).id != 4) {
                throw new BuildException(responce.toString());
            } else if ((responce instanceof ErrorVO) || ((responce instanceof DataVO) && ((DataVO) responce).target.equals("fcsh_start"))) {
                compile(os);
            } else {
                throw new BuildException("Build failed: " + responce.toString());
            }

            /*responce = readResponce(is);

            if (responce instanceof ErrorVO) {
                throw new BuildException(responce.toString());
            } else {
                //mxmlc -output=C:\\realworld.swf -load-config+=C:\work\realworld\FLX\src\flex-config.xml
                CommandVO compileCommand = new CommandVO("fcsh", "mxmlc -output=C:\\\\realworld.swf -load-config+=C:\\work\\realworld\\FLX\\src\\flex-config.xml");
                compileCommand.serialize(os);
                os.flush();
            }*/

            responce = readResponce(is);

            if (responce instanceof ErrorVO) {
                throw new BuildException(responce.toString());
            } else if (responce instanceof DataVO) {
                DataVO dataVO = (DataVO) responce;
                System.out.println(dataVO.data);
            } else {
                throw new BuildException("Build failed: " + responce.toString());
            }


            socket.close();
        }
        catch (Exception e) {
            throw new BuildException(e);
        }
    }

    private void compile(DataOutputStream os) throws IOException {
        System.out.println("mxmlc -output=C:\\\\realworld.swf -load-config+=C:\\work\\realworld\\FLX\\src\\flex-config.xml");
        CommandVO compileCommand = new CommandVO("fcsh", "mxmlc -output=C:\\\\realworld.swf -load-config+=C:\\work\\realworld\\FLX\\src\\flex-config.xml");
        compileCommand.serialize(os);
        os.flush();
    }

    private Object readResponce(DataInputStream is) throws IOException {
        int size = is.readInt();
        //System.out.println("Object size " + size);


        int classSize = is.readInt();
        byte[] name = new byte[classSize];
        is.readFully(name);
        String className = new String(name, "UTF-16LE");

        //System.out.println("Object is " + className);

        if (CommandVO.isClass(className)) {
            CommandVO commandVO = new CommandVO();
            commandVO.deSerialize(is);
            //System.out.println("Command: " + commandVO.toString());
            return commandVO;
        } else if (ErrorVO.isClass(className)) {
            ErrorVO errorVO = new ErrorVO();
            errorVO.deSerialize(is);
            //System.out.println("Error: " + errorVO.toString());
            return errorVO;
        } else if (DataVO.isClass(className)) {
            DataVO dataVO = new DataVO();
            dataVO.deSerialize(is);
            //System.out.println("Data: " + dataVO.toString());
            return dataVO;
        } else {
            throw new BuildException("Unknown object: " + className);
        }

    }
}

