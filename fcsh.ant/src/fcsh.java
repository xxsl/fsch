import org.apache.tools.ant.Task;
import org.apache.tools.ant.BuildException;
import org.apache.tools.ant.taskdefs.Echo;

import java.net.Socket;
import java.net.SocketAddress;
import java.net.InetSocketAddress;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

import vo.CommandVO;
import vo.ErrorVO;
import vo.DataVO;
import vo.Encodings;
import constants.BUILD;

/**
 * User: Dookie
 * Date: 18.03.2009
 * Time: 22:48:21
 */

public class fcsh extends Task {
    private List<Arg> args = new ArrayList<Arg>();

    private String consoleEncoding = "cp866";

    public String getConsoleEncoding() {
        return consoleEncoding;
    }

    public void setConsoleEncoding(String consoleEncoding) {
        this.consoleEncoding = consoleEncoding;
    }

//    public void addArg(Arg argument) {
//        args.add(argument);
//    }

    public Object createArg() {
        Arg aNewArg = new Arg();
        args.add(aNewArg);
        return aNewArg;
    }

    public static void main(String args[]) {
        new fcsh().execute();
    }

    public void execute() throws BuildException {
        SocketAddress socketAddress = new InetSocketAddress(40000);
        DataOutputStream os;
        DataInputStream is;
        Socket socket = new Socket();

        try {
            socket.connect(socketAddress, 60000);
        } catch (IOException e) {
            System.out.println("Server is not responding. Probably it is stopped. Trying to launch...");
            tryToLaunchServer(e);
        }

        socket = tryToestablishConnection(socket, socketAddress, 5);


        try {
            os = new DataOutputStream(socket.getOutputStream());
            is = new DataInputStream(socket.getInputStream());
        }
        catch (IOException io) {
            throw new BuildException(io);
        }

        try {
            CommandVO startFcshCommand = new CommandVO("fcsh_start", CommandVO.DEFAULT_COMMAND);

            startFcshCommand.serialize(os);

            os.flush();

            Object responce = readResponce(is);

            //fcsh error: busy or stopped
            if ((responce instanceof ErrorVO) && ((ErrorVO) responce).id != 4) {
                throw new BuildException(responce.toString());
                //fcsh is already running or started => compile
            }
            else if ((responce instanceof ErrorVO) || ((responce instanceof DataVO) && ((DataVO) responce).target.equals("fcsh_start"))) {
                compile(os);
                //any other object is error
            }
            else {
                throw new BuildException("Build failed: " + responce.toString());
            }

            responce = readResponce(is);

            //something happend?
            if (responce instanceof ErrorVO) {
                throw new BuildException(responce.toString());
                //check build result
            }
            else if (responce instanceof DataVO) {
                DataVO dataVO = (DataVO) responce;
                printRU(dataVO.data);
                System.out.println("");
                if (BUILD.FCSH_BUILD_ERROR.equals(dataVO.target)) {
                    throw new BuildException("Total crap...");
                }
                else if (BUILD.FCSH_BUILD_WARNING.equals(dataVO.target)) {
                    System.out.println("Fix this warnings... Dude!");
                }
                else if (BUILD.FCSH_BUILD_SUCCESSFULL.equals(dataVO.target)) {
                    System.out.println("Awesome!");
                }
                else if ("fcsh_stop".equals(dataVO.target)) {
                    System.out.println("Flex Compile SHell failed. Check your server.ini");
                }
                else {
                    System.out.println("WTF?!");
                }
                //any other is error
            }
            else {
                throw new BuildException("Build failed: " + responce.toString());
            }

            socket.close();
        }
        catch (IOException e) {
            throw new BuildException(e);
        }
    }

    private Socket tryToestablishConnection(Socket socket, SocketAddress socketAddress, int attempts) {
        if (!socket.isConnected()) {
            for (int i = 0; i < attempts; i++) {
                try {
                    System.out.println("Trying to connect... Attempt " + i + " of " + attempts);
                    socket = new Socket();
                    socket.connect(socketAddress, 60000);
                } catch (IOException e) {
                    System.out.println("Pause 2 seconds...");
                    try {
                        synchronized (this) {
                            wait(2000);
                        }
                    } catch (InterruptedException e1) {
                        throw new BuildException(e1);
                    }
                }
                if (socket.isConnected()) {
                    System.out.println("Server is up!");
                    break;
                }
            }
        }
        return socket;
    }

    private void tryToLaunchServer(IOException e) {
        String executable = System.getenv("FCSHServer");
        if (executable != null) {
            Runtime rt = Runtime.getRuntime();
            try {
                rt.exec(executable + "\\FCSHServer.exe");
            } catch (IOException e1) {
                throw new BuildException("Cant start Server", e1);
            }
            System.out.println("Server started");
        }
        else {
            throw new BuildException("Cant start Server, environment variable {FCSHServer} is not set.", e);
        }
    }

    private void compile(DataOutputStream os) throws IOException {
        String cmd = "";
        for (Arg argument : args) {
            cmd += argument.getValue() + " ";
        }
        System.out.println("Command: " + cmd);
        CommandVO compileCommand = new CommandVO("fcsh", cmd);
        compileCommand.serialize(os);
        os.flush();
    }

    private Object readResponce(DataInputStream is) throws IOException {
        int size = is.readInt();

        int classSize = is.readInt();
        byte[] name = new byte[classSize];
        is.readFully(name);
        String className = new String(name, Encodings.in);


        if (CommandVO.isClass(className)) {
            CommandVO commandVO = new CommandVO();
            commandVO.deSerialize(is);
            return commandVO;
        }
        else if (ErrorVO.isClass(className)) {
            ErrorVO errorVO = new ErrorVO();
            errorVO.deSerialize(is);
            return errorVO;
        }
        else if (DataVO.isClass(className)) {
            DataVO dataVO = new DataVO();
            dataVO.deSerialize(is);
            return dataVO;
        }
        else {
            throw new BuildException("Unknown object: " + className);
        }
    }

    private void printRU(String javaString) {
        try {
            // output to the console
            Writer w =
                    new BufferedWriter
                            (new OutputStreamWriter(System.out, getConsoleEncoding()));
            w.write(javaString);
            w.flush();
            //w.close();
        }
        catch (Exception e) {
            throw new BuildException(e);
        }
    }
}

