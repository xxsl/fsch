import org.apache.tools.ant.Task;
import org.apache.tools.ant.BuildException;

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
import constants.Build;
import constants.Command;

/**
 * User: Dookie
 * Date: 18.03.2009
 * Time: 22:48:21
 */

public class fcsh extends Task {
    private static final String EXECUTABLE = "FCSHServer.exe";
    private static final String ENVIRONMENT = "FCSHServer";
    private static final int PORT = 40000;

    private List<Arg> args = new ArrayList<Arg>();
    private String consoleEncoding = "cp866";

    public String getConsoleEncoding() {
        return consoleEncoding;
    }

    public void setConsoleEncoding(String consoleEncoding) {
        this.consoleEncoding = consoleEncoding;
    }

    public Object createArg() {
        Arg aNewArg = new Arg();
        args.add(aNewArg);
        return aNewArg;
    }


    public void execute() throws BuildException {
        SocketAddress localhost = new InetSocketAddress(PORT);

        DataOutputStream outputStream;
        DataInputStream inputStream;

        Socket socket = getSocket(localhost, 5);

        try {
            outputStream = new DataOutputStream(socket.getOutputStream());
            inputStream = new DataInputStream(socket.getInputStream());
        }
        catch (IOException io) {
            throw new BuildException("Get input/output stream failed", io);
        }

        try {
            CommandVO startFcshCommand = new CommandVO(Command.FCSH_START, CommandVO.DEFAULT_COMMAND);

            startFcshCommand.serialize(outputStream);

            outputStream.flush();

            Object responce = readResponce(inputStream);

            //fcsh error: busy or stopped
            if ((responce instanceof ErrorVO) && ((ErrorVO) responce).id != 4) {
                throw new BuildException(responce.toString());
                //fcsh is already running or started => compile
            } else if ((responce instanceof ErrorVO) || ((responce instanceof DataVO) && ((DataVO) responce).target.equals(Command.FCSH_START))) {
                compile(outputStream);
                //any other object is error
            } else if(((responce instanceof DataVO) && ((DataVO) responce).target.equals(Command.FCSH_STOP))){
                throw new BuildException("Flex Compile SHell failed to start. Check your server.ini");
            }
            else{
                throw new BuildException("Build failed: " + responce.toString());
            }

            responce = readResponce(inputStream);

            //something happend?
            if (responce instanceof ErrorVO) {
                throw new BuildException(responce.toString());
                //check build result
            } else if (responce instanceof DataVO) {
                DataVO dataVO = (DataVO) responce;
                if (!CommandVO.DEFAULT_COMMAND.equals(dataVO.data)) {
                    printEcoded(dataVO.data);
                }
                System.out.println("");
                if (Build.FCSH_BUILD_ERROR.equals(dataVO.target)) {
                    throw new BuildException("Total crap...");
                } else if (Build.FCSH_BUILD_WARNING.equals(dataVO.target)) {
                    System.out.println("Fix this warnings... Dude!");
                } else if (Build.FCSH_BUILD_SUCCESSFULL.equals(dataVO.target)) {
                    System.out.println("Awesome!");
                } else if (Command.FCSH_STOP.equals(dataVO.target)) {
                    System.out.println("Flex Compile SHell failed to start. Check your server.ini");
                    throw new BuildException("Flex Compile SHell failed to start. Check your server.ini");
                } else {
                    System.out.println("WTF?!:" + responce.toString());
                    throw new BuildException("Unknown responce:" + responce.toString());
                }
                //any other is error
            } else {
                throw new BuildException("Build failed: " + responce.toString());
            }

            socket.close();
        }
        catch (IOException e) {
            throw new BuildException(e);
        }
    }

    private Socket getSocket(SocketAddress socketAddress, int attempts) {
        Socket socket = new Socket();
        try {
            socket.connect(socketAddress);
        } catch (IOException e) {
            System.out.println("Server is not responding. Probably it is stopped. Trying to launch...");
            startServer();
        }
        if (!socket.isConnected()) {
            for (int i = 0; i < attempts; i++) {
                try {
                    System.out.println("Trying to connect... Attempt " + (i + 1) + " of " + attempts);
                    socket = new Socket();
                    socket.connect(socketAddress);
                } catch (IOException e) {
                    System.out.println("Paused for 2 seconds...");
                    try {
                        synchronized (this) {
                            wait(2000);
                        }
                    } catch (InterruptedException e1) {
                        throw new BuildException(e1);
                    }
                }
                if (socket.isConnected()) {
                    System.out.println("Connection established");
                    break;
                }
            }
        }
        return socket;
    }

    private void startServer() {
        String executable = System.getenv(ENVIRONMENT);

        if (executable != null) {
            Runtime runtime = Runtime.getRuntime();
            String command = executable + File.separator + EXECUTABLE;
            try {
                runtime.exec(command);
            } catch (IOException e1) {
                throw new BuildException("Server failed to start, command line was: " + command, e1);
            }
            System.out.println("Server started");
        } else {
            throw new BuildException("Can not start server, environment variable {" + ENVIRONMENT + "} is not set.");
        }
    }

    private void compile(DataOutputStream outputStream) throws IOException {
        String commandLine = "";

        for (Arg argument : args) {
            commandLine += argument.getValue() + " ";
        }
        System.out.println("Send command: " + commandLine);

        CommandVO commandVO = new CommandVO(Command.FCSH_EXEC, commandLine);
        commandVO.serialize(outputStream);
        outputStream.flush();
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
        } else if (ErrorVO.isClass(className)) {
            ErrorVO errorVO = new ErrorVO();
            errorVO.deSerialize(is);
            return errorVO;
        } else if (DataVO.isClass(className)) {
            DataVO dataVO = new DataVO();
            dataVO.deSerialize(is);
            return dataVO;
        } else {
            throw new BuildException("Unknown object: " + className);
        }
    }

    private void printEcoded(String javaString) {
        try {
            Writer writer = new BufferedWriter(new OutputStreamWriter(System.out, getConsoleEncoding()));
            writer.write(javaString);
            writer.flush();
        }
        catch (Exception e) {
            throw new BuildException("Print to console failed", e);
        }
    }
}

