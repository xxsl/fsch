package jtv;

import jtv.ndx.NDXFile;
import jtv.ndx.NDXTime;
import jtv.pdt.PDTFile;
import jtv.vo.JChannel;
import jtv.vo.JProgramme;

import java.io.File;
import java.io.IOException;
import java.util.*;

public class JFileChannel
{
    private File folder;
    private JChannel channel;
    private String name;
    private String charsetName;


    public JFileChannel(File folder, String name, String charsetName)
    {
        this.folder = folder;
        this.name = name;
        this.charsetName = charsetName;
    }

    public JFileChannel(File folder, JChannel channel, String charsetName)
    {
        this(folder, channel.getName(), charsetName);
        this.channel = channel;
    }

    public JChannel read() throws IOException
    {
        checkFolderExists();

        ArrayList<JProgramme> programmes = new ArrayList<JProgramme>();
        JChannel jChannel = new JChannel(name, programmes);

        NDXFile ndxFile = new jtv.ndx.NDXFile(folder, name);
        PDTFile pdtFile = new PDTFile(folder, name, charsetName);

        if (ndxFile.read() > 0 && pdtFile.read() > 0)
        {
            List<NDXTime> times = ndxFile.getNdxTimes();
            Map<Long, String> names = pdtFile.getOffset2Title();

            for (NDXTime time : times)
            {
                programmes.add(new JProgramme(names.get((long)time.getOffset()), new Date(time.getTime())));
            }
        }
        return jChannel;
    }

    public void write() throws IOException
    {
        checkFolderExists();

        List<NDXTime> times = new ArrayList<NDXTime>();
        List<String> titles = new ArrayList<String>();
        short offset = (short)(PDTFile.FILE_OFFSET + 3);

        //optimize pdt size, remove repeated titles
        Map<String, Short> dateStringMap = new HashMap<String, Short>();

        for(JProgramme jProgramme:channel.getProgrammes())
        {
            if(dateStringMap.get(jProgramme.getName()) == null)
            {
                offset += 2 + jProgramme.getName().getBytes(charsetName).length;
                titles.add(jProgramme.getName());
                dateStringMap.put(jProgramme.getName(), offset);
            }
            times.add(new NDXTime(dateStringMap.get(jProgramme.getName()), jProgramme.getStart().getTime()));
        }

        NDXFile ndxFile = new NDXFile(folder, name, times);
        ndxFile.write();

        PDTFile pdtFile = new PDTFile(folder, name, charsetName, titles);
        pdtFile.write();
    }

    public File getFolder()
    {
        return folder;
    }

    public JChannel getChannel()
    {
        return channel;
    }

    public String getName()
    {
        return name;
    }

    private void checkFolderExists() throws IOException
    {
        if(!folder.exists())
        {
            throw new IOException("JTV folder must exist: " + folder.getPath()); 
        }
    }
}
