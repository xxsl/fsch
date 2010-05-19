package jtv;

import jtv.vo.JChannel;
import jtv.vo.JProgramme;

import java.io.File;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * User: aturtsevitch
 * Date: May 19, 2010
 * Time: 12:21:07 PM
 */
public class JFileChannel
{
    private File folder;
    private JChannel channel;
    private String name;


    public JFileChannel(File folder, String name)
    {
        this.folder = folder;
        this.name = name;
    }

    public JFileChannel(File folder, JChannel channel)
    {
        this(folder, channel.getName());
        this.channel = channel;
    }

    public JChannel read() throws IOException
    {
        ArrayList<JProgramme> programmes = new ArrayList<JProgramme>();
        JChannel jChannel = new JChannel(name, programmes);

        NDXFile ndxFile = new NDXFile(new File(folder, getNdxName()));
        PDTFile pdtFile = new PDTFile(new File(folder, getPdtName()));

        if (ndxFile.read() > 0 && pdtFile.read() > 0)
        {
            List<NDXTime> times = ndxFile.getNdxTimes();
            Map<Long, String> names = pdtFile.getPdtTitles();

            for (NDXTime time : times)
            {
                programmes.add(new JProgramme(names.get(time.getOffset()), new Date(time.getTime())));
            }
        }
        return jChannel;
    }

    public void write() throws IOException
    {
        NDXFile ndxFile = new NDXFile(new File(folder, getNdxName()));
        PDTFile pdtFile = new PDTFile(new File(folder, getPdtName()));

        List<NDXTime> times = new ArrayList<NDXTime>();
        short offset = PDTFile.FILE_OFFSET;
        for(JProgramme jProgramme:channel.getProgrammes())
        {
            offset += 2 + jProgramme.getName().getBytes("Cp1251").length;
            times.add(new NDXTime(offset, jProgramme.getStart().getTime()));
        }

        ndxFile.write();
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

    private String getPdtName()
    {
        return name + ".pdt";
    }

    private String getNdxName()
    {
        return name + ".ndx";
    }
}
