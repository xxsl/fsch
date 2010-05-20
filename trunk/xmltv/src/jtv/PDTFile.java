package jtv;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * User: aturtsevitch
 * Date: May 19, 2010
 * Time: 12:37:28 PM
 */
public class PDTFile
{
    private static final String FILE_START = "JTV 3.x TV Program Data";
    public static final Short FILE_OFFSET = (short) FILE_START.length();

    private File file;
    private long size;
    private Map<Long, String> pdtTitles;
    private List<String> titles;

    public PDTFile(File folder, String name)
    {
        this.file = new File(folder, getPdtName(name));
    }

    public PDTFile(File file, String name, List<String> titles)
    {
        this(file, name);
        this.titles = titles;
        this.size = titles.size();
    }

    public Long read() throws IOException
    {
        pdtTitles = new HashMap<Long, String>();
        LEDataInputStream in = null;
        try
        {
            in = new LEDataInputStream(new BufferedInputStream(new FileInputStream(file)));
            int offset = 0;
            size = 0;
            //first read FILE_START
            int fileOffset = 3 + FILE_START.getBytes().length;
            in.skipBytes(fileOffset);
            offset += fileOffset;

            while (in.available() > 0)
            {
                int recordSize = in.readShort();
                byte[] nameBuff = new byte[recordSize];
                in.readFully(nameBuff);
                pdtTitles.put((long) offset, new String(nameBuff, "Cp1251"));//todo options
                offset += recordSize + 2;
                size++;
            }
        }
        finally
        {
            if (in != null)
            {
                in.close();
            }
        }
        return size;
    }

    public void write() throws IOException
    {
        if(file.exists() && !file.delete())
        {
            //todo
        }
        LEDataOutputStream out = null;
        try
        {
            out = new LEDataOutputStream(new BufferedOutputStream(new FileOutputStream(file)));
            //first 2 bytes is number of records
            out.write(FILE_START.getBytes());
            out.writeByte(160);
            out.writeByte(160);
            out.writeByte(160);
            for (int i = 0; i < size; i++)
            {
                String title = titles.get(i);
                out.writeShort(title.getBytes("Cp1251").length);
                out.write(title.getBytes("Cp1251"));
            }
            out.flush();
        }
        finally
        {
            if (out != null)
            {
                out.close();
            }
        }
    }

    public long getSize()
    {
        return size;
    }

    public Map<Long, String> getOffset2Title()
    {
        return pdtTitles;
    }

    public List<String> getTitles()
    {
        return titles;
    }

    private String getPdtName(String name)
    {
        return name + ".pdt";
    }
}
