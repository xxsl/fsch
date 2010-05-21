/*
 * Copyright(c) Nimrod97, 2010.
 *
 * Email: Nimrod97@gmail.com
 * Project: http://code.google.com/p/xmltv2jtv/
 */

package jtv.pdt;

import jtv.bigendian.LEDataInputStream;
import jtv.bigendian.LEDataOutputStream;
import org.apache.log4j.Logger;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class PDTFile
{
    private static final Logger LOGGER = Logger.getLogger(PDTFile.class.getName());

    private static final String FILE_START = "JTV 3.x TV Program Data";
    public static final Short FILE_OFFSET = (short) FILE_START.length();

    private File file;
    private long size;
    private String charsetName;
    private Map<Long, String> pdtTitles;
    private List<String> titles;

    public PDTFile(File folder, String name, String charsetName)
    {
        this.file = new File(folder, getPdtName(name));
        this.charsetName = charsetName;
    }

    public PDTFile(File file, String name, String charsetName, List<String> titles)
    {
        this(file, name, charsetName);
        this.titles = titles;
        this.size = titles.size();
    }

    public Long read() throws IOException
    {
        LOGGER.debug("Parsing pdt file " + file.getPath());

        pdtTitles = new HashMap<Long, String>();
        LEDataInputStream in = null;
        try
        {
            in = new LEDataInputStream(new BufferedInputStream(new FileInputStream(file)));
            size = 0;
            //first skip FILE_START bytes + 3 bytes A0
            int offset = 3 + FILE_START.getBytes().length;
            in.skipBytes(offset);

            while (in.available() > 0)
            {
                //title size
                int recordSize = in.readShort();
                byte[] nameBuff = new byte[recordSize];
                in.readFully(nameBuff);
                pdtTitles.put((long) offset, new String(nameBuff, charsetName));
                offset += recordSize + 2;
                size++;
            }
            LOGGER.debug("Successfully parsed pdt file: " + file.getPath() + ", programs: " + size);
        }
        finally
        {
            closeQuetly(in);
        }
        return size;
    }

    public void write() throws IOException
    {
        if (file.exists() && !file.delete())
        {
            throw new IOException("Unable to delete file " + file.getPath());
        }
        jtv.bigendian.LEDataOutputStream out = null;
        try
        {
            out = new LEDataOutputStream(new BufferedOutputStream(new FileOutputStream(file)));
            //write FILE_START
            out.write(FILE_START.getBytes());
            //3 bytes A0
            out.writeByte(160);
            out.writeByte(160);
            out.writeByte(160);
            for (int i = 0; i < size; i++)
            {
                byte[] title = titles.get(i).getBytes(charsetName);
                //write title length
                out.writeShort(title.length);
                //write title bytes
                out.write(title);
            }
            out.flush();
        }
        finally
        {
            closeQuetly(out);
        }
    }

    private void closeQuetly(Closeable out)
    {
        if (out != null)
        {
            try
            {
                out.close();
            }
            catch (IOException e)
            {
                LOGGER.warn("Close stream error" + e);
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

    private String getPdtName(String name)
    {
        return name + ".pdt";
    }
}
