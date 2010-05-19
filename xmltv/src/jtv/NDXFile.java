package jtv;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * User: aturtsevitch
 * Date: May 19, 2010
 * Time: 12:37:36 PM
 */
public class NDXFile
{
    private File file;
    private long size;
    private List<NDXTime> ndxTimes;

    public NDXFile(File file)
    {
        this.file = file;
    }

    public NDXFile(File file, List<NDXTime> ndxTimes)
    {
        this(file);
        this.ndxTimes = ndxTimes;
        this.size = ndxTimes.size();
    }

    public Long read() throws IOException
    {
        ndxTimes = new ArrayList<NDXTime>();
        LEDataInputStream in = null;
        try
        {
            in = new LEDataInputStream(new BufferedInputStream(new FileInputStream(file)));
            //first 2 bytes is number of records
            size = in.readShort();

            for (long i = 0; i < size; i++)
            {
                //2 zero bytes
                in.skipBytes(2);
                long time = FileTimes.getJavaTime(in.readLong());
                short offset = in.readShort();
                NDXTime ndxTime = new NDXTime(offset, time);
                ndxTimes.add(ndxTime);
            }
        }
        finally
        {
            if(in != null)
            {
                in.close();
            }
        }
        return size;
    }

    public void write() throws IOException
    {
        LEDataOutputStream out = null;
        try
        {
            out = new LEDataOutputStream(new BufferedOutputStream(new FileOutputStream(file)));
            //first 2 bytes is number of records
            out.writeShort((short)size);
            for (int i = 0; i < size; i++)
            {
                //2 zero bytes
                out.writeShort((short)0);
                NDXTime ndxTime = ndxTimes.get(i);
                out.writeLong(FileTimes.getWindowsTime(ndxTime.getTime()));
                out.writeShort(ndxTime.getOffset().shortValue());
            }
            out.flush();
        }
        finally
        {
            if(out != null)
            {
                out.close();
            }
        }
    }

    public Long getSize()
    {
        return size;
    }

    public List<NDXTime> getNdxTimes()
    {
        return ndxTimes;
    }
}
