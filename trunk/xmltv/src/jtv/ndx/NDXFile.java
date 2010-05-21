package jtv.ndx;

import jtv.bigendian.LEDataInputStream;
import jtv.bigendian.LEDataOutputStream;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class NDXFile
{
    private File file;
    private long size;
    private List<jtv.ndx.NDXTime> ndxTimes;

    public NDXFile(File folder, String name)
    {
        this.file = new File(folder, getNdxName(name));
    }

    public NDXFile(File folder, String name, List<jtv.ndx.NDXTime> ndxTimes)
    {
        this(folder, name);
        this.ndxTimes = ndxTimes;
        this.size = ndxTimes.size();
    }

    public Long read() throws IOException
    {
        ndxTimes = new ArrayList<jtv.ndx.NDXTime>();
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
                long time = TimeConverter.getJavaTime(in.readLong());
                short offset = in.readShort();
                jtv.ndx.NDXTime ndxTime = new jtv.ndx.NDXTime(offset, time);
                ndxTimes.add(ndxTime);
            }
        }
        finally
        {
            closeQuetly(in);
        }
        return size;
    }

    public void write() throws IOException
    {
        if(file.exists() && !file.delete())
        {
            throw new IOException("Unable to delete file " + file.getPath());
        }

        LEDataOutputStream out = null;
        try
        {
            out = new jtv.bigendian.LEDataOutputStream(new BufferedOutputStream(new FileOutputStream(file)));
            //first 2 bytes is number of records
            out.writeShort((short)size);
            for (int i = 0; i < size; i++)
            {
                //2 zero bytes
                out.writeShort((short)0);
                jtv.ndx.NDXTime ndxTime = ndxTimes.get(i);
                out.writeLong(TimeConverter.getWindowsTime(ndxTime.getTime()));
                out.writeShort(ndxTime.getOffset().shortValue());
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
                //ignore
            }
        }
    }

    public Long getSize()
    {
        return size;
    }

    public List<jtv.ndx.NDXTime> getNdxTimes()
    {
        return ndxTimes;
    }

    private String getNdxName(String name)
    {
        return name + ".ndx";
    }
}
