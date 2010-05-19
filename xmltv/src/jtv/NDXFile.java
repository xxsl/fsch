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

    public Long read() throws IOException
    {
        ndxTimes = new ArrayList<NDXTime>();
        DataInputStream in = null;
        try
        {
            in = new DataInputStream(new BufferedInputStream(new FileInputStream(file)));
            //first 2 bytes is number of records
            size = Short.reverseBytes(in.readShort());
            for (long i = 0; i < size; i++)
            {
                //in.skipBytes(12);
                //2 zero bytes
                in.skipBytes(2);
                //      : dwHighDateTime : 30076942 : Long : converter.shiftFile
                //      : dwLowDateTime : 1746311168 : Long : converter.shiftFile
                long time = Long.reverseBytes(in.readLong());
                long offset = Short.reverseBytes(in.readShort());
                NDXTime ndxTime = new NDXTime(offset, time);
                ndxTimes.add(ndxTime);
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

    public Long getSize()
    {
        return size;
    }

    public List<NDXTime> getNdxTimes()
    {
        return ndxTimes;
    }
}
