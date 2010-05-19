package jtv;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

/**
 * User: aturtsevitch
 * Date: May 19, 2010
 * Time: 12:37:28 PM
 */
public class PDTFile
{
    private static final String FILE_START = "JTV 3.x TV Program Data";

    private File file;
    private long size;
    private Map<Long, String> pdtTitles;

    public PDTFile(File file)
    {
        this.file = file;
    }

    public Long read() throws IOException
    {
        pdtTitles = new HashMap<Long, String>();
        DataInputStream in = null;
        try
        {
            in = new DataInputStream(new BufferedInputStream(new FileInputStream(file)));
            //first read FILE_START
            byte[] b = new byte[FILE_START.getBytes().length];
            in.readFully(b);
            in.skipBytes(3);
            int result = 1;
            while (in.available() > 0)
            {
                //2 zero bytes
                int recordSize = Short.reverseBytes(in.readShort());
                byte[] nameBuff = new byte[recordSize];
                in.readFully(nameBuff);
                pdtTitles.put((long) recordSize, new String(nameBuff));
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

    public long getSize()
    {
        return size;
    }

    public Map<Long, String> getPdtTitles()
    {
        return pdtTitles;
    }
}
