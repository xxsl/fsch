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
    public static final Short FILE_OFFSET = (short)FILE_START.length();

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

    public long getSize()
    {
        return size;
    }

    public Map<Long, String> getPdtTitles()
    {
        return pdtTitles;
    }
}
