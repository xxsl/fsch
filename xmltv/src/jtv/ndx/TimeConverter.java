/*
 * Copyright(c) Nimrod97, 2010.
 *
 * Email: Nimrod97@gmail.com
 * Project: http://code.google.com/p/xmltv2jtv/
 */

package jtv.ndx;

/**
 * Utility class to convert between Java and Windows times.
 */
public final class TimeConverter
{
    /**
     * Java timestamps use 64-bit milliseconds since 1970 GMT. Windows timestamps use 64-bit value representing the
     * number of 100-nanosecond intervals since January 1, 1601, with ten thousand times as much precision. <br>
     * DIFF_IN_MILLIS is the difference between January 1 1601 and January 1 1970 in milliseconds. This magic number
     * came from com.mindprod.common11.TestDate. Done according to Gregorian Calendar, no correction for 1752-09-02
     * Wednesday was followed immediately by 1752-09-14 Thursday dropping 12 days. Also according to
     * http://gcc.gnu.org/ml/java-patches/2003-q1/msg00565.html
     */
    private static final long DIFF_IN_MILLIS = 11644473600000L;

    /**
     * Returns Java milliseconds.
     * @param windowsTime Windows milliseconds.
     * @return Java milliseconds.
     */
    public static long getJavaTime(long windowsTime)
    {
        return (windowsTime / 10000) - DIFF_IN_MILLIS;
    }

    /**
     * Returns Windows milliseconds.
     * @param javaTime Java milliseconds.
     * @return Windows milliseconds.
     */
    public static long getWindowsTime(long javaTime)
    {
        return (javaTime + DIFF_IN_MILLIS) * 10000;
    }
}
