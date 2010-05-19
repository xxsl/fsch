package jtv;

public final class FileTimes
{
    /**
     * 86,400,000 the number of milliseconds in 24 hour day. Easily fits into an int.
     */
    private static final int MILLISECONDS_PER_DAY = 24 * 60 * 60 * 1000;

    /**
     * Java timestamps use 64-bit milliseconds since 1970 GMT. Windows timestamps use 64-bit value representing the
     * number of 100-nanosecond intervals since January 1, 1601, with ten thousand times as much precision. <br>
     * DIFF_IN_MILLIS is the difference between January 1 1601 and January 1 1970 in milliseconds. This magic number
     * came from com.mindprod.common11.TestDate. Done according to Gregorian Calendar, no correction for 1752-09-02
     * Wednesday was followed immediately by 1752-09-14 Thursday dropping 12 days. Also according to
     * http://gcc.gnu.org/ml/java-patches/2003-q1/msg00565.html
     */
    private static final long DIFF_IN_MILLIS = 11644473600000L;

    public static long getJavaTime(Long windowsTime)
    {
        return (windowsTime / 10000) - DIFF_IN_MILLIS;
    }

    public static long getWindowsTime(Long javaTime)
    {
        return (javaTime + DIFF_IN_MILLIS) * 10000;
    }
}
