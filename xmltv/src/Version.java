/**
 * Created by IntelliJ IDEA.
 * User: Dookie
 * Date: 21.05.2010
 * Time: 22:53:13
 * To change this template use File | Settings | File Templates.
 */
public class Version
{
    private static long VERSION_ = 0;
    private static long MAJOR_ = 0;
    private static long MINOR_ = 1;

    private static final Version VERSION = new Version(VERSION_, MAJOR_, MINOR_);

    private long version = 0;
    private long major = 0;
    private long minor = 1;
    private String asString;

    private Version(long version, long major, long minor)
    {
        this.version = version;
        this.major = major;
        this.minor = minor;
        this.asString = getVersion() + "." + getMajor() + "." + getMinor();
    }

    public static Version getInfo()
    {
        return VERSION;
    }

    public static String getInfoString()
    {
        return VERSION.toString();
    }

    public long getVersion()
    {
        return version;
    }

    public long getMajor()
    {
        return major;
    }

    public long getMinor()
    {
        return minor;
    }

    @Override
    public String toString()
    {
        return asString;
    }
}
