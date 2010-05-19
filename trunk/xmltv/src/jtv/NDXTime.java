package jtv;

/**
 * User: aturtsevitch
 * Date: May 19, 2010
 * Time: 12:53:50 PM
 */
public class NDXTime
{
    private Long offset;
    private Long time;

    public NDXTime(Long offset, Long time)
    {
        this.offset = offset;
        this.time = time;
    }

    public Long getOffset()
    {
        return offset;
    }

    public Long getTime()
    {
        return time;
    }
}
