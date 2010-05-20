package jtv.ndx;

public class NDXTime
{
    private Short offset;
    private Long time;

    public NDXTime(Short offset, Long time)
    {
        this.offset = offset;
        this.time = time;
    }

    public Short getOffset()
    {
        return offset;
    }

    public Long getTime()
    {
        return time;
    }
}
