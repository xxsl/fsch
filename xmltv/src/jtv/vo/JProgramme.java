package jtv.vo;

import java.util.Date;

/**
 * User: aturtsevitch
 * Date: May 19, 2010
 * Time: 12:04:03 PM
 */
public class JProgramme
{
    private String name;
    private Date start;

    public JProgramme(String s, Date date)
    {
        name = s;
        start = date;
    }

    public Date getStart()
    {
        return start;
    }

    public void setStart(Date start)
    {
        this.start = start;
    }

    public String getName()
    {
        return name;
    }

    public void setName(String name)
    {
        this.name = name;
    }
}
