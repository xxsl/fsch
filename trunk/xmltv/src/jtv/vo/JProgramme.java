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

    public String getName()
    {
        return name;
    }

    public Date getStart()
    {
        return start;
    }
}
