package jtv.vo;

import java.util.List;

/**
 * User: aturtsevitch
 * Date: May 19, 2010
 * Time: 12:03:37 PM
 */
public class JChannel
{
    private String name;
    private List<JProgramme> programmes;
    
    public JChannel(String s, List<JProgramme> list)
    {
        name = s;
        programmes = list;
    }

    public String getName()
    {
        return name;
    }

    public List getProgrammes()
    {
        return programmes;
    }
}
