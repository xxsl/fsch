package jtv.vo;

import java.util.List;

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

    public List<JProgramme> getProgrammes()
    {
        return programmes;
    }
}
