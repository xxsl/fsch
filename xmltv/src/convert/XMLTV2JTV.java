package convert;

import jtv.vo.JChannel;
import jtv.vo.JProgramme;
import org.joda.time.DateTime;
import org.joda.time.format.*;
import xmltv.generated.Channel;
import xmltv.generated.Programme;
import xmltv.generated.Tv;

import javax.xml.bind.DatatypeConverter;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class XMLTV2JTV
{
    private Tv xmltv;

    public XMLTV2JTV(Tv xmltv)
    {
        this.xmltv = xmltv;
    }

    public List<JChannel> convert() throws ParseException
    {
        List<JChannel> jChannels = new ArrayList<JChannel>();

        Map<String, List<Programme>> programmeMap = createProgrammMap();
        Map<String, Channel> channelMap = createChannelMap();

        for (String key : programmeMap.keySet())
        {
            List<Programme> programmes = programmeMap.get(key);
            JChannel jChannel = new JChannel(channelMap.get(key).getDisplayName().get(0).getvalue(), new ArrayList<JProgramme>());//todo lang
            for (Programme programme : programmes)
            {
                jChannel.getProgrammes().add(new JProgramme(getTitle(programme), getDate(programme)));
            }
            jChannels.add(jChannel);
        }

        return jChannels;
    }

    //ISO 8601
    private Date getDate(Programme programme) throws ParseException
    {
        SimpleDateFormat timeParser =  new SimpleDateFormat();
        timeParser.applyPattern("yyyyMMddHHmmss");
        DateTimeFormatter parser2 = ISODateTimeFormat.dateTimeNoMillis();

        return DatatypeConverter.parseDateTime(programme.getStart()).getTime();
    }

    private String getTitle(Programme programme)
    {
        //todo lang
        return programme.getTitle().get(0).getvalue();
    }

    private Map<String, List<Programme>> createProgrammMap()
    {
        Map<String, List<Programme>> programmeMap = new HashMap<String, List<Programme>>();
        for (Programme proramme : xmltv.getProgramme())
        {
            List<Programme> programmes = programmeMap.get(proramme.getChannel());
            if (programmes == null)
            {
                programmes = new ArrayList<Programme>();
                programmeMap.put(proramme.getChannel(), programmes);
            }
            programmes.add(proramme);
        }
        return programmeMap;
    }

    private Map<String, Channel> createChannelMap()
    {
        Map<String, Channel> channelMap = new HashMap<String, Channel>();
        for (Channel channel : xmltv.getChannel())
        {
            channelMap.put(channel.getId(), channel);
        }
        return channelMap;
    }
}
