import convert.XMLTV2JTV;
import jtv.JFileChannel;
import jtv.vo.JChannel;
import jtv.vo.JProgramme;
import org.apache.commons.cli.*;
import xmltv.generated.Tv;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.Unmarshaller;
import java.io.File;
import java.util.Date;
import java.util.List;

/**
 * User: aturtsevitch
 * Date: May 18, 2010
 * Time: 3:40:21 PM
 */
public class Main
{
    public static void main(String[] args) throws Exception
    {
        Long start = System.currentTimeMillis();
        Options options = createOptions();
        CommandLine cl = parseArgs(args, options);


        JAXBContext jc = JAXBContext.newInstance("xmltv.generated");

        // unmarshal from foo.xml
        Unmarshaller u = jc.createUnmarshaller();
        Tv tv = (Tv) u.unmarshal(new File("G:\\work\\xmltv\\dtd\\program_xml.xml"));


        XMLTV2JTV xmltv2JTV = new XMLTV2JTV(tv);
        List<JChannel> channels = xmltv2JTV.convert();
        // marshal to System.out
        //Marshaller m = jc.createMarshaller();
        //m.marshal(tv, System.out);

        for(JChannel jChannel:channels)
        {
            for (JProgramme jProgramme : jChannel.getProgrammes())
            {
                jProgramme.setStart(new Date((jProgramme.getStart().getTime() - 1000 * 60 * 60)));
            }

            JFileChannel jFileChannel = new JFileChannel(new File("G:\\work\\xmltv\\jtv\\program_jtv1"), jChannel);
            jFileChannel.write();
        }


        if (cl.hasOption("h"))
        {
            HelpFormatter f = new HelpFormatter();
            f.printHelp("XMLTV2JTV", options);
        }

        System.out.println("Time elapsed: " + (System.currentTimeMillis() - start) + " ms");
    }

    private static CommandLine parseArgs(String[] args, Options options) throws ParseException
    {
        BasicParser parser = new BasicParser();
        return parser.parse(options, args);
    }

    private static Options createOptions()
    {
        Options opt = new Options();
        opt.addOption("h", false, "Print help for this application");
        opt.addOption("u", true, "The username to use");
        opt.addOption("dsn", true, "The data source to use");
        return opt;
    }
}
