/*
 * Copyright(c) Nimrod97, 2010.
 *
 * Email: Nimrod97@gmail.com
 * Project: http://code.google.com/p/xmltv2jtv/
 */

import convert.Tv2JTvConverter;
import jtv.JFileChannel;
import jtv.vo.JChannel;
import jtv.vo.JProgramme;
import options.OptionsEx;
import org.apache.commons.cli.*;
import org.apache.log4j.Logger;
import xmltv.generated.Tv;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.Unmarshaller;
import java.io.File;
import java.util.Date;
import java.util.List;


public class Main
{
    private static final Logger LOGGER = Logger.getLogger(Main.class.getName());


    public static void main(String[] args) throws Exception
    {
        //setup log4j
        String userDir = (String)System.getProperties().get("user.dir");
        System.getProperties().put("log4j.configuration", new File(userDir, "log4j.properties").getPath());

        Long start = System.currentTimeMillis();
        OptionsEx options = createOptions();
        CommandLine cl = parseArgs(args, options);
        options.getOptionEx("c").getDefaultValue();


        //System.exit(0);
        try
        {

            JAXBContext jc = JAXBContext.newInstance("xmltv.generated");

            // unmarshal from foo.xml
            Unmarshaller u = jc.createUnmarshaller();
            Tv tv = (Tv) u.unmarshal(new File("J:\\Projects\\fsch\\xmltv\\dtd\\program_xml.xml"));


            Tv2JTvConverter xmltv2JTV = new Tv2JTvConverter(tv);
            List<JChannel> channels = xmltv2JTV.convert();
            // marshal to System.out
            //Marshaller m = jc.createMarshaller();
            //m.marshal(tv, System.out);

            //System.exit(0);

            for (JChannel jChannel : channels)
            {
                for (JProgramme jProgramme : jChannel.getProgrammes())
                {
                    jProgramme.setStart(new Date((jProgramme.getStart().getTime() - 1000 * 60 * 60)));
                }

                JFileChannel jFileChannel = new JFileChannel(new File("J:\\Projects\\fsch\\xmltv\\jtv\\program_jtv1"), jChannel, "Cp1251");
                jFileChannel.write();

                JFileChannel jFileChannel2 = new JFileChannel(new File("J:\\Projects\\fsch\\xmltv\\jtv\\program_jtv1"), jChannel, "Cp1251");
                jFileChannel2.read();
            }
        }
        catch (Exception e)
        {
            LOGGER.error("Convertion failed", e);
        }

        if (cl.hasOption("h"))
        {
            HelpFormatter f = new HelpFormatter();
            f.printHelp("XMLTV2JTV", options);
        }

        LOGGER.info("Time elapsed: " + (System.currentTimeMillis() - start) + " ms");
    }

    private static CommandLine parseArgs(String[] args, Options options) throws ParseException
    {
        BasicParser parser = new BasicParser();
        return parser.parse(options, args);
    }

    private static OptionsEx createOptions()
    {
        OptionsEx opt = new OptionsEx();
        opt.addOption("h", false, "Print help for this application");
        opt.addOption("c", true, "The charsetName to use for JTV format encoding/decoding", "Cp1251");
        opt.addOption("dsn", true, "The data source to use");
        return opt;
    }
}
