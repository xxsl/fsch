import xmltv.generated.Tv;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.Marshaller;
import javax.xml.bind.Unmarshaller;
import java.io.File;

/**
 * User: aturtsevitch
 * Date: May 18, 2010
 * Time: 3:40:21 PM
 */
public class XMLTV2JTV
{
    public static void main(String[] args) throws Exception
    {
        JAXBContext jc = JAXBContext.newInstance("xmltv.generated");

        // unmarshal from foo.xml
        Unmarshaller u = jc.createUnmarshaller();
        Tv tv = (Tv) u.unmarshal(new File("j:\\Projects\\XMLTV\\dtd\\program.xml"));

        // marshal to System.out
        //Marshaller m = jc.createMarshaller();
        //m.marshal(tv, System.out);
    }
}
