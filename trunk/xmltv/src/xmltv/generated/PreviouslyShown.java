//
// This file was generated by the JavaTM Architecture for XML Binding(JAXB) Reference Implementation, v2.0-b26-ea3 
// See <a href="http://java.sun.com/xml/jaxb">http://java.sun.com/xml/jaxb</a> 
// Any modifications to this file will be lost upon recompilation of the source schema. 
// Generated on: 2010.05.18 at 03:48:27 PM EEST 
//


package xmltv.generated;

import javax.annotation.Generated;
import javax.xml.bind.annotation.*;
import javax.xml.bind.annotation.adapters.NormalizedStringAdapter;
import javax.xml.bind.annotation.adapters.XmlJavaTypeAdapter;


/**
 * 
 */
@XmlAccessorType(AccessType.FIELD)
@XmlType(name = "")
@XmlRootElement(name = "previously-shown")
@Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
public class PreviouslyShown {

    @XmlAttribute
    @XmlJavaTypeAdapter(NormalizedStringAdapter.class)
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    protected String start;
    @XmlAttribute
    @XmlJavaTypeAdapter(NormalizedStringAdapter.class)
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    protected String channel;

    /**
     * Gets the value of the start property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    public String getStart() {
        return start;
    }

    /**
     * Sets the value of the start property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    public void setStart(String value) {
        this.start = value;
    }

    /**
     * Gets the value of the channel property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    public String getChannel() {
        return channel;
    }

    /**
     * Sets the value of the channel property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    public void setChannel(String value) {
        this.channel = value;
    }

}
