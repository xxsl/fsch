//
// This file was generated by the JavaTM Architecture for XML Binding(JAXB) Reference Implementation, v2.0-b26-ea3 
// See <a href="http://java.sun.com/xml/jaxb">http://java.sun.com/xml/jaxb</a> 
// Any modifications to this file will be lost upon recompilation of the source schema. 
// Generated on: 2010.05.18 at 03:48:27 PM EEST 
//


package xmltv.generated;

import javax.annotation.Generated;
import javax.xml.bind.annotation.*;


/**
 * 
 */
@XmlAccessorType(AccessType.FIELD)
@XmlType(name = "", propOrder = {
    "value"
})
@XmlRootElement(name = "editor")
@Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
public class Editor {

    @XmlValue
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    protected String value;

    /**
     * Gets the value of the value property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    public String getvalue() {
        return value;
    }

    /**
     * Sets the value of the value property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    public void setvalue(String value) {
        this.value = value;
    }

}
