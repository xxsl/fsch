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
import java.util.ArrayList;
import java.util.List;


/**
 * 
 */
@XmlAccessorType(AccessType.FIELD)
@XmlType(name = "", propOrder = {
    "value",
    "icon"
})
@XmlRootElement(name = "star-rating")
@Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
public class StarRating {

    @XmlAttribute
    @XmlJavaTypeAdapter(NormalizedStringAdapter.class)
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    protected String system;
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    protected String value;
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    protected List<Icon> icon;

    /**
     * Gets the value of the system property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    public String getSystem() {
        return system;
    }

    /**
     * Sets the value of the system property.
     * 
     * @param value
     *     allowed object is
     *     {@link String }
     *     
     */
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    public void setSystem(String value) {
        this.system = value;
    }

    /**
     * Gets the value of the value property.
     * 
     * @return
     *     possible object is
     *     {@link String }
     *     
     */
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    public String getValue() {
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
    public void setValue(String value) {
        this.value = value;
    }

    /**
     * Gets the value of the icon property.
     * 
     * <p>
     * This accessor method returns a reference to the live list,
     * not a snapshot. Therefore any modification you make to the
     * returned list will be present inside the JAXB object.
     * This is why there is not a <CODE>set</CODE> method for the icon property.
     * 
     * <p>
     * For example, to add a new item, do as follows:
     * <pre>
     *    getIcon().add(newItem);
     * </pre>
     * 
     * 
     * <p>
     * Objects of the following type(s) are allowed in the list
     * {@link Icon }
     * 
     * 
     */
    @Generated(value = "com.sun.tools.xjc.Driver", date = "2010-05-18T03:48:27+03:00", comments = "JAXB RI v2.0-b26-ea3")
    public List<Icon> getIcon() {
        if (icon == null) {
            icon = new ArrayList<Icon>();
        }
        return this.icon;
    }

}
