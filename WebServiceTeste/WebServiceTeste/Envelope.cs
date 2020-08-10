using System;
using System.Collections.Generic;
using System.Web;
using System.Xml.Serialization;

namespace WebServiceTeste
{
    //[XmlType(TypeName = "EnvelopeXXXX")]
    public class Envelope
    {
        //[XmlArrayItem("MemberName")]
        [XmlElement(ElementName = "EnvelopeID")]
        public string EnvelopeID { get; set; }
        [XmlElement(ElementName = "EnvelopeValor")]
        public string EnvelopeValor { get; set; }
    }
}