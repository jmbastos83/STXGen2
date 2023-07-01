using System.Xml.Serialization;

using System.Xml.Serialization;

namespace STXGen2
{
    [XmlRoot(ElementName = "Column")]
    public class TempDataTableColumn
    {
        [XmlAttribute(AttributeName = "Uid")]
        public string Uid { get; set; }

        [XmlAttribute(AttributeName = "Type")]
        public string Type { get; set; }

        [XmlAttribute(AttributeName = "MaxLength")]
        public string MaxLength { get; set; }
    }
}
