using System.Xml.Serialization;

namespace STXGen2
{
    [XmlRoot(ElementName = "Cell")]
    public class TempDataTableCell
    {
        [XmlElement(ElementName = "ColumnUid")]
        public string ColumnUid { get; set; }

        [XmlElement(ElementName = "Value")]
        public string Value { get; set; }
    }
}