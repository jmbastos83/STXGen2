using System.Xml.Serialization;

namespace STXGen2
{
    [XmlRoot(ElementName = "Row")]
    public class TempDataTableRow
    {
        [XmlElement(ElementName = "Cells")]
        public TempDataTableCells Cells { get; set; }
    }
}