using System.Collections.Generic;
using System.Xml.Serialization;

namespace STXGen2
{
    [XmlRoot(ElementName = "Rows")]
    public class TempDataTableRows
    {
        [XmlElement(ElementName = "Row")]
        public List<TempDataTableRow> RowList { get; set; }
    }
}