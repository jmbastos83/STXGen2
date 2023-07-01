using System.Collections.Generic;
using System.Xml.Serialization;

namespace STXGen2
{
    [XmlRoot(ElementName = "Cells")]
    public class TempDataTableCells
    {
        [XmlElement(ElementName = "Cell")]
        public List<TempDataTableCell> CellList { get; set; }
    }
}