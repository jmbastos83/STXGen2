using System.Collections.Generic;
using System.Xml.Serialization;

namespace STXGen2
{
    [XmlRoot(ElementName = "Columns")]
    public class TempDataTableColumns
    {
        [XmlElement(ElementName = "Column")]
        public List<TempDataTableColumn> ColumnList { get; set; }
    }
}