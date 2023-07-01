using System.Collections.Generic;
using System.Xml.Serialization;
using System.Runtime.Serialization;
using System;

namespace STXGen2
{
    [XmlRoot(ElementName = "DataTable")]
    public class TempDataTable
    {
        private Dictionary<string, int> columnIndices = new Dictionary<string, int>();
        private Dictionary<string, int> columnMaxValues = new Dictionary<string, int>();

        [XmlElement(ElementName = "Columns")]
        public TempDataTableColumns Columns { get; set; }

        [XmlElement(ElementName = "Rows")]
        public TempDataTableRows Rows { get; set; }

        [XmlAttribute(AttributeName = "Uid")]
        public string Uid { get; set; }

        public TempDataTable()
        {
        }

        [OnDeserialized]
        internal void OnDeserializedMethod(StreamingContext context)
        {
            for (int i = 0; i < Columns.ColumnList.Count; i++)
            {
                columnIndices[Columns.ColumnList[i].Uid] = i;
            }

            foreach (TempDataTableRow row in Rows.RowList)
            {
                foreach (TempDataTableCell cell in row.Cells.CellList)
                {
                    if (!columnMaxValues.ContainsKey(cell.ColumnUid) || int.Parse(cell.Value) > columnMaxValues[cell.ColumnUid])
                    {
                        columnMaxValues[cell.ColumnUid] = int.Parse(cell.Value);
                    }
                }
            }
        }

        public int colIndex(string column)
        {
            if (columnIndices.TryGetValue(column, out int index))
            {
                return index;
            }
            throw new Exception("Column not found " + column);
        }

        public int GetMaxValue(string column)
        {
            if (columnMaxValues.TryGetValue(column, out int maxVal))
            {
                return maxVal;
            }
            throw new Exception("Column not found " + column);
        }
    }
}
