using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace STXGen2
{
    class XMLDatasource
    {
        public class DbDataSources
        {
            public string Uid { get; set; }
            public List<Row> Rows { get; set; }
        }

        public class Row
        {
            public List<Cell> Cells { get; set; }
        }

        public class Cell
        {
            public string Uid { get; set; }
            public string Value { get; set; }
        }

        public static DbDataSources ParseXml(XDocument xml)
        {
            var dbDataSourcesElement = xml.Descendants("dbDataSources").FirstOrDefault();
            var rowsElements = dbDataSourcesElement?.Descendants("row") ?? Enumerable.Empty<XElement>();

            var dbDataSources = new DbDataSources
            {
                Uid = dbDataSourcesElement?.Attribute("uid")?.Value,
                Rows = rowsElements.Select(rowElement =>
                    new Row
                    {
                        Cells = rowElement.Descendants("cell").Select(cellElement =>
                            new Cell
                            {
                                Uid = cellElement.Element("uid")?.Value,
                                Value = cellElement.Element("value")?.Value
                            }
                        ).ToList()
                    }
                ).ToList()
            };

            return dbDataSources;
        }

        public static DbDataSources GetDbDataSourcesFromOperation(SAPbouiCOM.DataTable operations, Dictionary<string, string> columnToUidMappings)
        {
            var dbDataSources = new DbDataSources
            {
                Uid = "@STXQC19O",
                Rows = new List<Row>()
            };

            // loop over the rows of the DataTable
            for (int rowIndex = 0; rowIndex < operations.Rows.Count; rowIndex++)
            {
                var row = new Row
                {
                    Cells = new List<Cell>()
                };

                // loop over the columns of the DataTable
                for (int columnIndex = 0; columnIndex < operations.Columns.Count; columnIndex++)
                {
                    var columnName = operations.Columns.Item(columnIndex).Name;

                    // use the column name to get the corresponding Uid from the dictionary
                    var uid = columnToUidMappings.ContainsKey(columnName) ? columnToUidMappings[columnName] : null;

                    if (uid != null)
                    {
                        var cell = new Cell
                        {
                            Uid = uid,
                            Value = operations.GetValue(columnIndex, rowIndex).ToString()
                        };

                        row.Cells.Add(cell);
                    }
                }

                dbDataSources.Rows.Add(row);
            }

            return dbDataSources;
        }

        public static XDocument GenerateXml(DbDataSources dbDataSources)
        {
            var dbDataSourcesElement = new XElement("dbDataSources", new XAttribute("uid", dbDataSources.Uid));
            var rowsElement = new XElement("rows");
            dbDataSourcesElement.Add(rowsElement);

            foreach (var row in dbDataSources.Rows)
            {
                var rowElement = new XElement("row");
                var cellsElement = new XElement("cells");
                rowElement.Add(cellsElement);

                foreach (var cell in row.Cells)
                {
                    var cellElement = new XElement("cell",
                        new XElement("uid", cell.Uid),
                        new XElement("value", cell.Value)
                    );

                    cellsElement.Add(cellElement);
                }

                rowsElement.Add(rowElement);
            }

            return new XDocument(new XDeclaration("1.0", "utf-16", "yes"), dbDataSourcesElement);
        }
    }
}
