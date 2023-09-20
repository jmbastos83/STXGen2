using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace STXGen2
{
    [FormAttribute("STXGen2.RelationshipMap", "RelationshipMap.b1f")]
    class RelationshipMap : UserFormBase
    {
        public RelationshipMap()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_0").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Grid Grid0;

        private void OnCustomInitialize()
        {
            //SAPbouiCOM.DataTable dt = Grid0.DataTable ?? this.UIAPIRawForm.DataSources.DataTables.Add("DT_0");

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
           
            string query = @"
                            SELECT 
                                O.DocEntry as 'Order ID',
                                O.CardName as 'Customer',
                                L.LineNum as 'Line Number',
                                L.ItemCode as 'Product',
                                L.Quantity as 'Quantity'
                            FROM 
                                ORDR O
                            JOIN 
                                RDR1 L on O.DocEntry = L.DocEntry
                            ORDER BY 
                                O.DocEntry, L.LineNum";

            Grid0.DataTable.ExecuteQuery(query);
            Grid0.Item.Enabled = false;

            Grid0.CollapseLevel = 2;
            //oRecordSet.DoQuery(query);
            //string xml = oRecordSet.GetAsXML();
            

            ////dt.Rows.Clear();
            //Grid0.DataTable.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_All, xml);

            //// Set Collapse Level for Grid Grouping
            //Grid0.CollapseLevel = 2;
        }

    }
}
