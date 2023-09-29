using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using SAPbouiCOM.Framework;

namespace STXGen2
{
    [FormAttribute("RelationshipMap", "RelationshipMap.b1f")]
    class RelationshipMap : UserFormBase
    {
        private SAPbouiCOM.Grid Grid0;
        private Button Button0;
        private Button Button1;

        public static string relDocEntry { get; set; }


        public RelationshipMap()
        {

        }



        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Gresult").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }



        private void OnCustomInitialize()
        {
            //SAPbouiCOM.DataTable dt = Grid0.DataTable ?? this.UIAPIRawForm.DataSources.DataTables.Add("DT_0");

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = $"With SOInfo as(\n" +

                            "Select 20 as \"FlowOrder\",'Sales Order' as \"DocType\",T1.\"U_STXToolNum\",T1.\"U_STXPartNum\",T0.\"CardCode\",T0.\"CardName\",T1.\"ItemCode\",T0.\"DocNum\",T0.\"DocDate\",\n" +
                            "T0.\"DocDueDate\",T1.\"VisOrder\",T2.\"U_NAME\" as \"Updated By\",T1.\"DocEntry\",T0.\"ObjType\",T1.\"LineNum\",T1.\"BaseEntry\",T1.\"BaseLine\", T1.\"BaseType\"\n" +
                            "from ORDR T0\n" +
                            "inner join RDR1 T1 on T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                            "inner join OUSR T2 on coalesce(T0.\"UserSign2\",T0.\"UserSign\") = T2.\"USERID\"\n" +
                            "where T0.\"DocEntry\" = {0}),\n" +

                            "QUOTEInfo as (\n" +

                            "Select 1 as \"FlowOrder\",'Sales Quotation' as \"DocType\",T1.\"U_STXToolNum\",T1.\"U_STXPartNum\",T0.\"CardCode\",T0.\"CardName\",T1.\"ItemCode\",T0.\"DocNum\",T0.\"DocDate\",\n" +
                            "T0.\"DocDueDate\",T1.\"VisOrder\",T2.\"U_NAME\" as \"Updated By\",T0.\"ObjType\"\n" +
                            "from OQUT T0\n" +
                            "inner join QUT1 T1 on T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                            "inner join OUSR T2 on coalesce(T0.\"UserSign2\",T0.\"UserSign\") = T2.\"USERID\"\n" +
                            "inner join SOInfo T3 on T1.\"DocEntry\" = T3.\"BaseEntry\" and T1.\"LineNum\" = T3.\"BaseLine\" and T1.\"ObjType\" = T3.\"BaseType\"\n" +
                            "where T3.\"DocEntry\" = {0}),\n" +

                            "DELIVERYInfo as (\n" +

                            "Select 30 as \"FlowOrder\",'Delivery Note' as \"DocType\",T1.\"U_STXToolNum\",T1.\"U_STXPartNum\",T0.\"CardCode\",T0.\"CardName\",T1.\"ItemCode\",T0.\"DocNum\",T0.\"DocDate\",\n" +
                            "T0.\"DocDueDate\",T1.\"VisOrder\",T2.\"U_NAME\" as \"Updated By\",T0.\"ObjType\"\n" +
                            "from ODLN T0\n" +
                            "inner join DLN1 T1 on T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                            "inner join OUSR T2 on coalesce(T0.\"UserSign2\",T0.\"UserSign\") = T2.\"USERID\"\n" +
                            "inner join SOInfo T3 on T1.\"BaseEntry\" = T3.\"DocEntry\" and T1.\"BaseLine\" = T3.\"LineNum\" and T1.\"BaseType\" = T3.\"ObjType\"\n" +
                            "where T3.\"DocEntry\" = {0}),\n" +

                            "WOInfo as (\n" +

                            "select 25 as \"FlowOrder\",'Production Order' as \"DocType\",T3.U_ToolNum,T3.\"U_PartNum\",T2.\"CardCode\",T2.\"U_STXCustName\" AS \"CardName\",T2.\"ItemCode\", T2.\"DocNum\",\n" +
                            "T2.\"PostDate\" AS \"DocDate\", T2.\"DueDate\" AS \"DocDueDate\",T2.\"U_STXSOLineNum\" as \"VisOrder\",T4.\"U_NAME\" as \"Updated By\",T2.\"ObjType\"\n" +
                            "from ORDR T0\n" +
                            "inner join RDR1 T1 on T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                            "inner join OWOR T2 on T0.\"DocNum\" = T2.\"U_STXSONum\" and T1.\"LineNum\" = T2.\"U_STXSOLineNum\"\n" +
                            "inner join \"@STXQC19\" T3 on T2.\"U_STXQC19ID\" = T3.\"DocEntry\"\n" +
                            "inner join OUSR T4 on coalesce(T2.\"UserSign2\",T2.\"UserSign\") = T4.\"USERID\"\n" +
                            "where T0.\"DocEntry\" = {0}),\n" +

                            "INVOICEInfo as (\n" +

                            "Select 40 as \"FlowOrder\",'A/R Invoice' as \"DocType\",T1.\"U_STXToolNum\",T1.\"U_STXPartNum\",T0.\"CardCode\",T0.\"CardName\",T1.\"ItemCode\",T0.\"DocNum\",T0.\"DocDate\",\n" +
                            "T0.\"DocDueDate\",T1.\"VisOrder\",T2.\"U_NAME\" as \"Updated By\",T0.\"ObjType\"\n" +
                            "from OINV T0\n" +
                            "inner join INV1 T1 on T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                            "inner join OUSR T2 on coalesce(T0.\"UserSign2\",T0.\"UserSign\") = T2.\"USERID\"\n" +
                            "left join DLN1 T3 on T1.\"BaseEntry\" = T3.\"DocEntry\" and T1.\"BaseLine\" = T3.\"LineNum\" and T1.\"BaseType\" = T3.\"ObjType\"\n" +
                            "left join SOInfo T4 on T3.\"BaseEntry\" = T4.\"DocEntry\" and T3.\"BaseLine\" = T4.\"LineNum\" and T3.\"BaseType\" = T4.\"ObjType\"\n" +
                            "left join SOInfo T5 on T1.\"BaseEntry\" = T5.\"DocEntry\" and T1.\"BaseLine\" = T5.\"LineNum\" and T1.\"BaseType\" = T5.\"ObjType\"\n" +
                            "where coalesce(T4.\"DocEntry\",T5.\"DocEntry\") = {0})\n" +

                            "select coalesce(T0.\"U_STXToolNum\",'') as \"Tool Num.\",T0.\"U_STXPartNum\" as \"Part Num\",T0.\"FlowOrder\",T0.\"DocType\" as \"Doc. Type\",T0.\"DocNum\" as \"Doc. Number\",T0.\"VisOrder\" as \"Doc. Line\",T0.\"CardCode\",T0.\"CardName\",T0.\"ItemCode\",T0.\"DocDate\",\n" +
                            "T0.\"DocDueDate\",T0.\"Updated By\",T0.\"ObjType\" from QUOTEInfo T0\n" +
                            "union all\n" +
                            "select coalesce(T0.\"U_STXToolNum\",'') as \"Tool Num.\",T0.\"U_STXPartNum\" as \"Part Num\",T0.\"FlowOrder\",T0.\"DocType\" as \"Doc. Type\",T0.\"DocNum\" as \"Doc. Number\",T0.\"VisOrder\" as \"Doc. Line\",T0.\"CardCode\",T0.\"CardName\",T0.\"ItemCode\",T0.\"DocDate\",\n" +
                            "T0.\"DocDueDate\",T0.\"Updated By\",T0.\"ObjType\" from SOInfo T0\n" +
                            "union all\n" +
                            "select coalesce(T0.\"U_STXToolNum\",'') as \"Tool Num.\",T0.\"U_STXPartNum\" as \"Part Num\",T0.\"FlowOrder\",T0.\"DocType\" as \"Doc. Type\",T0.\"DocNum\" as \"Doc. Number\",T0.\"VisOrder\" as \"Doc. Line\",T0.\"CardCode\",T0.\"CardName\",T0.\"ItemCode\",T0.\"DocDate\",\n" +
                            "T0.\"DocDueDate\",T0.\"Updated By\",T0.\"ObjType\" from DELIVERYInfo T0\n" +
                            "union all\n" +
                            "select coalesce(T0.\"U_ToolNum\",'') as \"Tool Num.\",T0.\"U_PartNum\" as \"Part Num\",T0.\"FlowOrder\",T0.\"DocType\" as \"Doc. Type\",T0.\"DocNum\" as \"Doc. Number\",T0.\"VisOrder\" as \"Doc. Line\",T0.\"CardCode\",T0.\"CardName\",T0.\"ItemCode\",T0.\"DocDate\",\n" +
                            "T0.\"DocDueDate\",T0.\"Updated By\",T0.\"ObjType\" from WOInfo T0\n" +
                            "union all\n" +
                            "select coalesce(T0.\"U_STXToolNum\",'') as \"Tool Num.\",T0.\"U_STXPartNum\" as \"Part Num\",T0.\"FlowOrder\",T0.\"DocType\" as \"Doc. Type\",T0.\"DocNum\" as \"Doc. Number\",T0.\"VisOrder\" as \"Doc. Line\",T0.\"CardCode\",T0.\"CardName\",T0.\"ItemCode\",T0.\"DocDate\",\n" +
                            "T0.\"DocDueDate\",T0.\"Updated By\",T0.\"ObjType\" from INVOICEInfo T0\n" +

                            "order by coalesce(T0.\"U_STXToolNum\",''),T0.\"FlowOrder\",T0.\"U_STXPartNum\",T0.\"DocNum\",T0.\"VisOrder\",T0.\"CardCode\",T0.\"CardName\"";

            query = string.Format(query, relDocEntry);
            Grid0.DataTable.ExecuteQuery(query);

            // Setting up the columns
            for (int i = 0; i < Grid0.Columns.Count; i++)
            {
                SAPbouiCOM.EditTextColumn column = (SAPbouiCOM.EditTextColumn)Grid0.Columns.Item(i);

                // Example: if you want to hide a specific column
                if (column.UniqueID == "FlowOrder")
                {
                    column.Visible = false;
                }
                if (column.UniqueID == "ObjType")
                {
                    column.Visible = false;
                }
                if (Grid0.Columns.Item("Doc. Number") is SAPbouiCOM.EditTextColumn docNumColumn)
                {
                    docNumColumn.LinkedObjectType = "17";
                }
            }
            Grid0.Item.Enabled = false;
            Grid0.CollapseLevel = 1;
            Grid0.AutoResizeColumns();
        }



        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            Grid0.Rows.CollapseAll();
        }

        private void Button1_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            Grid0.Rows.ExpandAll();
        }

        private Button Button2;
    }
}
