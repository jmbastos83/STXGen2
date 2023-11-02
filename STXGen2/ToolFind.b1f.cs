using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using SAPbouiCOM.Framework;

namespace STXGen2
{
    [FormAttribute("STXGen2.ToolFind", "ToolFind.b1f")]
    class ToolFind : UserFormBase
    {

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText0;
        private string selectedColUID;
        private bool sortColumn = false;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;


        public ToolFind()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("gdTInfo").Specific));
            this.Grid0.PressedAfter += new SAPbouiCOM._IGridEvents_PressedAfterEventHandler(this.Grid0_PressedAfter);
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("findFlt").Specific));
            this.EditText0.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText0_KeyDownAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new ResizeAfterHandler(this.Form_ResizeAfter);

        }



        private void OnCustomInitialize()
        {

            PopulateGrid();
            Button0.Item.Enabled = false;
        }

        private void PopulateGrid()
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = $"select distinct T1.\"U_STXToolNum\" as \"Tool Number\",T1.\"U_STXPartName\" as \"Part Name\",coalesce(T2.\"ShipDate\",T3.\"DocDueDate\") as \"Expected Arrival\",T4.\"SlpName\" as \"Employee\",T0.\"DocNum\" as \"Sales Order\",\n" +
                            "T0.\"CardName\" as \"Customer\",T3.\"DocNum\" as \"Quote Number\",T5.\"U_PartPic\" as \"Picture\"\n" +
                            "from ORDR T0\n" +
                            "inner join RDR1 T1 on T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                            "left join QUT1 T2 on T1.\"BaseEntry\" = T2.\"DocEntry\" and T1.\"BaseLine\" = T2.\"LineNum\" and T1.\"BaseType\" = T2.\"ObjType\"\n" +
                            "left join OQUT T3 on T2.\"DocEntry\" = T3.\"DocEntry\"\n" +
                            "left join OSLP T4 on T0.\"SlpCode\" = T4.\"SlpCode\"\n" +
                            "left join \"@STXQC19\" T5 on T1.\"U_STXQC19ID\" = T5.\"DocEntry\"\n" +
                            "where T1.\"LineStatus\" = 'O' and T0.\"CANCELED\" = 'N'\n" +
                            "union all\n" +
                            "select distinct T1.\"U_STXToolNum\" as \"Tool Number\",T1.\"U_STXPartName\",coalesce(T1.\"ShipDate\",T0.\"DocDueDate\") as \"Expected Arrival\",T4.\"SlpName\",null as \"Sales Order\",\n" +
                            "T0.\"CardName\" as \"Customer\",T0.\"DocNum\" as \"Quote Number\",T5.\"U_PartPic\" as \"Picture\"\n" +
                            "from OQUT T0\n" +
                            "inner join QUT1 T1 on T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                            "left join OSLP T4 on T0.\"SlpCode\" = T4.\"SlpCode\"\n" +
                            "left join \"@STXQC19\" T5 on T1.\"U_STXQC19ID\" = T5.\"DocEntry\"\n" +
                            "where T1.\"LineStatus\" = 'O' and T0.\"CANCELED\" = 'N'";

            Grid0.DataTable.ExecuteQuery(query);

            for (int i = 0; i < Grid0.Columns.Count; i++)
            {
                Grid0.Columns.Item(i).TitleObject.Sortable = true;
            }

            SAPbouiCOM.GridColumn oColumn = Grid0.Columns.Item("Picture");
            oColumn.Visible = false;

            SAPbouiCOM.EditTextColumn toolImage = (SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("Tool Number");
            toolImage.LinkedObjectType = "17";

            SAPbouiCOM.EditTextColumn sorderObject = (SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("Sales Order");
            sorderObject.LinkedObjectType = "17";

            SAPbouiCOM.EditTextColumn quoteObject = (SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("Quote Number");
            quoteObject.LinkedObjectType = "23";

            

            Grid0.AutoResizeColumns();
        }



        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            Grid0.AutoResizeColumns();
        }



        private void EditText0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            FilterGrid(pVal);

        }

        private void FilterGrid(SBOItemEventArg pVal)
        {
            SAPbouiCOM.Grid grid = (SAPbouiCOM.Grid)this.UIAPIRawForm.Items.Item("gdTInfo").Specific;
            SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("findFlt").Specific;

            string filterValue = oEditText.Value.Trim().ToLower();

            int colIndex = -1;

            // Iterate through the columns to find the index
            for (int i = 0; i < grid.Columns.Count; i++)
            {
                if (grid.Columns.Item(i).UniqueID == selectedColUID)
                {
                    colIndex = i;
                    break; // Exit the loop once you find the matching column
                }
            }

            if (colIndex != -1)
            {
                for (int i = 0; i < grid.Rows.Count; i++)
                {
                    string cellValue = grid.DataTable.GetValue(colIndex, i).ToString().ToLower();
                    if (cellValue.Contains(filterValue))
                    {
                        // Show the row if the condition is true
                        grid.Rows.SelectedRows.Add(i);
                        break;
                    }
                }
            }
            grid.AutoResizeColumns();
        }

        private void Grid0_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            selectedColUID = pVal.ColUID;
            
        }

        private void Grid0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                this.UIAPIRawForm.Freeze(true);
                if (pVal.Row == -1)
                {
                    
                   // Grid0.Rows.SelectedRows.Add(0);
                }
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
            }

        }

    }
}
