using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;

namespace STXGen2
{
    [FormAttribute("STXGen2.ToolFind", "ToolFind.b1f")]
    class ToolFind : UserFormBase
    {

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText0;
        private string selectedColUID;
        private bool sortColumnAsc = false;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private bool closedoc = false;

        public object OrderGrid { get; private set; }

        public delegate void DataSelectedHandler(ToolFindData data);

        // Define an event based on that delegate
        public event DataSelectedHandler DataSelected;

        public class ToolFindData
        {
            public string ToolNumber { get; set; }
            public string SalesOrderNumber { get; set; }
            public string SalesEmployee { get; set; }
            public string CardCode { get; set; }
            public string CardName { get; set; }
            public string LicTradeNum { get; set; }
            public string Brand { get; set; }
            public string OEM { get; set; }
            public string KAM { get; set; }

        }

        public ToolFind()
        {
            //searchTimer = new System.Timers.Timer(500); // 300 ms for example
            //searchTimer.Elapsed += OnTimedEvent;
            //searchTimer.AutoReset = false;
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("gdTInfo").Specific));
            this.Grid0.ClickBefore += new SAPbouiCOM._IGridEvents_ClickBeforeEventHandler(this.Grid0_ClickBefore);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.Button0_PressedBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("findFlt").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            this.CloseBefore += new CloseBeforeHandler(this.Form_CloseBefore);

        }



        private void OnCustomInitialize()
        {
            sortColumnAsc = true;
            OrderGrid = 1;
            PopulateGrid(OrderGrid);
            Button0.Item.Enabled = false;
        }

        private void PopulateGrid(object order)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = $"select * ,ROW_NUMBER() over (order by (select null)) as \"Index\" from (\n" +
                            "select distinct T1.\"U_STXToolNum\" as \"Tool Number\",T1.\"U_STXPartName\" as \"Part Name\",coalesce(T2.\"ShipDate\",T3.\"DocDueDate\") as \"Expected Arrival\",T4.\"SlpName\" as \"Employee\",T0.\"DocNum\" as \"Sales Order\",\n" +
                            "T0.\"CardName\" as \"Customer\",T0.\"LicTradNum\",T3.\"DocNum\" as \"Quote Number\",T5.\"U_PartPic\" as \"Picture\",T0.\"CardCode\",T0.\"U_STXBrand\",T0.\"U_STXOEM\",T0.\"U_STXGKAM\"\n" +
                            "from ORDR T0\n" +
                            "inner join RDR1 T1 on T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                            "left join QUT1 T2 on T1.\"BaseEntry\" = T2.\"DocEntry\" and T1.\"BaseLine\" = T2.\"LineNum\" and T1.\"BaseType\" = T2.\"ObjType\"\n" +
                            "left join OQUT T3 on T2.\"DocEntry\" = T3.\"DocEntry\"\n" +
                            "left join OSLP T4 on T0.\"SlpCode\" = T4.\"SlpCode\"\n" +
                            "left join \"@STXQC19\" T5 on T1.\"U_STXQC19ID\" = T5.\"DocEntry\"\n" +
                            "where T1.\"LineStatus\" = 'O' and T0.\"CANCELED\" = 'N'\n" +
                            "union all\n" +
                            "select distinct T1.\"U_STXToolNum\" as \"Tool Number\",T1.\"U_STXPartName\",coalesce(T1.\"ShipDate\",T0.\"DocDueDate\") as \"Expected Arrival\",T4.\"SlpName\",null as \"Sales Order\",\n" +
                            "T0.\"CardName\" as \"Customer\",T0.\"LicTradNum\",T0.\"DocNum\" as \"Quote Number\",T5.\"U_PartPic\" as \"Picture\",T0.\"CardCode\",T0.\"U_STXBrand\",T0.\"U_STXOEM\",T0.\"U_STXGKAM\"\n" +
                            "from OQUT T0\n" +
                            "inner join QUT1 T1 on T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                            "left join OSLP T4 on T0.\"SlpCode\" = T4.\"SlpCode\"\n" +
                            "left join \"@STXQC19\" T5 on T1.\"U_STXQC19ID\" = T5.\"DocEntry\"\n" +
                            "where T1.\"LineStatus\" = 'O' and T0.\"CANCELED\" = 'N') T0 \n" +
                            "order by {0}";

            query = string.Format(query, order);
            Grid0.DataTable.ExecuteQuery(query);

            for (int i = 0; i < Grid0.Columns.Count; i++)
            {
                Grid0.Columns.Item(i).TitleObject.Sortable = true;
                if ((Convert.ToInt32(order) - 1) == i)
                {
                    selectedColUID = Grid0.Columns.Item(i).TitleObject.Caption;
                }
            }
            Grid0.Columns.Item(selectedColUID).TitleObject.Sort(BoGridSortType.gst_Ascending);

            EditText0.Item.AffectsFormMode = false;

            SAPbouiCOM.GridColumn oPicture = Grid0.Columns.Item("Picture");
            oPicture.Visible = false;

            SAPbouiCOM.GridColumn oIndex = Grid0.Columns.Item("Index");
            oIndex.Visible = false;

            SAPbouiCOM.GridColumn oCardCode = Grid0.Columns.Item("CardCode");
            oCardCode.Visible = false;
            SAPbouiCOM.GridColumn oBrand = Grid0.Columns.Item("U_STXBrand");
            oBrand.Visible = false;
            SAPbouiCOM.GridColumn oOEM = Grid0.Columns.Item("U_STXOEM");
            oOEM.Visible = false;
            SAPbouiCOM.GridColumn oKAM = Grid0.Columns.Item("U_STXGKAM");
            oKAM.Visible = false;

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



        private void FilterGrid(string pVal)
        {
            try
            {
                this.UIAPIRawForm.Freeze(true);

                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("findFlt").Specific;
                string filterValue = oEditText.Value.Trim().ToLower();

                int dataTableIndex = -1;
                int gridVisualIndex = -1;

                // Find the index in the DataTable
                for (int i = 0; i < Grid0.DataTable.Rows.Count; i++)
                {
                    object cellObject = Grid0.DataTable.GetValue(selectedColUID, i);
                    string value;

                    if (cellObject != null)
                    {
                        // Check if the object is a date and format it appropriately.
                        if (cellObject is DateTime dateTimeValue)
                        {
                            Program.SBO_Application.SetStatusBarMessage("Search not possible.", BoMessageTime.bmt_Short, false);
                            value = "";
                        }
                        else
                        {
                            // For non-date values, just convert to string as usual
                            value = cellObject.ToString().ToLower();
                        }
                    }
                    else
                    {
                        // Handle null or empty cells
                        value = string.Empty;
                    }

                    if (value.Equals(filterValue) || value.Contains(filterValue))
                    {
                        dataTableIndex = i;
                        break;
                    }
                }
                // Find the visual index in the Grid that matches the DataTable index
                if (dataTableIndex != -1)
                {
                    for (int j = 0; j < Grid0.Rows.Count; j++)
                    {
                        if (Grid0.GetDataTableRowIndex(j) == dataTableIndex)
                        {
                            gridVisualIndex = j;
                            break;
                        }
                    }
                }

                // 'GridVisualIndex' holds the index of the row in the grid that matches the value.
                if (gridVisualIndex != -1)
                {
                    Grid0.Rows.SelectedRows.Add(gridVisualIndex);
                    Button0.Item.Enabled = true;
                }
            }
            finally
            {
                Grid0.AutoResizeColumns();
                this.UIAPIRawForm.Freeze(false);
            }

        }

        private void Grid0_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = false;
            if (pVal.ColUID != "RowsHeader" && pVal.Row == -1)
            {
                if (selectedColUID == pVal.ColUID)
                {
                    if (sortColumnAsc)
                    {
                        Grid0.Columns.Item(selectedColUID).TitleObject.Sort(BoGridSortType.gst_Descending);
                        sortColumnAsc = false;
                    }
                    else
                    {
                        Grid0.Columns.Item(selectedColUID).TitleObject.Sort(BoGridSortType.gst_Ascending);
                        sortColumnAsc = true;
                    }

                }
                else
                {
                    selectedColUID = pVal.ColUID;
                    Grid0.Columns.Item(selectedColUID).TitleObject.Sort(BoGridSortType.gst_Ascending);
                    sortColumnAsc = true;
                }
            }
            else
            {
                Grid0.Rows.SelectedRows.Add(pVal.Row);
                Button0.Item.Enabled = true;
            }
        }

        

        private void Button2_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            Button2.Item.AffectsFormMode = false;
            Grid0.Rows.SelectedRows.Clear();
            FilterGrid(EditText0.Value.ToString());


        }

        private void Form_CloseBefore(SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (closedoc)
            {
                closedoc = false;
                if (Grid0.Rows.SelectedRows.Count > 0)
                {
                    int selectedIndex = Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                    string toolNumber = Grid0.DataTable.GetValue("Tool Number", selectedIndex).ToString();
                    string salesOrderNumber = Grid0.DataTable.GetValue("Sales Order", selectedIndex).ToString();
                    string SalesEmployee = Grid0.DataTable.GetValue("Employee", selectedIndex).ToString();

                    string selectedCardCode = Grid0.DataTable.GetValue("CardCode", selectedIndex).ToString();
                    string selectedCardName = Grid0.DataTable.GetValue("Customer", selectedIndex).ToString();
                    string selectedLicTradeNum = Grid0.DataTable.GetValue("LicTradNum", selectedIndex).ToString();
                    string selectedBrand = Grid0.DataTable.GetValue("U_STXBrand", selectedIndex).ToString();
                    string selectedOEM = Grid0.DataTable.GetValue("U_STXOEM", selectedIndex).ToString();
                    string selectedKam = Grid0.DataTable.GetValue("U_STXGKAM", selectedIndex).ToString();

                    ToolFindData data = new ToolFindData()
                    {
                        ToolNumber = toolNumber,
                        SalesOrderNumber = salesOrderNumber,
                        SalesEmployee = SalesEmployee,

                        CardCode = selectedCardCode,
                        CardName = selectedCardName,
                        LicTradeNum = selectedLicTradeNum,
                        Brand = selectedBrand,
                        OEM = selectedOEM,
                        KAM = selectedKam

                    };

                    // Raise the event with the ToolFindData instance
                    DataSelected?.Invoke(data);
                }
            
            }

        }

        private void Button0_PressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            closedoc = true;


        }
    }
}
