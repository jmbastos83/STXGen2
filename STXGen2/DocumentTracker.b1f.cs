using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using SAPbouiCOM.Framework;

namespace STXGen2
{
    [FormAttribute("STXGen2.DocumentTracker", "DocumentTracker.b1f")]
    class DocumentTracker : UserFormBase
    {

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;

        private SAPbouiCOM.Matrix Matrix0;
        private DataTable oDataTable;
        public static string openDocEntry { get; set; }

        public DocumentTracker()
        {

        }



        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("mtDTrac").Specific));
            this.Matrix0.PressedBefore += new SAPbouiCOM._IMatrixEvents_PressedBeforeEventHandler(this.Matrix0_PressedBefore);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_3").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            this.UnloadBefore += new UnloadBeforeHandler(this.Form_UnloadBefore);

        }



        private void OnCustomInitialize()
        {
            BindDataTableToMatrix("DocTrackInfo", "mtDTrac", openDocEntry);

        }

        private void BindDataTableToMatrix(string tableName, string matrixUID, string docEntry)
        {
            bool tableExists = false;
            //SAPbouiCOM.DataTable oDataTable;

            var dataTables = this.UIAPIRawForm.DataSources.DataTables;
            if (dataTables.Count != 0)
            {
                foreach (SAPbouiCOM.DataTable dt in dataTables)
                {
                    if (dt.UniqueID == tableName)
                    {
                        tableExists = true;
                        break;
                    }
                    if (!tableExists)
                    {
                        oDataTable = this.UIAPIRawForm.DataSources.DataTables.Add(tableName);
                    }
                }
            }
            else
            {
                oDataTable = this.UIAPIRawForm.DataSources.DataTables.Add(tableName);
            }

            DBCalls.DocumentTrackerInfo(oDataTable, docEntry);
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item(matrixUID).Specific;

            oMatrix.Clear();
            for (int i = 0; i < oMatrix.Columns.Count; i++)
            {
                string colUid = oMatrix.Columns.Item(i).UniqueID;
                if (ColumnExists(oDataTable, colUid))
                {
                    oMatrix.Columns.Item(colUid).DataBind.Bind(tableName, colUid);
                }
            }

            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();
        }

        private bool ColumnExists(DataTable dt, string colName)
        {
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                if (dt.Columns.Item(i).Name == colName)
                    return true;
            }
            return false;
        }

        private void Matrix0_PressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.ItemUID == "mtDTrac" && pVal.ColUID == "Check")
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("mtDTrac").Specific;
                SAPbouiCOM.EditText woNum = (SAPbouiCOM.EditText)oMatrix.Columns.Item("WONum").Cells.Item(pVal.Row).Specific;
                if (!string.IsNullOrEmpty(woNum.Value))
                {
                    SAPbouiCOM.CheckBox checkB = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Check").Cells.Item(pVal.Row).Specific;
                    checkB.Checked = false;
                    Program.SBO_Application.SetStatusBarMessage("This line already has a linked Work Order!", BoMessageTime.bmt_Short, false);
                }
            }
        }

        private void Form_ResizeAfter(SBOItemEventArg pVal)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("mtDTrac").Specific;
            oMatrix.AutoResizeColumns();
        }

        private void Button2_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("mtDTrac").Specific;

            // 1. Get all selected rows from Matrix
            List<int> selectedRows = new List<int>();
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                SAPbouiCOM.CheckBox check = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Check").Cells.Item(i).Specific;
                if (check.Checked == true)
                {
                    selectedRows.Add(i);
                }
            }

            //2. Get the information required from QCID and The Sales Order to create Production order for each line
            foreach (int rowIndex in selectedRows)
            {
                string salesOrder = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("SONum").Cells.Item(rowIndex).Specific).Value;
                string LineNum = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("docLine").Cells.Item(rowIndex).Specific).Value;

                // 3. Create the Production order based on the retrieved salesOrder. You will need to implement this logic.
                DBCalls.CreateProductionOrder(salesOrder, LineNum);
            }

            BindDataTableToMatrix("DocTrackInfo", "mtDTrac", openDocEntry);
        }

        private void Form_UnloadBefore(SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form parentForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(Utils.ParentFormUID);
            parentForm.Select();
            Program.SBO_Application.ActivateMenuItem("1304");

        }
    }
}
