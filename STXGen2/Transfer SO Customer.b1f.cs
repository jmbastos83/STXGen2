using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using SAPbouiCOM;
using SAPbobsCOM;

namespace STXGen2
{
    [FormAttribute("STXGen2.Form1", "Transfer SO Customer.b1f")]
    class Form1 : UserFormBase
    {
        private SAPbouiCOM.Conditions oCons;
        private SAPbouiCOM.Condition oCon;

        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.EditText EditText2;
        public static string soTrfDocEntry { get; set; }

        private bool shouldUpdateCardCode = false;
        private bool shouldUpdateCardName = false;
        private string selectedCardName = string.Empty;
        private string selectedCardCode = string.Empty;

        private SAPbouiCOM.EditText EditText0;
        private string cardCode;
        private string cardName;

        public Form1(string docCardCode, string docCardName)
        {
            cardCode = docCardCode;
            cardName = docCardName;
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("CardCode").Specific));
            this.EditText0.ValidateAfter += new SAPbouiCOM._IEditTextEvents_ValidateAfterEventHandler(this.EditText0_ValidateAfter);
            this.EditText0.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.EditText0_LostFocusAfter);
            this.EditText0.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText0_ChooseFromListAfter);
            this.EditText0.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.EditText0_ChooseFromListBefore);
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lbCardCode").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lCardCode").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("DocData").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_5").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("CardName").Specific));
            this.EditText2.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.EditText2_LostFocusAfter);
            this.EditText2.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText2_ChooseFromListAfter);
            this.EditText2.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.EditText2_ChooseFromListBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ActivateAfter += new ActivateAfterHandler(this.Form_ActivateAfter);
            

        }

        private void OnCustomInitialize()
        {
            this.Grid0.Item.Enabled = false;
            this.Button0.Item.Enabled = false;

            PopulateGrid();
        }

        private void PopulateGrid()
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = $"select T1.\"LineNum\",T1.ItemCode, T1.\"Dscription\",T1.\"Quantity\",T1.\"ShipDate\",T1.\"Price\",T1.\"WhsCode\",T1.\"OcrCode\",T1.\"OcrCode2\",T1.\"OcrCode3\",T1.\"OcrCode4\",T1.\"OcrCode5\",\n" +
                            "T1.\"U_STXWONum\",T1.\"U_STXToolNum\",T1.\"U_STXPartNum\",T1.\"U_STXPartName\",T1.\"U_STXLeadTime\",T1.\"U_STXQC19ID\"\n" +
                            "from ORDR T0\n" +
                            "inner join RDR1 T1 on T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                            "where T0.\"DocEntry\" = {0} and T1.\"LineStatus\" = 'O'";

            query = string.Format(query, soTrfDocEntry);
            Grid0.DataTable.ExecuteQuery(query);
        }

        private void EditText0_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ChooseFromList oCfl = this.UIAPIRawForm.ChooseFromLists.Item("cflOCRD");
                oCons = GetCFLConditions(oCfl);

                oCon = oCons.Add();
                oCon.Alias = "CardCode";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_CONTAIN;
                oCon.CondVal = EditText0.Value.ToString();
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCon = oCons.Add();
                oCon.Alias = "CardType";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "C";
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCon = oCons.Add();
                oCon.Alias = "ValidFor";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "Y";
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCon = oCons.Add();
                oCon.Alias = "CardCode";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oCon.CondVal = cardCode;

                oCfl.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private Conditions GetCFLConditions(SAPbouiCOM.ChooseFromList oCfl)
        {
            oCfl.SetConditions(null);
            return oCfl.GetConditions();
        }

        private void EditText0_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                this.UIAPIRawForm.Freeze(true);
                SAPbouiCOM.DataTable selectedDataTable = null;
                SBOChooseFromListEventArg chooseFromListEventArg = (SBOChooseFromListEventArg)pVal;
                selectedDataTable = chooseFromListEventArg.SelectedObjects;
                if (selectedDataTable != null && !selectedDataTable.IsEmpty)
                {
                    selectedCardName = selectedDataTable.GetValue("CardName", 0).ToString();
                    selectedCardCode = selectedDataTable.GetValue("CardCode", 0).ToString();

                    shouldUpdateCardName = true;  // set the flag

                    SAPbouiCOM.EditText cardCode = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("CardCode").Specific;
                    cardCode.Value = selectedCardCode;
                }
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
            }
        }

        private void EditText2_ChooseFromListBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ChooseFromList oCfl = this.UIAPIRawForm.ChooseFromLists.Item("cflOCRD2");
                oCons = GetCFLConditions(oCfl);

                oCon = oCons.Add();
                oCon.Alias = "CardName";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_CONTAIN;
                oCon.CondVal = EditText2.Value.ToString();
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCon = oCons.Add();
                oCon.Alias = "CardType";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "C";
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCon = oCons.Add();
                oCon.Alias = "ValidFor";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "Y";
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCon = oCons.Add();
                oCon.Alias = "CardName";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oCon.CondVal = cardName;

                oCfl.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void EditText2_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                this.UIAPIRawForm.Freeze(true);
                SAPbouiCOM.DataTable selectedDataTable = null;
                SBOChooseFromListEventArg chooseFromListEventArg = (SBOChooseFromListEventArg)pVal;
                selectedDataTable = chooseFromListEventArg.SelectedObjects;
                // Check if the selectedDataTable is not null
                if (selectedDataTable != null && !selectedDataTable.IsEmpty)
                {
                    selectedCardName = selectedDataTable.GetValue("CardName", 0).ToString();
                    selectedCardCode = selectedDataTable.GetValue("CardCode", 0).ToString();

                    shouldUpdateCardCode = true;  // set the flag

                    SAPbouiCOM.EditText cardName = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("CardName").Specific;
                    cardName.Value = selectedCardName;
                }
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
            }
        }

        private void Form_ActivateAfter(SBOItemEventArg pVal)
        {
            if (shouldUpdateCardCode)
            {
                this.UIAPIRawForm.Freeze(true);
                try
                {
                    // Populate CardCode
                    SAPbouiCOM.EditText cardCode = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("CardCode").Specific;
                    cardCode.Value = selectedCardCode;

                    shouldUpdateCardCode = false;  // reset the flag
                }
                finally
                {
                    this.UIAPIRawForm.Freeze(false);
                }
            }

            if (shouldUpdateCardName)
            {
                this.UIAPIRawForm.Freeze(true);
                try
                {
                    // Populate CardName
                    SAPbouiCOM.EditText cardName = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("CardName").Specific;
                    cardName.Value = selectedCardName;

                    shouldUpdateCardName = false;  // reset the flag
                }
                finally
                {
                    this.UIAPIRawForm.Freeze(false);
                }
            }
        }

        private void EditText2_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (shouldUpdateCardCode)
            {
                this.UIAPIRawForm.Freeze(true);
                try
                {
                    // Populate CardCode
                    SAPbouiCOM.EditText cardCode = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("CardCode").Specific;
                    cardCode.Value = selectedCardCode;

                    shouldUpdateCardCode = false;  // reset the flag
                }
                finally
                {
                    this.UIAPIRawForm.Freeze(false);
                }
            }

        }

        private void EditText0_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (shouldUpdateCardName)
            {
                this.UIAPIRawForm.Freeze(true);
                try
                {
                    // Populate CardName
                    SAPbouiCOM.EditText cardName = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("CardName").Specific;
                    cardName.Value = selectedCardName;

                    shouldUpdateCardName = false;  // reset the flag
                }
                finally
                {
                    this.UIAPIRawForm.Freeze(false);
                }
            }
        }

        private void EditText0_ValidateAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (!string.IsNullOrEmpty(EditText0.Value))
            {
                Button0.Item.Enabled = true;
            }

        }

        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)this.UIAPIRawForm.Items.Item("DocData").Specific;
            List<int> selectedRowsIndices = new List<int>();

            for (int rowIndex = 0; rowIndex < oGrid.Rows.Count; rowIndex++)
            {
                if (oGrid.Rows.IsSelected(rowIndex))
                {
                    selectedRowsIndices.Add(rowIndex);
                }
            }

            if (selectedRowsIndices.Count == 0)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Please select at least one row to proceed.");
                return;
            }
            DocTransfInfo(soTrfDocEntry, selectedRowsIndices);

            this.UIAPIRawForm.Refresh();

        }

        private bool DocCloseTrfLines(string soTrfDocEntry, List<int> selectedRowsIndices, string newDocEntryStr)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)this.UIAPIRawForm.Items.Item("DocData").Specific;
            SAPbobsCOM.Documents oOrder = (SAPbobsCOM.Documents)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

            if (oOrder.GetByKey(int.Parse(soTrfDocEntry)))
            {

                foreach (int rowIndex in selectedRowsIndices)
                {
                    oOrder.Lines.SetCurrentLine(int.Parse(oGrid.DataTable.GetValue("LineNum", rowIndex).ToString()));
                    oOrder.Lines.LineStatus = BoStatus.bost_Close;
                }

                oOrder.DocumentReferences.ReferencedObjectType = ReferencedObjectTypeEnum.rot_SalesOrder;
                oOrder.DocumentReferences.ReferencedDocEntry = int.Parse(newDocEntryStr);
                oOrder.DocumentReferences.IssueDate = DateTime.Today.AddDays(2);
                oOrder.DocumentReferences.Remark = "Transfer Customer: Original Document";


                // Commit the changes
                int result = oOrder.Update();
                if (result == 0)
                    return true;
                else
                {
                    int errCode;
                    string errMsg;
                    Utils.oCompany.GetLastError(out errCode, out errMsg);
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox($"Error updating sales order: {errCode} - {errMsg}");
                    return false;
                }
            }
            else
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox($"Sales Order with DocEntry {soTrfDocEntry} not found.");
                return false;
            }
        }

        private void DocTransfInfo(string soTrfDocEntry, List<int> selectedRowsIndices)
        {
            DateTime docDueDate = DateTime.Today;
            string numAtCard = "";
            string docCur = "";

            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)this.UIAPIRawForm.Items.Item("DocData").Specific;

            string query = "select T0.\"DocNum\", T0.\"ObjType\",T0.\"DocDueDate\",T0.\"NumAtCard\",T0.\"DocCur\",T0.\"SlpCode\",\n" +
                           "T0.\"U_STXMSEGID1\",T0.\"U_STXMarSeg\",T0.\"U_STXMSEGID2\",T0.\"U_STXIndCode\",T0.\"U_STXBrand\",T0.\"U_STXBRANDID\",T0.\"U_STXNBOID\",T0.\"U_STXOEMPgm\",T0.\"U_STXOEM\"\n" +
                           "from ORDR T0\n" +
                           "where T0.\"DocEntry\" = {0}";
            query = string.Format(query, soTrfDocEntry);
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(query);

            docDueDate = (DateTime)rs.Fields.Item("DocDueDate").Value;
            numAtCard = rs.Fields.Item("NumAtCard").Value.ToString();
            docCur = rs.Fields.Item("DocCur").Value.ToString();
            int objType = int.Parse(rs.Fields.Item("ObjType").Value.ToString());

            SAPbobsCOM.Documents oOrder = (SAPbobsCOM.Documents)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

            // Set Sales Order Header info. Assuming CardCode is set to a specific Business Partner for this example.
            oOrder.CardCode = this.EditText0.Value;
            oOrder.DocDueDate = docDueDate;
            oOrder.DocCurrency = docCur;
            oOrder.NumAtCard = numAtCard;
            oOrder.SalesPersonCode = (int)rs.Fields.Item("SlpCode").Value;
            oOrder.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;

            oOrder.UserFields.Fields.Item("U_STXMSEGID1").Value = rs.Fields.Item("U_STXMSEGID1").Value.ToString();
            oOrder.UserFields.Fields.Item("U_STXMarSeg").Value = rs.Fields.Item("U_STXMarSeg").Value.ToString();
            oOrder.UserFields.Fields.Item("U_STXMSEGID2").Value = rs.Fields.Item("U_STXMSEGID2").Value.ToString();
            oOrder.UserFields.Fields.Item("U_STXIndCode").Value = rs.Fields.Item("U_STXIndCode").Value.ToString();
            oOrder.UserFields.Fields.Item("U_STXBrand").Value = rs.Fields.Item("U_STXBrand").Value.ToString();
            oOrder.UserFields.Fields.Item("U_STXBRANDID").Value = rs.Fields.Item("U_STXBRANDID").Value.ToString();
            oOrder.UserFields.Fields.Item("U_STXNBOID").Value = rs.Fields.Item("U_STXNBOID").Value.ToString();
            oOrder.UserFields.Fields.Item("U_STXOEMPgm").Value = rs.Fields.Item("U_STXOEMPgm").Value.ToString();
            oOrder.UserFields.Fields.Item("U_STXOEM").Value = rs.Fields.Item("U_STXOEM").Value.ToString();
            



            // Loop through selected rows and add them to the Sales Order
            foreach (int rowIndex in selectedRowsIndices)
            {
                string itemCode = oGrid.DataTable.GetValue("ItemCode", rowIndex).ToString();
                double quantity = Convert.ToDouble(oGrid.DataTable.GetValue("Quantity", rowIndex));

                oOrder.Lines.ItemCode = itemCode;
                oOrder.Lines.Quantity = quantity;

                oOrder.Lines.ShipDate = (DateTime)oGrid.DataTable.GetValue("ShipDate", rowIndex);
                oOrder.Lines.Price = Convert.ToDouble(oGrid.DataTable.GetValue("Price", rowIndex));
                oOrder.Lines.WarehouseCode = oGrid.DataTable.GetValue("WhsCode", rowIndex).ToString();

                oOrder.Lines.CostingCode = oGrid.DataTable.GetValue("OcrCode", rowIndex).ToString();
                oOrder.Lines.COGSCostingCode = oGrid.DataTable.GetValue("OcrCode", rowIndex).ToString();
                oOrder.Lines.CostingCode2 = oGrid.DataTable.GetValue("OcrCode2", rowIndex).ToString();
                oOrder.Lines.COGSCostingCode2 = oGrid.DataTable.GetValue("OcrCode2", rowIndex).ToString();
                oOrder.Lines.CostingCode3 = oGrid.DataTable.GetValue("OcrCode3", rowIndex).ToString();
                oOrder.Lines.COGSCostingCode3 = oGrid.DataTable.GetValue("OcrCode3", rowIndex).ToString();
                oOrder.Lines.CostingCode4 = oGrid.DataTable.GetValue("OcrCode4", rowIndex).ToString();
                oOrder.Lines.COGSCostingCode4 = oGrid.DataTable.GetValue("OcrCode4", rowIndex).ToString();
                oOrder.Lines.CostingCode5 = oGrid.DataTable.GetValue("OcrCode5", rowIndex).ToString();
                oOrder.Lines.COGSCostingCode5 = oGrid.DataTable.GetValue("OcrCode5", rowIndex).ToString();

                oOrder.Lines.UserFields.Fields.Item("U_STXWONum").Value = oGrid.DataTable.GetValue("U_STXWONum", rowIndex).ToString();
                oOrder.Lines.UserFields.Fields.Item("U_STXToolNum").Value = oGrid.DataTable.GetValue("U_STXToolNum", rowIndex).ToString();
                oOrder.Lines.UserFields.Fields.Item("U_STXPartNum").Value = oGrid.DataTable.GetValue("U_STXPartNum", rowIndex).ToString();
                oOrder.Lines.UserFields.Fields.Item("U_STXPartName").Value = oGrid.DataTable.GetValue("U_STXPartName", rowIndex).ToString();
                oOrder.Lines.UserFields.Fields.Item("U_STXLeadTime").Value = oGrid.DataTable.GetValue("U_STXLeadTime", rowIndex).ToString();
                oOrder.Lines.UserFields.Fields.Item("U_STXQC19ID").Value = oGrid.DataTable.GetValue("U_STXQC19ID", rowIndex).ToString();

                // Add the line to the Sales Order object
                oOrder.Lines.Add();
            }

            oOrder.DocumentReferences.ReferencedObjectType = ReferencedObjectTypeEnum.rot_SalesOrder;
            oOrder.DocumentReferences.ReferencedDocEntry = int.Parse(soTrfDocEntry);
            oOrder.DocumentReferences.IssueDate = DateTime.Today.AddDays(2);
            oOrder.DocumentReferences.Remark = "Transfer Customer: Original Document";
            oOrder.DocumentReferences.Add();

            oOrder.Comments = $"Generated via Customer Transfer from Sales Order {rs.Fields.Item("DocNum").Value.ToString()}";

            int result = oOrder.Add();

            if (result != 0)
            {
                // Error handling
                string errMsg = Utils.oCompany.GetLastErrorDescription();
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox($"Failed to create Sales Order. Error: {errMsg}");
            }
            else
            {
                string newDocEntryStr = Utils.oCompany.GetNewObjectKey();
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Sales Order created successfully!", BoMessageTime.bmt_Short, false);

                #region Open new sales order

                SAPbouiCOM.BoFormObjectEnum formType = SAPbouiCOM.BoFormObjectEnum.fo_Order;
                string formTypeStr = ((int)formType).ToString();

                // Open the Sales Order form with the specified DocEntry
                Program.SBO_Application.OpenForm(formType, "DocEntry", newDocEntryStr);
                #endregion

                DocCloseTrfLines(soTrfDocEntry, selectedRowsIndices, newDocEntryStr);
                moveWOtoNewSO(newDocEntryStr);
                PopulateGrid();
            }
        }

        private void moveWOtoNewSO(string newDocEntryStr)
        {
            string query = "select T0.\"DocNum\",T0.\"DocEntry\",T0.\"CardCode\", T0.\"CardName\",T1.\"LineNum\",T0.\"LicTradNum\",T1.\"U_STXWONum\"\n" +
                           "from ORDR T0\n" +
                           "inner join RDR1 T1 on T0.\"DocEntry\" = T1.\"DocEntry\"\n" +
                           "where T0.\"DocEntry\" = {0}";

       
            var formattedQuery = string.Format(query, newDocEntryStr);
            var rs = (Recordset)Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rs.DoQuery(formattedQuery);

            while (!rs.EoF)
            {
                UpdateWOFields(rs, newDocEntryStr);
                rs.MoveNext();
            }
        }

        private void UpdateWOFields(Recordset rs, string newDocEntryStr)
        {
            string docentryWO = DBCalls.getWODocEntry(rs.Fields.Item("U_STXWONum").Value.ToString());
            var prodOrder = (SAPbobsCOM.ProductionOrders)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);

            if (prodOrder.GetByKey(int.Parse(docentryWO)))
            {
                SetProdOrderUserFields(prodOrder, rs);
                HandleDocumentReferences(prodOrder, newDocEntryStr, docentryWO);

                int updateResult = prodOrder.Update();
                if (updateResult != 0)
                {
                    Utils.oCompany.GetLastError(out int errCode, out string errMsg);
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox($"Error updating production order: {errCode} - {errMsg}");
                }
            }
        }

        private void SetProdOrderUserFields(SAPbobsCOM.ProductionOrders prodOrder, Recordset rs)
        {
            prodOrder.UserFields.Fields.Item("U_STXSONum").Value = rs.Fields.Item("DocNum").Value.ToString();
            prodOrder.UserFields.Fields.Item("U_STXSOLineNum").Value = rs.Fields.Item("LineNum").Value.ToString();
            prodOrder.CustomerCode = rs.Fields.Item("CardCode").Value.ToString();
            prodOrder.UserFields.Fields.Item("U_STXCustName").Value = rs.Fields.Item("CardName").Value;
            prodOrder.UserFields.Fields.Item("U_STXLicTradNum").Value = rs.Fields.Item("LicTradNum").Value.ToString();
        }

        private void HandleDocumentReferences(SAPbobsCOM.ProductionOrders prodOrder, string newDocEntryStr, string docentryWO)
        {
            var query2 = "select T0.\"DocEntry\",T0.\"LineNum\",T0.\"RefObjType\",T0.\"RefDocEntr\" from WOR5 T0 where T0.\"DocEntry\" = {0}";
            var formattedQuery2 = string.Format(query2, docentryWO);
            var rs2 = (Recordset)Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rs2.DoQuery(formattedQuery2);

            if (rs2.RecordCount > 0)
            {
                while (!rs2.EoF)
                {
                    if (rs2.Fields.Item("RefObjType").Value.ToString() == "17")
                    {
                        int refLine = int.Parse(rs2.Fields.Item("LineNum").Value.ToString());
                        prodOrder.DocumentReferences.SetCurrentLine(refLine - 1);
                        prodOrder.DocumentReferences.ReferencedDocEntry = int.Parse(newDocEntryStr);
                    }
                    rs2.MoveNext();
                }
            }
            else
            {
                prodOrder.DocumentReferences.Add();
                prodOrder.DocumentReferences.ReferencedObjectType = ReferencedObjectTypeEnum.rot_SalesOrder;
                prodOrder.DocumentReferences.ReferencedDocEntry = int.Parse(newDocEntryStr);
                prodOrder.DocumentReferences.IssueDate = DateTime.Today.AddDays(2);
            }
        }

    }
}
