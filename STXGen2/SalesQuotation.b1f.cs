using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System.Xml;

namespace STXGen2
{
    [FormAttribute("149", "SalesQuotation.b1f")]
    class SystemForm1 : SystemFormBase
    {
        private Matrix Matrix0;
        private string ItemCC1;
        private string ItemCC2;
        private string activeFormUID;
        private double resCC1;
        private double resCC2;
        private ButtonCombo ButtonCombo0;

        private bool isChooseFromListTriggered = false;
        private bool isChooseFromListPickerTriggered = false;

        private bool isUpdatingDimensions;
        private string newItmCode;
        private string prevToolNum;
        private string prevPartNum;
        private string prevPartName;
        private string prevLeadTime;
        private Form activeForm;
        private string itemChosen;
        private ItemData selectedItem;

        public string prevItemCode { get; private set; }
        public string STXQCID { get; private set; }
        public bool itmChange { get; private set; }

        public SystemForm1()
        {

        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {

            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("38").Specific));
            this.Matrix0.GotFocusAfter += new SAPbouiCOM._IMatrixEvents_GotFocusAfterEventHandler(this.Matrix0_GotFocusAfter);
            this.Matrix0.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix0_ChooseFromListBefore);
            this.Matrix0.PickerClickedBefore += new SAPbouiCOM._IMatrixEvents_PickerClickedBeforeEventHandler(this.Matrix0_PickerClickedBefore);
            this.Matrix0.KeyDownBefore += new SAPbouiCOM._IMatrixEvents_KeyDownBeforeEventHandler(this.Matrix0_KeyDownBefore);
            this.Matrix0.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix0_ChooseFromListAfter);
            this.ButtonCombo0 = this.GetItem("1").Specific as SAPbouiCOM.ButtonCombo;
            if (this.ButtonCombo0 != null)
            {
                // Handle the error here.
                // This means the cast was unsuccessful.
                this.ButtonCombo0 = ((SAPbouiCOM.ButtonCombo)(this.GetItem("1").Specific));
                this.ButtonCombo0.PressedBefore += new SAPbouiCOM._IButtonComboEvents_PressedBeforeEventHandler(this.ButtonCombo0_PressedBefore);
            }
            

            
            this.OnCustomInitialize();

        }



        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadBefore += new SAPbouiCOM.Framework.FormBase.LoadBeforeHandler(this.Form_LoadBefore);
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.ResizeAfter += new ResizeAfterHandler(this.Form_ResizeAfter);

        }


        private void OnCustomInitialize()
        {

        }

        private void Matrix0_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.DataTable selectedDataTable = null;
            try
            {
                this.UIAPIRawForm.Freeze(true);

                if (pVal.ItemUID == "38" && (pVal.ColUID == "1" || pVal.ColUID == "3") && pVal.ActionSuccess == true)
                {

                    SBOChooseFromListEventArg chooseFromListEventArg = (SBOChooseFromListEventArg)pVal;
                    selectedDataTable = chooseFromListEventArg.SelectedObjects;

                    List<ItemData> itemsList = ItemData.ConvertDataTableToList(selectedDataTable);
                    if (itemsList.Count > 0)
                    {
                        selectedItem = itemsList[0];
                        itemChosen = selectedItem.ItemCode;

                        if (itemChosen != prevItemCode)
                        {
                            itmChange = true;
                        }

                        if (isChooseFromListPickerTriggered == false && isChooseFromListTriggered == true && itmChange == true)
                        {
                            SetItemProperties(selectedItem, pVal.Row); // Adjust method signature and implementation
                            itemChangedVal(selectedItem, pVal.Row);    // Adjust method signature and implementation
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Program.SBO_Application.MessageBox("Error: " + ex.ToString());
            }
            finally
            {
                isChooseFromListTriggered = false;
                this.UIAPIRawForm.Freeze(false);
            }


        }

        private void Matrix0_GotFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {

                if (pVal.ItemUID == "38" && (pVal.ColUID == "1" || pVal.ColUID == "3") && isChooseFromListPickerTriggered == true && isChooseFromListTriggered == false && isUpdatingDimensions == false && itmChange == true)
                {
                    isChooseFromListPickerTriggered = false;

                    SetItemProperties(selectedItem, pVal.Row);
                    itemChangedVal(selectedItem, pVal.Row);
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void itemChangedVal(ItemData selectedItem, int row)
        {
            if (isChooseFromListPickerTriggered == false && isChooseFromListTriggered == false && isUpdatingDimensions == false)
            {
                SAPbouiCOM.DBDataSource oDBDS = this.UIAPIRawForm.DataSources.DBDataSources.Item("OQUT");
                string sapdocEntry = oDBDS.GetValue("DocEntry", 0);
                string sapObjType = oDBDS.GetValue("ObjType", 0);

                SAPbouiCOM.Matrix itemsMatrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("38").Specific;
                SAPbouiCOM.EditText itemCode = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("1").Cells.Item(row).Specific;
                SAPbouiCOM.EditText intLineNo = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("110").Cells.Item(row).Specific;
                SAPbouiCOM.EditText qcidCell = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("U_STXQC19ID").Cells.Item(row).Specific;

                string currentItemCode = selectedItem.ItemCode;

                if (prevItemCode != currentItemCode && !string.IsNullOrEmpty(STXQCID) && itmChange == true)
                {
                    bool confirmGetOper = Program.SBO_Application.MessageBox("Do you want to keep all tool information?", 1, "Yes", "No") == 1;
                    if (confirmGetOper)
                    {
                        QCIDColumnsEnable();
                        //this.UIAPIRawForm.Freeze(true);
                        SAPbouiCOM.EditText toolNum = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("U_STXToolNum").Cells.Item(row).Specific;
                        SAPbouiCOM.EditText partNum = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("U_STXPartNum").Cells.Item(row).Specific;
                        SAPbouiCOM.EditText partName = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("U_STXPartName").Cells.Item(row).Specific;
                        SAPbouiCOM.EditText leadTime = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("U_STXLeadTime").Cells.Item(row).Specific;

                        toolNum.Value = prevToolNum;
                        partNum.Value = prevPartNum;
                        partName.Value = prevPartName;
                        leadTime.Value = prevLeadTime;

                        string newQCID = DBCalls.duplicateQCID(STXQCID, sapdocEntry, sapObjType, intLineNo.Value, itmChange);

                        qcidCell.Active = true;
                        qcidCell.Value = newQCID;
                        //this.UIAPIRawForm.Freeze(false);
                    }
                    itemCode.Active = true;
                    QCIDColumnsDisable();
                    itmChange = false;
                }
            }
        }

        private void SetItemProperties(ItemData selectedDataTable, int row)
        {
            if (selectedDataTable != null)
            {
                newItmCode = selectedDataTable.ItemCode.ToString();
                ItemCC1 = selectedDataTable.U_STXCC1.ToString();
                ItemCC2 = selectedDataTable.U_STXCC2.ToString();

                resCC1 = DBCalls.VerifyCC1(ItemCC1);
                resCC2 = DBCalls.VerifyCC2(ItemCC2);

                SAPbouiCOM.Matrix itemsMatrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("38").Specific;
                UpdateItemDimensions(itemsMatrix, row);
            }
        }

        private void UpdateItemDimensions(Matrix itemsMatrix, int row)
        {
            try
            {
                isUpdatingDimensions = true;

                this.UIAPIRawForm.Freeze(true);
                this.UIAPIRawForm.Select();

                SAPbouiCOM.EditText itemCode = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("1").Cells.Item(row).Specific;
                SAPbouiCOM.EditText itemDim1 = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("2004").Cells.Item(row).Specific;
                SAPbouiCOM.EditText itemDim3 = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("2002").Cells.Item(row).Specific;
                SAPbouiCOM.EditText itemDim1Cogs = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("110000310").Cells.Item(row).Specific;
                SAPbouiCOM.EditText itemDim3Cogs = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("10002041").Cells.Item(row).Specific;

                if (resCC1 == 1)
                {
                    itemDim1Cogs.Active = true;
                    itemDim1Cogs.Value = ItemCC1;
                    itemDim1.Active = true;
                    itemDim1.Value = ItemCC1;
                }
                else if (!string.IsNullOrEmpty(ItemCC1))
                {
                    Program.SBO_Application.SetStatusBarMessage("Invalid Dimension1 Value!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }

                if (resCC2 == 1)
                {
                    itemDim3Cogs.Active = true;
                    itemDim3Cogs.Value = ItemCC2;
                    itemDim3.Active = true;
                    itemDim3.Value = ItemCC2;
                }
                else if (!string.IsNullOrEmpty(ItemCC2))
                {
                    Program.SBO_Application.SetStatusBarMessage("Invalid Dimension3 Value!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }

                itemCode.Active = true;
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);

                isUpdatingDimensions = false;
                isChooseFromListTriggered = false;
            }

        }

        private void Matrix0_KeyDownBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.ItemUID == "38" && (pVal.ColUID == "1" || pVal.ColUID == "3") && pVal.CharPressed == 9)
            {
                isChooseFromListPickerTriggered = true;
            }
            else
            {
                isChooseFromListTriggered = true;
            }
        }

        private void Matrix0_PickerClickedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.ItemUID == "38" && (pVal.ColUID == "1" || pVal.ColUID == "3") && pVal.ActionSuccess == true)
            {
                isChooseFromListPickerTriggered = true;
            }

        }

        private void ButtonCombo0_PressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                this.UIAPIRawForm.Freeze(true);

                itmChange = false;
                SAPbouiCOM.DBDataSource oDBDS = this.UIAPIRawForm.DataSources.DBDataSources.Item("OQUT");
                string sapdocEntry = oDBDS.GetValue("DocEntry", 0);
                string sapObjType = oDBDS.GetValue("ObjType", 0);

                if (pVal.FormMode == 3)
                {

                    string qcidValue = "";
                    SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("38").Specific;


                    for (int i = 1; i <= matrix.RowCount; i++)
                    {
                        SAPbouiCOM.EditText qcidCell = (SAPbouiCOM.EditText)matrix.Columns.Item("U_STXQC19ID").Cells.Item(i).Specific;
                        SAPbouiCOM.EditText intLineNo = (SAPbouiCOM.EditText)matrix.Columns.Item("110").Cells.Item(i).Specific;
                        if (!string.IsNullOrEmpty(qcidCell.Value))
                        {

                            qcidValue = qcidCell.Value;
                            string newQCID = DBCalls.duplicateQCID(qcidValue, sapdocEntry, sapObjType, intLineNo.Value, itmChange);
                            qcidCell.Value = newQCID;
                        }
                    }
                }
                else
                {
                    if (pVal.FormMode == 2)
                    {
                        string qcidValue = "";
                        SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("38").Specific;

                        for (int i = 1; i <= matrix.RowCount; i++)
                        {
                            SAPbouiCOM.EditText qcidCell = (SAPbouiCOM.EditText)matrix.Columns.Item("U_STXQC19ID").Cells.Item(i).Specific;
                            SAPbouiCOM.EditText intLineNo = (SAPbouiCOM.EditText)matrix.Columns.Item("110").Cells.Item(i).Specific;
                            if (!string.IsNullOrEmpty(qcidCell.Value))
                            {

                                qcidValue = qcidCell.Value;
                                int QCIDdocLine = DBCalls.getDocLineofQCID(qcidValue, sapdocEntry, sapObjType);
                                if (QCIDdocLine != int.Parse(intLineNo.Value))
                                {
                                    string newQCID = DBCalls.duplicateQCID(qcidValue, sapdocEntry, sapObjType, intLineNo.Value, itmChange);
                                    qcidCell.Active = true;
                                    qcidCell.Value = newQCID;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
            }
        }

        private void Matrix0_ChooseFromListBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.ItemUID == "38" && (pVal.ColUID == "1" || pVal.ColUID == "3"))
            {
                isChooseFromListTriggered = true;
                SAPbouiCOM.Matrix itemsMatrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("38").Specific;
                SAPbouiCOM.EditText itemCode = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific;
                SAPbouiCOM.EditText qcidCell = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("U_STXQC19ID").Cells.Item(pVal.Row).Specific;
                SAPbouiCOM.EditText toolNum = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("U_STXToolNum").Cells.Item(pVal.Row).Specific;
                SAPbouiCOM.EditText partNum = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("U_STXPartNum").Cells.Item(pVal.Row).Specific;
                SAPbouiCOM.EditText partName = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("U_STXPartName").Cells.Item(pVal.Row).Specific;
                SAPbouiCOM.EditText leadTime = (SAPbouiCOM.EditText)itemsMatrix.Columns.Item("U_STXLeadTime").Cells.Item(pVal.Row).Specific;


                prevItemCode = itemCode.Value;
                prevToolNum = toolNum.Value;
                prevPartNum = partNum.Value;
                prevPartName = partName.Value;
                prevLeadTime = leadTime.Value;
                STXQCID = qcidCell.Value;
            }
        }

        private void Form_LoadBefore(SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            activeForm = (SAPbouiCOM.Form)this.UIAPIRawForm;
            activeFormUID = activeForm.UniqueID;

        }

        private void Form_LoadAfter(SBOItemEventArg pVal)
        {
            QCIDColumnsDisable();
        }

        private void QCIDColumnsDisable()
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("38").Specific;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("U_STXQC19ID");
            oColumn.Editable = false;
            oColumn = oMatrix.Columns.Item("U_STXToolNum");
            oColumn.Editable = false;
            oColumn = oMatrix.Columns.Item("U_STXPartNum");
            oColumn.Editable = false;
            oColumn = oMatrix.Columns.Item("U_STXPartName");
            oColumn.Editable = false;
            oColumn = oMatrix.Columns.Item("U_STXLeadTime");
            oColumn.Editable = false;
        }

        private void QCIDColumnsEnable()
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("38").Specific;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("U_STXQC19ID");
            oColumn.Editable = true;
            oColumn = oMatrix.Columns.Item("U_STXToolNum");
            oColumn.Editable = true;
            oColumn = oMatrix.Columns.Item("U_STXPartNum");
            oColumn.Editable = true;
            oColumn = oMatrix.Columns.Item("U_STXPartName");
            oColumn.Editable = true;
            oColumn = oMatrix.Columns.Item("U_STXLeadTime");
            oColumn.Editable = true;
        }

        private void Form_ResizeAfter(SBOItemEventArg pVal)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("38").Specific;
            oMatrix.AutoResizeColumns();

        }
    }
}
