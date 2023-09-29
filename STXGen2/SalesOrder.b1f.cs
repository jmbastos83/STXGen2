using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using SAPbouiCOM;

namespace STXGen2
{
    [FormAttribute("139", "SalesOrder.b1f")]
    class SystemForm3 : SystemFormBase
    {
        private SAPbouiCOM.EditText stxMKSeg1;
        private SAPbouiCOM.EditText stxMKSEG2;
        private SAPbouiCOM.EditText stxBrand;
        private SAPbouiCOM.EditText stxNBOID;
        private SAPbouiCOM.EditText stxOEMPgm;
        private SAPbouiCOM.EditText stxOEM;
        private SAPbouiCOM.EditText stxGKAM;
        private SAPbouiCOM.EditText stxMK1ID;
        private SAPbouiCOM.EditText stxBrandID;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.Button Button0;
        private ButtonCombo ButtonCombo0;
        private bool isHistoric = false;
        private bool itmChange;

        private Matrix Matrix0;
        private ItemData selectedItem;
        private string itemChosen;
        private bool isChooseFromListPickerTriggered;
        private bool isChooseFromListTriggered;
        private bool isUpdatingDimensions;
        private string prevToolNum = string.Empty;
        private string prevPartNum = string.Empty;
        private string prevPartName = string.Empty;
        private string prevLeadTime = string.Empty;
        private string newItmCode;
        private string ItemCC1;
        private string ItemCC2;
        private double resCC1;
        private double resCC2;

        private DocumentTracker formDocTracker;

        public string prevItemCode { get; private set; }
        public string STXQCID { get; private set; }

        public SystemForm3()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.stxMKSeg1 = ((SAPbouiCOM.EditText)(this.GetItem("MKSeg1").Specific));
            this.stxMKSEG2 = ((SAPbouiCOM.EditText)(this.GetItem("MKSEG2").Specific));
            this.stxBrand = ((SAPbouiCOM.EditText)(this.GetItem("STXBrand").Specific));
            this.stxNBOID = ((SAPbouiCOM.EditText)(this.GetItem("NBOID").Specific));
            this.stxOEMPgm = ((SAPbouiCOM.EditText)(this.GetItem("OEMPgm").Specific));
            this.stxOEM = ((SAPbouiCOM.EditText)(this.GetItem("OEM").Specific));
            this.stxGKAM = ((SAPbouiCOM.EditText)(this.GetItem("GKAM").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lMKSeg1").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lMKSEG2").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_9").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("lNBOID").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("lSTXBrand").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("lOEM").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("lGKAM").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("ClrNBO").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.stxMK1ID = ((SAPbouiCOM.EditText)(this.GetItem("MK1ID").Specific));
            this.stxBrandID = ((SAPbouiCOM.EditText)(this.GetItem("MK2ID").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("38").Specific));
            this.Matrix0.GotFocusAfter += new SAPbouiCOM._IMatrixEvents_GotFocusAfterEventHandler(this.Matrix0_GotFocusAfter);
            this.Matrix0.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix0_ChooseFromListAfter);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Revision").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("DocTrak").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("RelMap").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadBefore += new LoadBeforeHandler(this.Form_LoadBefore);

        }

        private void OnCustomInitialize()
        {
            Button0.Item.Left = this.stxMKSeg1.Item.Left + this.stxMKSeg1.Item.Width + 3;
            Button0.Item.Top = this.stxMKSeg1.Item.Top + 2;
            this.ButtonCombo0 = this.GetItem("1").Specific as SAPbouiCOM.ButtonCombo;
            if (this.ButtonCombo0 != null)
            {
                isHistoric = false;
                this.ButtonCombo0 = ((SAPbouiCOM.ButtonCombo)(this.GetItem("1").Specific));
                this.ButtonCombo0.PressedBefore += new SAPbouiCOM._IButtonComboEvents_PressedBeforeEventHandler(this.ButtonCombo0_PressedBefore);
                this.ButtonCombo0.PressedAfter += new SAPbouiCOM._IButtonComboEvents_PressedAfterEventHandler(this.ButtonCombo0_PressedAfter);
            }
            else
            {
                isHistoric = true;
            }
        }

        private void ButtonCombo0_PressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (SystemForm1.SOcopyToTrigger == true)
                {
                    return;
                }
                this.UIAPIRawForm.Freeze(true);

                itmChange = false;
                SAPbouiCOM.DBDataSource oDBDS = this.UIAPIRawForm.DataSources.DBDataSources.Item("ORDR");
                string sapdocEntry = oDBDS.GetValue("DocEntry", 0);
                string sapBaseEntry = oDBDS.GetValue("BaseEntry", 0);
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
                            SAPbouiCOM.EditText baseType = (SAPbouiCOM.EditText)matrix.Columns.Item("43").Cells.Item(i).Specific;
                            SAPbouiCOM.EditText baseEntry = (SAPbouiCOM.EditText)matrix.Columns.Item("45").Cells.Item(i).Specific;
                            SAPbouiCOM.EditText baseLineNo = (SAPbouiCOM.EditText)matrix.Columns.Item("46").Cells.Item(i).Specific;

                            if (!string.IsNullOrEmpty(qcidCell.Value))
                            {
                                if (!string.IsNullOrEmpty(baseType.Value))
                                {
                                    sapObjType = baseType.Value;
                                }
                                if (!string.IsNullOrEmpty(baseEntry.Value))
                                {
                                    sapdocEntry = baseEntry.Value;
                                }
                                if (!string.IsNullOrEmpty(baseLineNo.Value))
                                {
                                    sapdocEntry = baseEntry.Value;
                                }

                                qcidValue = qcidCell.Value;
                                int QCIDdocLine = DBCalls.getDocLineofQCID(qcidValue, sapdocEntry, sapObjType);
                                if (QCIDdocLine != (!string.IsNullOrEmpty(baseLineNo.Value) ? int.Parse(baseLineNo.Value) : int.Parse(intLineNo.Value)))
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

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                this.UIAPIRawForm.Freeze(true);
                ItemData.DisableNBOinfo(this.UIAPIRawForm);
                this.stxGKAM.Value = "";
                this.stxOEM.Value = "";
                this.stxOEMPgm.Value = "";
                this.stxNBOID.Value = "";
                this.stxBrand.Value = "";
                this.stxMKSEG2.Value = "";
                this.stxMKSeg1.Value = "";
                this.stxMK1ID.Value = "";
                this.stxBrandID.Value = "";
                ItemData.EnableNBOinfo(this.UIAPIRawForm);
            }
            catch (Exception ex)
            {
                Program.SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, true);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
            }
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
                            SetItemProperties(selectedItem, pVal.Row);
                            itemChangedVal(selectedItem, pVal.Row);
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

        private void itemChangedVal(ItemData selectedItem, int row)
        {
            if (isChooseFromListPickerTriggered == false && isChooseFromListTriggered == false && isUpdatingDimensions == false)
            {
                SAPbouiCOM.DBDataSource oDBDS = this.UIAPIRawForm.DataSources.DBDataSources.Item("ORDR");
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
                        ItemData.QCIDColumnsEnable(this.UIAPIRawForm);
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
                    ItemData.QCIDColumnsDisable(this.UIAPIRawForm);
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
            catch (Exception ex)
            {
                Program.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }

        }

        private void ButtonCombo0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ActionSuccess == true)
            {
                SystemForm1.SOcopyToTrigger = false;
            }
        }

        private EditText EditText0;
        private Button Button1;
        

        private void Button1_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.DBDataSource dbDataSource = (SAPbouiCOM.DBDataSource)this.UIAPIRawForm.DataSources.DBDataSources.Item(0);
            Utils.ParentFormUID = this.UIAPIRawForm.UniqueID;
            string docEntry = dbDataSource.GetValue("DocEntry", 0).Trim();

            if (!IsFormOpen("DocTracker"))
            {
                DocumentTracker.openDocEntry = docEntry;
                formDocTracker = new DocumentTracker();
                
                formDocTracker.UIAPIRawForm.Visible = true;
            }
            else
            {
               SAPbouiCOM.Form existingForm = Program.SBO_Application.Forms.Item("DocTracker");
               existingForm.Visible = true;
            }
        }

        private bool IsFormOpen(string formUID)
        {
            for (int i = 0; i < Program.SBO_Application.Forms.Count; i++)
            {
                if (Program.SBO_Application.Forms.Item(i).UniqueID == formUID)
                {
                    return true; 
                }
            }
            return false;
        }

        private Button Button2;
        private RelationshipMap formRelationshipMap;

        private void Button2_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.DBDataSource dbDataSource = (SAPbouiCOM.DBDataSource)this.UIAPIRawForm.DataSources.DBDataSources.Item(0);
            Utils.ParentFormUID = this.UIAPIRawForm.UniqueID;
            string docEntry = dbDataSource.GetValue("DocEntry", 0).Trim();

            if (!IsFormOpen("RelationMap"))
            {
                RelationshipMap.relDocEntry = docEntry;
                formRelationshipMap = new RelationshipMap();

                formRelationshipMap.UIAPIRawForm.Visible = true;
            }
            else
            {
                SAPbouiCOM.Form existingForm = Program.SBO_Application.Forms.Item("RelationMap");
                existingForm.Visible = true;
            }

        }

        private void Form_LoadBefore(SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.FormMode == 0 || pVal.FormMode == 3)
            {

            }
            throw new System.NotImplementedException();

        }
    }
}
