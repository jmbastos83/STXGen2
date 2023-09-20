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

        public SAPbouiCOM.Conditions oCons;
        public SAPbouiCOM.Condition oCon;

        private string ItemCC1;
        private string ItemCC2;
        private string activeFormUID;
        private double resCC1;
        private double resCC2;

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

        private EditText stxMKSeg1;
        private EditText stxMKSEG2;
        private EditText stxBrand;
        private EditText stxNBOID;
        private EditText stxOEMPgm;
        private EditText stxOEM;
        private EditText stxGKAM;

        private EditText stxMK1ID;
        private EditText stxMK2ID;
        private EditText stxBrandID;

        private EditText stxRevision;

        private StaticText StaticText0;
        private StaticText StaticText1;
        private StaticText StaticText2;
        private StaticText StaticText3;
        private StaticText StaticText4;
        private StaticText StaticText5;
        private StaticText StaticText6;

        private Button Button0;
        private bool isHistoric = false;
        private bool addNewDocTrigger = false;

        private Matrix Matrix0;
        private ButtonCombo ButtonCombo0;
        private ComboBox ComboBox0;

        public string prevItemCode { get; private set; }
        public string STXQCID { get; private set; }
        public bool itmChange { get; private set; }
        public static bool SOcopyToTrigger { get; set; }

        public SystemForm1()
        {
            SOcopyToTrigger = false;
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.stxMKSeg1 = ((SAPbouiCOM.EditText)(this.GetItem("MKSeg1").Specific));
            this.stxMKSeg1.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.stxMKSeg1_ChooseFromListAfter);
            this.stxMKSEG2 = ((SAPbouiCOM.EditText)(this.GetItem("MKSEG2").Specific));
            this.stxMKSEG2.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.stxMKSEG2_ChooseFromListAfter);
            this.stxMKSEG2.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.stxMKSEG2_ChooseFromListBefore);
            this.stxBrand = ((SAPbouiCOM.EditText)(this.GetItem("STXBrand").Specific));
            this.stxBrand.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.stxBrand_ChooseFromListAfter);
            this.stxBrand.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.stxBrand_ChooseFromListBefore);
            this.stxNBOID = ((SAPbouiCOM.EditText)(this.GetItem("NBOID").Specific));
            this.stxNBOID.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.stxNBOID_ChooseFromListAfter);
            this.stxNBOID.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.stxNBOID_ChooseFromListBefore);
            this.stxOEMPgm = ((SAPbouiCOM.EditText)(this.GetItem("OEMPgm").Specific));
            this.stxOEM = ((SAPbouiCOM.EditText)(this.GetItem("OEM").Specific));
            this.stxGKAM = ((SAPbouiCOM.EditText)(this.GetItem("GKAM").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("38").Specific));
            this.Matrix0.GotFocusAfter += new SAPbouiCOM._IMatrixEvents_GotFocusAfterEventHandler(this.Matrix0_GotFocusAfter);
            this.Matrix0.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix0_ChooseFromListBefore);
            this.Matrix0.PickerClickedBefore += new SAPbouiCOM._IMatrixEvents_PickerClickedBeforeEventHandler(this.Matrix0_PickerClickedBefore);
            this.Matrix0.KeyDownBefore += new SAPbouiCOM._IMatrixEvents_KeyDownBeforeEventHandler(this.Matrix0_KeyDownBefore);
            this.Matrix0.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix0_ChooseFromListAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lMKSeg1").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lMKSEG2").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lSTXBrand").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("lNBOID").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("lOEMPgm").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("lOEM").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("lGKAM").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("ClrNBO").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.stxMK1ID = ((SAPbouiCOM.EditText)(this.GetItem("MK1ID").Specific));
            this.stxMK2ID = ((SAPbouiCOM.EditText)(this.GetItem("MK2ID").Specific));
            this.stxBrandID = ((SAPbouiCOM.EditText)(this.GetItem("BrandID").Specific));
            this.stxRevision = ((SAPbouiCOM.EditText)(this.GetItem("Revision").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("10000329").Specific));
            this.ComboBox0.ComboSelectBefore += new SAPbouiCOM._IComboBoxEvents_ComboSelectBeforeEventHandler(this.ComboBox0_ComboSelectBefore);
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
            ItemData.QCIDColumnsDisable(this.UIAPIRawForm);
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

                string currentItemCode = selectedItem.ItemCode;

                if (prevItemCode != currentItemCode && !string.IsNullOrEmpty(STXQCID) && itmChange == true)
                {
                    bool confirmGetOper = Program.SBO_Application.MessageBox("Do you want to keep all tool information?", 1, "Yes", "No") == 1;
                    if (confirmGetOper)
                    {
                        ItemData.QCIDColumnsEnable(this.UIAPIRawForm);
                  
                        itemsMatrix.SetCellWithoutValidation(row, "U_STXToolNum", prevToolNum);
                        itemsMatrix.SetCellWithoutValidation(row, "U_STXPartNum", prevPartNum);
                        itemsMatrix.SetCellWithoutValidation(row, "U_STXPartName", prevPartName);
                        itemsMatrix.SetCellWithoutValidation(row, "U_STXLeadTime", prevLeadTime);

                        string newQCID = DBCalls.duplicateQCID(STXQCID, sapdocEntry, sapObjType, intLineNo.Value, itmChange);
                        itemsMatrix.SetCellWithoutValidation(row, "U_STXQC19ID", newQCID);
                   
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
                    addNewDocTrigger = true;
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
                            matrix.SetCellWithoutValidation(i, "U_STXQC19ID", newQCID);
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

        private void ButtonCombo0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                this.UIAPIRawForm.Freeze(true);

                SAPbouiCOM.DBDataSource oDBDS = this.UIAPIRawForm.DataSources.DBDataSources.Item("OQUT");
                string sapdocEntry = oDBDS.GetValue("DocEntry", 0);
                string sapObjType = oDBDS.GetValue("ObjType", 0);

                if (pVal.FormMode == 1 && addNewDocTrigger == true)
                {
                    addNewDocTrigger = false;
                    string qcidValue = "";
                    SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("38").Specific;

                    for (int i = 1; i <= matrix.RowCount; i++)
                    {
                        SAPbouiCOM.EditText qcidCell = (SAPbouiCOM.EditText)matrix.Columns.Item("U_STXQC19ID").Cells.Item(i).Specific;
                        SAPbouiCOM.EditText intLineNo = (SAPbouiCOM.EditText)matrix.Columns.Item("110").Cells.Item(i).Specific;
                        if (!string.IsNullOrEmpty(qcidCell.Value))
                        {
                            qcidValue = qcidCell.Value;
                            DBCalls.UpdateQCIDBaseDoc(qcidValue, sapdocEntry,intLineNo.Value);
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
            ItemData.QCIDColumnsDisable(this.UIAPIRawForm);
            if (pVal.FormMode == 3)
            {
                stxRevision.Value = "A";
            }

        }

        private void Form_ResizeAfter(SBOItemEventArg pVal)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("38").Specific;
            oMatrix.AutoResizeColumns();

        }


        private SAPbouiCOM.Conditions GetCFLConditions(SAPbouiCOM.ChooseFromList oCfl)
        {
            oCfl.SetConditions(null);
            return oCfl.GetConditions();
        }

        private void stxMKSEG2_ChooseFromListBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ChooseFromList oCfl = this.UIAPIRawForm.ChooseFromLists.Item("CFLMKSEG2");
                oCons = GetCFLConditions(oCfl);

                oCon = oCons.Add();
                oCon.Alias = "U_MKSeg1Name";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = stxMKSeg1.Value.ToString();

                oCfl.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void stxBrand_ChooseFromListBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ChooseFromList oCfl = this.UIAPIRawForm.ChooseFromLists.Item("CFLBRANDS");
                oCons = GetCFLConditions(oCfl);

                oCon = oCons.Add();
                oCon.Alias = "U_MKSeg1Name";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = stxMKSeg1.Value.ToString();
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

                oCon = oCons.Add();
                oCon.Alias = "Code";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oCon.CondVal = "-1";

                oCfl.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void stxNBOID_ChooseFromListBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ChooseFromList oCfl = this.UIAPIRawForm.ChooseFromLists.Item("CFLNBO");
                oCons = GetCFLConditions(oCfl);
                //oCfl.SetConditions(null);
                //oCons = oCfl.GetConditions();

                if (string.IsNullOrEmpty(stxNBOID.Value.ToString()) && (!string.IsNullOrEmpty(stxBrand.Value.ToString())))
                {
                    SetNBOIDCFLConditions(stxBrand.Value.ToString());
                }

                oCfl.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void SetNBOIDCFLConditions(string v)
        {
            oCon = oCons.Add();
            oCon.BracketOpenNum = 2;
            oCon.Alias = "U_BrandName";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = stxBrand.Value.ToString();
            oCon.BracketCloseNum = 1;
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

            oCon = oCons.Add();
            oCon.BracketOpenNum = 1;
            oCon.Alias = "U_NickName";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_START;
            oCon.CondVal = "PH-";
            oCon.BracketCloseNum = 1;
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

            oCon = oCons.Add();
            oCon.BracketOpenNum = 1;
            oCon.Alias = "U_NickName";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_START;
            oCon.CondVal = "PH_";
            oCon.BracketCloseNum = 2;
        }

        private void stxNBOID_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ItemUID == "NBOID" && pVal.ActionSuccess == true)
                {
                    SAPbouiCOM.DataTable selectedDataTable = null;

                    SBOChooseFromListEventArg chooseFromListEventArg = (SBOChooseFromListEventArg)pVal;
                    selectedDataTable = chooseFromListEventArg.SelectedObjects;

                    if (selectedDataTable != null)
                    {
                        this.UIAPIRawForm.Freeze(true);

                        var result = DBCalls.GetDataByNBO(selectedDataTable.GetValue("Code", 0).ToString());
                        ItemData.DisableNBOinfo(this.UIAPIRawForm);
                        SAPbouiCOM.EditText eMkseg1 = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("MKSeg1").Specific;
                        SAPbouiCOM.EditText eMk1ID = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("MK1ID").Specific;
                        SAPbouiCOM.EditText eBrand = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("STXBrand").Specific;
                        SAPbouiCOM.EditText eBrandID = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("BrandID").Specific;
                        SAPbouiCOM.EditText eOEMPgm = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("OEMPgm").Specific;
                        SAPbouiCOM.EditText eOEM = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("OEM").Specific;
                        SAPbouiCOM.EditText eGKam = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("GKAM").Specific;
                        (string sMkSeg1Name,string sMkseg1ID, string sBrandName,string sBrandID, string sOEM, string sOEMProgram, string sGKAM) = result.Value;

                        if ((sOEMProgram.StartsWith("PH-") || sOEMProgram.StartsWith("PH_")) && eMkseg1.Value.ToString() != "")
                        {
                            eOEMPgm.Value = sOEMProgram;
                            eOEM.Value = sOEM;
                            eGKam.Value = sGKAM;
                        }
                        else
                        {
                            eMkseg1.Value = sMkSeg1Name;
                            eMk1ID.Value = sMkseg1ID;
                            eBrand.Value = sBrandName;
                            eBrandID.Value = sBrandID;
                            eOEMPgm.Value = sOEMProgram;
                            eOEM.Value = sOEM;
                            eGKam.Value = sGKAM;
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                ItemData.EnableNBOinfo(this.UIAPIRawForm);
                this.UIAPIRawForm.Freeze(false);
            }
        }


        private void stxMKSeg1_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.DataTable selectedDataTable = null;
            SBOChooseFromListEventArg chooseFromListEventArg = (SBOChooseFromListEventArg)pVal;
            selectedDataTable = chooseFromListEventArg.SelectedObjects;

            SAPbouiCOM.EditText mk1segID =(SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("MK1ID").Specific;
            mk1segID.Value = selectedDataTable.GetValue("Code", 0).ToString();

        }

        private void stxMKSEG2_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.DataTable selectedDataTable = null;
            SBOChooseFromListEventArg chooseFromListEventArg = (SBOChooseFromListEventArg)pVal;
            selectedDataTable = chooseFromListEventArg.SelectedObjects;

            SAPbouiCOM.EditText mk2segID = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("MK2ID").Specific;
            mk2segID.Value = selectedDataTable.GetValue("Code", 0).ToString();

        }

        private void stxBrand_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.DataTable selectedDataTable = null;
            SBOChooseFromListEventArg chooseFromListEventArg = (SBOChooseFromListEventArg)pVal;
            selectedDataTable = chooseFromListEventArg.SelectedObjects;

            SAPbouiCOM.EditText brandID = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("BrandID").Specific;
            brandID.Value = selectedDataTable.GetValue("Code", 0).ToString();
        }


        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
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
            catch (Exception)
            {
                throw;
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
            }

        }



        private void ComboBox0_ComboSelectBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (ComboBox0 != null)
            {
                // To get the currently selected value before the change
                string currentValue = ComboBox0.Value;

                // To get the newly selected value 
                string newValue = ComboBox0.ValidValues.Item(pVal.PopUpIndicator).Value;

                if (currentValue == "Copy To" && newValue == "Sales Order")
                {
                    SOcopyToTrigger = true;
                }
            }

        }
    }
}
