﻿using System;
using System.Collections.Generic;
using System.Globalization;
using SAPbouiCOM;
using SAPbobsCOM;

namespace STXGen2
{
    internal class SAPEvents
    {

        public static QuoteCalculator frmQCalc;
        public static List<int> deletedTexturesList = new List<int>();

        private static SAPbouiCOM.Form oForm;

        private static float exRate = 1;
        private static bool isEventBeingProcessed = false;
        private static bool docCancelation = false;
        private static string cancelDocEntry;
        public static bool manualQCIDCreation;

        public static string itemCode { get; private set; }
        public static string itemName { get; private set; }
        public static string mLinenum { get; private set; }
        public static string qcid { get; private set; }
        public static string lastClickedMatrixUID { get; set; }
        public static int selectedRow { get; set; }
        public static int SysFormLine { get; private set; }
        public static bool cancelSAPOperation { get; private set; }

        internal static void SBO_Application_RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;

                SAPbouiCOM.MenuItem oMenuItem = null;
                SAPbouiCOM.Menus oMenus = null;
                if ((oForm.TypeEx == "149" || oForm.TypeEx == "139" || oForm.TypeEx == "140" || oForm.TypeEx == "133" || oForm.TypeEx == "179") && eventInfo.BeforeAction == true && eventInfo.ItemUID == "38")
                {
                    selectedRow = eventInfo.Row;
                    if (oForm.Mode == BoFormMode.fm_UPDATE_MODE || oForm.Mode == BoFormMode.fm_OK_MODE)
                    {
                        try
                        {

                            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
                            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SAPbouiCOM.Framework.Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));

                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            oCreationPackage.UniqueID = "QCalc";
                            oCreationPackage.String = "Quote Calculator";
                            oCreationPackage.Position = 2;
                            oCreationPackage.Enabled = true;
                            oMenuItem = SAPbouiCOM.Framework.Application.SBO_Application.Menus.Item("1280"); // Data'
                            oMenus = oMenuItem.SubMenus;
                            if (!oMenus.Exists("QCalc"))
                            {
                                oMenus.AddEx(oCreationPackage);
                            }
                        }
                        catch (Exception ex)
                        {
                            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

                        }
                    }
                }
                else
                {
                    try
                    {
                        bool bMenuQC = SAPbouiCOM.Framework.Application.SBO_Application.Menus.Exists("QCalc");
                        if (bMenuQC == true)
                        {
                            SAPbouiCOM.Framework.Application.SBO_Application.Menus.RemoveEx("QCalc");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);
                    }
                }

                if (eventInfo.FormUID == oForm.UniqueID && eventInfo.ItemUID == "mTextures" && eventInfo.EventType == BoEventTypes.et_RIGHT_CLICK && eventInfo.BeforeAction)
                {
                    BubbleEvent = false;
                    selectedRow = eventInfo.Row;
                    oForm.EnableMenu("1292", true);
                    oForm.EnableMenu("1293", eventInfo.Row > 0);
                }

                if (eventInfo.FormUID == oForm.UniqueID && eventInfo.ItemUID == "mOper" && eventInfo.EventType == BoEventTypes.et_RIGHT_CLICK && eventInfo.BeforeAction)
                {
                    BubbleEvent = false;
                    selectedRow = eventInfo.Row;
                    oForm.EnableMenu("772", true);
                    oForm.EnableMenu("784", true);
                    oForm.EnableMenu("1292", true);
                    oForm.EnableMenu("1293", eventInfo.Row > 0);
                    oForm.EnableMenu("1294", eventInfo.Row > 0);
                }

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        internal static void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {

            BubbleEvent = true;

            if (FormUID == "RelationMap" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.ItemUID == "Gresult" && pVal.BeforeAction)
            {
                SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.Item(FormUID);
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Gresult").Specific;

                int rowIndex = oGrid.GetDataTableRowIndex(pVal.Row); // Get the Index on the datatable of the row selected on the grid

                if (rowIndex >= 0 && oGrid.Rows.Count > rowIndex)
                {
                    try
                    {
                        string objtType = oGrid.DataTable.GetValue("ObjType", rowIndex).ToString();  // The value in the "ObjType" column of the clicked row

                        SAPbouiCOM.EditTextColumn docNumColumn = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Doc. Number");
                        docNumColumn.LinkedObjectType = objtType;  // Change LinkedObjectType based on the value in the "ObjType" column

                    }
                    catch (Exception ex)
                    {
                        Program.SBO_Application.SetStatusBarMessage("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                }
            }

            if (pVal.FormType == 0 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction)
            {
                if (pVal.ItemUID == "1")
                {
                    cancelSAPOperation = false;
                }
                if (pVal.ItemUID == "2")
                {
                    cancelSAPOperation = true;
                }

            }

            if (pVal.FormTypeCount == 2 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "1" && !pVal.BeforeAction && pVal.ActionSuccess)
            {
                docCancelation = true;
            }


            if (FormUID == "DocTracker" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.ItemUID == "mtDTrac" && pVal.ColUID == "WONum" && pVal.BeforeAction)
            {
                SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.Item(FormUID);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("mtDTrac").Specific;


                try
                {
                    string woNumValue = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("WONum").Cells.Item(pVal.Row).Specific).Value;
                    string docentryWO = DBCalls.getWODocEntry(woNumValue);

                    // Cancel the default linked button behavior
                    BubbleEvent = false;

                    Program.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_ProductionOrder, "", docentryWO);
                }
                catch (Exception ex)
                {
                    Program.SBO_Application.SetStatusBarMessage("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.ItemUID == "38" && pVal.ColUID == "U_STXWONum" && pVal.BeforeAction)
            {
                SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.Item(FormUID);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                try
                {
                    string woNumValue = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_STXWONum").Cells.Item(pVal.Row).Specific).Value;
                    string docentryWO = DBCalls.getWODocEntry(woNumValue);

                    // Cancel the default linked button behavior
                    BubbleEvent = false;

                    Program.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_ProductionOrder, "", docentryWO);
                }
                catch (Exception ex)
                {
                    Program.SBO_Application.SetStatusBarMessage("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }

            if (FormUID == "ToolFindUI")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.ItemUID == "gdTInfo" && pVal.ColUID == "Sales Order" && pVal.BeforeAction)
                {
                    SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.Item(FormUID);
                    SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("gdTInfo").Specific;

                    try
                    {
                        string soNum = oGrid.DataTable.GetValue("Sales Order", pVal.Row).ToString();
                        string docentrySO = DBCalls.getSODocEntry(soNum);

                        // Cancel the default linked button behavior
                        BubbleEvent = false;

                        Program.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_Order, "", docentrySO);
                    }
                    catch (Exception ex)
                    {
                        Program.SBO_Application.SetStatusBarMessage("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.ItemUID == "gdTInfo" && pVal.ColUID == "Quote Number" && pVal.BeforeAction)
                {
                    SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.Item(FormUID);
                    SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("gdTInfo").Specific;

                    try
                    {
                        string soNum = oGrid.DataTable.GetValue("Quote Number", pVal.Row).ToString();
                        string docentryQT = DBCalls.getQTDocEntry(soNum);

                        // Cancel the default linked button behavior
                        BubbleEvent = false;

                        Program.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_Quotation, "", docentryQT);
                    }
                    catch (Exception ex)
                    {
                        Program.SBO_Application.SetStatusBarMessage("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.ItemUID == "gdTInfo" && pVal.ColUID == "Tool Number" && pVal.BeforeAction)
                {
                    SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.Item(FormUID);
                    SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("gdTInfo").Specific;

                    try
                    {
                        string imagePath = oGrid.DataTable.GetValue("Picture", pVal.Row).ToString();

                        if (!string.IsNullOrEmpty(imagePath))
                        {
                            // Open the image using your custom method
                            Utils.OpenImage(imagePath);
                        }
                        BubbleEvent = false;
                    }
                    catch (Exception ex)
                    {
                        Program.SBO_Application.SetStatusBarMessage("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                }
            }
        }

        internal static void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (isEventBeingProcessed)
            {
                return;
            }

            Form activeForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;

            if (activeForm.TypeEx == "139" && (activeForm.Mode == BoFormMode.fm_FIND_MODE || activeForm.Mode == BoFormMode.fm_ADD_MODE))
            {

                activeForm.Freeze(true);
                SAPbouiCOM.Item itemRelmap = activeForm.Items.Item("RelMap");
                if (itemRelmap != null)
                {
                    SAPbouiCOM.Button relMap = itemRelmap.Specific as SAPbouiCOM.Button;
                    if (relMap != null)
                    {
                        relMap.Item.Enabled = false;
                    }
                }

                SAPbouiCOM.Item itemTracker = activeForm.Items.Item("DocTrak");
                if (itemTracker != null)
                {
                    SAPbouiCOM.Button docTrack = itemTracker.Specific as SAPbouiCOM.Button;
                    if (docTrack != null)
                    {
                        docTrack.Item.Enabled = false;
                    }
                }
                activeForm.Freeze(false);
            }

            if ((activeForm.TypeEx == "149" || activeForm.TypeEx == "139" || activeForm.TypeEx == "140" || activeForm.TypeEx == "133" || activeForm.TypeEx == "179") && pVal.BeforeAction && pVal.MenuUID == "QCalc")
            {
                Matrix itemMatrix = (Matrix)activeForm.Items.Item("38").Specific;
                

                if (selectedRow > -1)
                {
                    SysFormLine = selectedRow;
                    SAPbouiCOM.EditText etItemCode = (SAPbouiCOM.EditText)itemMatrix.Columns.Item("1").Cells.Item(selectedRow).Specific;
                    itemCode = etItemCode.Value;
                    SAPbouiCOM.EditText etDescription = (SAPbouiCOM.EditText)itemMatrix.Columns.Item("3").Cells.Item(selectedRow).Specific;
                    itemName = etDescription.Value;
                    SAPbouiCOM.EditText etLineNum = (SAPbouiCOM.EditText)itemMatrix.Columns.Item("110").Cells.Item(selectedRow).Specific;
                    mLinenum = etLineNum.Value;
                    SAPbouiCOM.EditText etQCID = (SAPbouiCOM.EditText)itemMatrix.Columns.Item("U_STXQC19ID").Cells.Item(selectedRow).Specific;
                    qcid = etQCID.Value;

                    SAPbouiCOM.EditText UnPrice = (SAPbouiCOM.EditText)itemMatrix.Columns.Item("14").Cells.Item(selectedRow).Specific;
                    string unPrice = UnPrice.Value;

                    string docCur = "";

                    SAPbouiCOM.ComboBox CurSource = (SAPbouiCOM.ComboBox)activeForm.Items.Item("70").Specific;
                    switch (CurSource.Value)
                    {
                        case "L":
                            docCur = Utils.MainCurrency;
                            exRate = 1;
                            break;
                        case "S":
                            docCur = Utils.SystemCurrency;
                            exRate = 1;
                            break;
                        case "C":
                            SAPbouiCOM.ComboBox DocCur = (SAPbouiCOM.ComboBox)activeForm.Items.Item("63").Specific;
                            docCur = DocCur.Value;
                            SAPbouiCOM.EditText ExRate = (SAPbouiCOM.EditText)activeForm.Items.Item("64").Specific;
                            float.TryParse(ExRate.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out exRate);
                            break;
                    }
                    try
                    {
                        activeForm.Freeze(true);
                        if (string.IsNullOrEmpty(qcid) || qcid == "0")
                        {
                            frmQCalc = new QuoteCalculator();
                            Utils.ParentFormUID = activeForm.UniqueID;

                            string formTypeEx = activeForm.TypeEx;
                            string objectTypeCode = DBCalls.GetObjectTypeCodeByFormType(formTypeEx);

                            SAPbouiCOM.DBDataSource oDBDS = activeForm.DataSources.DBDataSources.Item(objectTypeCode);
                            string DocEntry = oDBDS.GetValue("DocEntry", 0);
                            string ObjType = oDBDS.GetValue("ObjType", 0);

                            qcid = frmQCalc.AddUDO(DocEntry, ObjType, mLinenum);
                            //nQCalc.UpdateUDO(qcid, DocEntry, ObjType, mLinenum);

                            SAPbouiCOM.Matrix mDocLines = ((SAPbouiCOM.Matrix)(activeForm.Items.Item("38").Specific));
                            mDocLines.Columns.Item("U_STXQC19ID").Editable = true;
                            etQCID.Value = qcid;

                            manualQCIDCreation = true;
                            activeForm.Items.Item("1").Click();
                            manualQCIDCreation = false;
                            if (!cancelSAPOperation)
                            {
                                mDocLines.Columns.Item("U_STXQC19ID").Editable = false;
                                frmQCalc.LoadFrmByKey(qcid, itemCode, itemName, docCur, unPrice, exRate, DocEntry, ObjType, mLinenum);
                            }
                            else
                            {
                                //DBCalls.revertQCIDCreation(qcid);
                                //etQCID.Value = string.Empty;
                                ((SAPbouiCOM.EditText)mDocLines.Columns.Item("1").Cells.Item(selectedRow).Specific).Active = true;
                                mDocLines.Columns.Item("U_STXQC19ID").Editable = false;
                            }
                        }
                        else
                        {
                            string formTypeEx = activeForm.TypeEx;
                            string objectTypeCode = DBCalls.GetObjectTypeCodeByFormType(formTypeEx);

                            SAPbouiCOM.DBDataSource oDBDS = activeForm.DataSources.DBDataSources.Item(objectTypeCode);
                            string DocEntry = oDBDS.GetValue("DocEntry", 0);
                            string ObjType = oDBDS.GetValue("ObjType", 0);

                            frmQCalc = new QuoteCalculator();

                            Utils.ParentFormUID = activeForm.UniqueID;
                            if (activeForm.Mode == BoFormMode.fm_UPDATE_MODE)
                            {
                                activeForm.Items.Item("1").Click();
                            }
                            frmQCalc.LoadFrmByKey(qcid, itemCode, itemName, docCur, unPrice, exRate, DocEntry, ObjType, mLinenum);
                        }
                    }
                    catch (Exception ex)
                    {
                        Program.SBO_Application.SetStatusBarMessage(ex.ToString(), BoMessageTime.bmt_Short, false);
                    }
                    finally
                    {
                        activeForm.Freeze(false);
                    }
                    

                }
                else
                {
                    Program.SBO_Application.SetStatusBarMessage("Please select a row in the item matrix.", BoMessageTime.bmt_Short, false);
                }
            }

            if (activeForm.TypeEx == "STXGen2.QuoteCalculator")
            {
                // Handle events for the add-on form
                if ((pVal.MenuUID == "1292" || pVal.MenuUID == "1293") && !pVal.BeforeAction)
                {
                    if (!string.IsNullOrEmpty(lastClickedMatrixUID))
                    {
                        SAPbouiCOM.Matrix activeMatrix = (SAPbouiCOM.Matrix)activeForm.Items.Item(lastClickedMatrixUID).Specific;
                        HandleQCMatrixMenuEvent(Program.SBO_Application, ref pVal, activeMatrix);
                        return;
                    }
                    else
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("No matrix clicked.");
                    }
                }
            }
            if (activeForm.TypeEx == "139" && pVal.BeforeAction && pVal.MenuUID == "1293")
            {
                bool canDelete = DBCalls.VerifyWOCreated(activeForm, selectedRow);
                if (!canDelete)
                {
                    BubbleEvent = false;
                }
            }

            if (activeForm.TypeEx == "65211" && pVal.BeforeAction && pVal.MenuUID == "1284")
            {
                SAPbouiCOM.DBDataSource dbDataSource = (SAPbouiCOM.DBDataSource)activeForm.DataSources.DBDataSources.Item(0);
                cancelDocEntry = dbDataSource.GetValue("DocEntry", 0).Trim();

            }
            if (activeForm.TypeEx == "65211" && !pVal.BeforeAction && pVal.MenuUID == "1284")
            {
                if (docCancelation)
                {
                    DBCalls.QCIDUpdateWOinfo(activeForm, cancelDocEntry);
                    docCancelation = false;
                }
               
            }
        }

        private static void HandleQCMatrixMenuEvent(SAPbouiCOM.Application sBO_Application, ref MenuEvent pVal, Matrix activeMatrix)
        {
            try
            {
                SAPbouiCOM.Form oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
                if (pVal.MenuUID == "1292" && !pVal.BeforeAction)
                {
                    oForm.Freeze(true);
                    if (SAPEvents.lastClickedMatrixUID == "mTextures")
                    {
                        QCEvents.AddLineToTexturesMatrix(oForm, activeMatrix, selectedRow);
                    }
                    else if (SAPEvents.lastClickedMatrixUID == "mOper")
                    {
                        QCEvents.AddLineToOperationMatrix(oForm, activeMatrix, selectedRow);
                    }
                }
                else if (pVal.MenuUID == "1293" && !pVal.BeforeAction)
                {
                    oForm.Freeze(true);
                    if (SAPEvents.lastClickedMatrixUID == "mTextures")
                    {
                        QCEvents.RemoveLinefromTexturesMatrix(oForm, activeMatrix, selectedRow);


                    }
                    else if (SAPEvents.lastClickedMatrixUID == "mOper")
                    {
                        QCEvents.RemoveLinefromOperationMatrix(oForm, activeMatrix, selectedRow);
                    }


                }

            }
            catch (Exception ex)
            {
                Program.SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }

        }


        internal static void SBO_Application_AppEvent(BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}