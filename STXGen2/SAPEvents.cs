using System;
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
        private static bool deleteGRPORow = false;

        public static string ItemCode { get; private set; }
        public static string ItemName { get; private set; }
        public static string MLinenum { get; private set; }
        public static string Qcid { get; private set; }
        public static string LastClickedMatrixUID { get; set; }
        public static int SelectedRow { get; set; }
        public static int SysFormLine { get; private set; }
        public static bool CancelSAPOperation { get; private set; }

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
                    SelectedRow = eventInfo.Row;
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

                if ((oForm.TypeEx == "721") && eventInfo.BeforeAction == true && eventInfo.ItemUID == "13" && eventInfo.EventType == BoEventTypes.et_RIGHT_CLICK)
                {
                    SelectedRow = eventInfo.Row;
                }

                if (eventInfo.FormUID == oForm.UniqueID && eventInfo.ItemUID == "mTextures" && eventInfo.EventType == BoEventTypes.et_RIGHT_CLICK && eventInfo.BeforeAction)
                {
                    BubbleEvent = false;
                    SelectedRow = eventInfo.Row;
                    oForm.EnableMenu("1292", true);
                    oForm.EnableMenu("1293", eventInfo.Row > 0);
                }

                if (eventInfo.FormUID == oForm.UniqueID && eventInfo.ItemUID == "mOper" && eventInfo.EventType == BoEventTypes.et_RIGHT_CLICK && eventInfo.BeforeAction)
                {
                    BubbleEvent = false;
                    SelectedRow = eventInfo.Row;
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

            if (FormUID == "RelationMap")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.ItemUID == "Gresult" && pVal.ColUID == "Doc. Number" && pVal.BeforeAction)
                {
                    SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.Item(FormUID);
                    SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Gresult").Specific;

                    int rowIndex = oGrid.GetDataTableRowIndex(pVal.Row);

                    if (rowIndex >= 0 && oGrid.Rows.Count > rowIndex)
                    {
                        try
                        {
                            string objtType = oGrid.DataTable.GetValue("ObjType", rowIndex).ToString();
                            string oDocentry = oGrid.DataTable.GetValue("DocEntry", rowIndex).ToString();
                            string oDocLine = oGrid.DataTable.GetValue("Doc. Line", rowIndex).ToString();
                            SAPbouiCOM.EditTextColumn docNumColumn = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Doc. Number");
                            docNumColumn.LinkedObjectType = objtType;
                            BubbleEvent = false;

                            Program.SBO_Application.OpenForm((SAPbouiCOM.BoFormObjectEnum)Enum.Parse(typeof(SAPbouiCOM.BoFormObjectEnum), objtType), "", oDocentry);

                        }
                        catch (Exception ex)
                        {
                            Program.SBO_Application.SetStatusBarMessage("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        }
                    }
                }
            }


            if (pVal.FormType == 0 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction)
            {
                if (pVal.ItemUID == "1")
                {
                    CancelSAPOperation = false;
                }
                if (pVal.ItemUID == "2")
                {
                    CancelSAPOperation = true;
                }

            }

            if (pVal.FormTypeCount == 2 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "1" && !pVal.BeforeAction && pVal.ActionSuccess)
            {
                docCancelation = true;
            }


            if (FormUID == "DocTracker")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.ItemUID == "mtDTrac" && pVal.ColUID == "WONum" && pVal.BeforeAction)
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

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.ItemUID == "mtDTrac" && pVal.ColUID == "SONum" && pVal.BeforeAction)
                {
                    SAPbouiCOM.Form oForm = Program.SBO_Application.Forms.Item(FormUID);
                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("mtDTrac").Specific;

                    try
                    {
                        string soNum = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("SONum").Cells.Item(pVal.Row).Specific).Value;
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
            }

            if (pVal.FormTypeEx == "139" && (pVal.FormMode == (int)BoFormMode.fm_FIND_MODE || pVal.FormMode == (int)BoFormMode.fm_ADD_MODE) && !pVal.BeforeAction && pVal.EventType == BoEventTypes.et_FORM_DRAW)
            {
                SAPbouiCOM.Form activeForm = Program.SBO_Application.Forms.Item(FormUID);

                activeForm.Freeze(true);
                try
                {
                    SAPbouiCOM.Item itemRelmap = activeForm.Items.Item("RelMap");
                    if (itemRelmap?.Specific is SAPbouiCOM.Button relMap)
                    {
                        relMap.Item.Enabled = false;
                    }

                    SAPbouiCOM.Item itemTracker = activeForm.Items.Item("DocTrak");
                    if (itemTracker?.Specific is SAPbouiCOM.Button docTrack)
                    {
                        docTrack.Item.Enabled = false;
                    }

                }
                catch (Exception)
                {

                }
                finally
                {
                    activeForm.Freeze(false);
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

            Form activeForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;

            if (isEventBeingProcessed)
            {
                return;
            }


            if ((activeForm.TypeEx == "149" || activeForm.TypeEx == "139" || activeForm.TypeEx == "140" || activeForm.TypeEx == "133" || activeForm.TypeEx == "179") && pVal.BeforeAction && pVal.MenuUID == "QCalc")
            {
                Matrix itemMatrix = (Matrix)activeForm.Items.Item("38").Specific;


                if (SelectedRow > -1)
                {
                    SysFormLine = SelectedRow;
                    SAPbouiCOM.EditText etItemCode = (SAPbouiCOM.EditText)itemMatrix.Columns.Item("1").Cells.Item(SelectedRow).Specific;
                    ItemCode = etItemCode.Value;
                    SAPbouiCOM.EditText etDescription = (SAPbouiCOM.EditText)itemMatrix.Columns.Item("3").Cells.Item(SelectedRow).Specific;
                    ItemName = etDescription.Value;
                    SAPbouiCOM.EditText etLineNum = (SAPbouiCOM.EditText)itemMatrix.Columns.Item("110").Cells.Item(SelectedRow).Specific;
                    MLinenum = etLineNum.Value;
                    SAPbouiCOM.EditText etQCID = (SAPbouiCOM.EditText)itemMatrix.Columns.Item("U_STXQC19ID").Cells.Item(SelectedRow).Specific;
                    Qcid = etQCID.Value;

                    SAPbouiCOM.EditText UnPrice = (SAPbouiCOM.EditText)itemMatrix.Columns.Item("14").Cells.Item(SelectedRow).Specific;
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
                        if (string.IsNullOrEmpty(Qcid) || Qcid == "0")
                        {
                            frmQCalc = new QuoteCalculator();
                            Utils.ParentFormUID = activeForm.UniqueID;

                            string formTypeEx = activeForm.TypeEx;
                            string objectTypeCode = DBCalls.GetObjectTypeCodeByFormType(formTypeEx);

                            SAPbouiCOM.DBDataSource oDBDS = activeForm.DataSources.DBDataSources.Item(objectTypeCode);
                            string DocEntry = oDBDS.GetValue("DocEntry", 0);
                            string ObjType = oDBDS.GetValue("ObjType", 0);

                            Qcid = frmQCalc.AddUDO(DocEntry, ObjType, MLinenum);
                            //nQCalc.UpdateUDO(qcid, DocEntry, ObjType, mLinenum);

                            SAPbouiCOM.Matrix mDocLines = ((SAPbouiCOM.Matrix)(activeForm.Items.Item("38").Specific));
                            mDocLines.Columns.Item("U_STXQC19ID").Editable = true;
                            etQCID.Value = Qcid;

                            manualQCIDCreation = true;
                            activeForm.Items.Item("1").Click();
                            manualQCIDCreation = false;
                            if (!CancelSAPOperation)
                            {
                                mDocLines.Columns.Item("U_STXQC19ID").Editable = false;
                                frmQCalc.LoadFrmByKey(Qcid, ItemCode, ItemName, docCur, unPrice, exRate, DocEntry, ObjType, MLinenum);
                            }
                            else
                            {
                                //DBCalls.revertQCIDCreation(qcid);
                                //etQCID.Value = string.Empty;
                                ((SAPbouiCOM.EditText)mDocLines.Columns.Item("1").Cells.Item(SelectedRow).Specific).Active = true;
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
                            frmQCalc.LoadFrmByKey(Qcid, ItemCode, ItemName, docCur, unPrice, exRate, DocEntry, ObjType, MLinenum);
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
                    if (!string.IsNullOrEmpty(LastClickedMatrixUID))
                    {
                        SAPbouiCOM.Matrix activeMatrix = (SAPbouiCOM.Matrix)activeForm.Items.Item(LastClickedMatrixUID).Specific;
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
                bool canDelete = DBCalls.VerifyWOCreated(activeForm, SelectedRow);
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

            if (activeForm.TypeEx == "721" && pVal.BeforeAction && pVal.MenuUID == "1293")
            {
                deleteGRPORow = true;
            }

            if (activeForm.TypeEx == "721" && !pVal.BeforeAction && pVal.MenuUID == "1293")
            {
                if (deleteGRPORow)
                {
                    Goods_Receipt.TrigrDeleteRow = true;
                    Goods_Receipt.UpdateitmsCboxes(activeForm, SelectedRow);
                    deleteGRPORow = false;
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
                    if (SAPEvents.LastClickedMatrixUID == "mTextures")
                    {
                        QCEvents.AddLineToTexturesMatrix(oForm, activeMatrix, SelectedRow);
                    }
                    else if (SAPEvents.LastClickedMatrixUID == "mOper")
                    {
                        QCEvents.AddLineToOperationMatrix(oForm, activeMatrix, SelectedRow);
                    }
                }
                else if (pVal.MenuUID == "1293" && !pVal.BeforeAction)
                {
                    oForm.Freeze(true);
                    if (SAPEvents.LastClickedMatrixUID == "mTextures")
                    {
                        QCEvents.RemoveLinefromTexturesMatrix(oForm, activeMatrix, SelectedRow);


                    }
                    else if (SAPEvents.LastClickedMatrixUID == "mOper")
                    {
                        QCEvents.RemoveLinefromOperationMatrix(oForm, activeMatrix, SelectedRow);
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