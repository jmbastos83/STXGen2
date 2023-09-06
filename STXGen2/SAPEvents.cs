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

        public static string itemCode { get; private set; }
        public static string itemName { get; private set; }
        public static string mLinenum { get; private set; }
        public static string qcid { get; private set; }
        public static string lastClickedMatrixUID { get; set; }
        public static int selectedRow { get; set; }
        public static int SysFormLine { get; private set; }

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

        //internal static void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        //{
        //    BubbleEvent = true;
        //    if (pVal.ItemUID == "10000329" && pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.PopUpIndicator == 0)
        //    {

        //    }
            
        //}

        internal static void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (isEventBeingProcessed)
            {
                return;
            }

            Form activeForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
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

                    if (string.IsNullOrEmpty(qcid))
                    {
                        frmQCalc = new QuoteCalculator();
                        frmQCalc.ParentFormUID = activeForm.UniqueID;

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

                        activeForm.Items.Item("1").Click();
                        mDocLines.Columns.Item("U_STXQC19ID").Editable = false;
                        frmQCalc.LoadFrmByKey(qcid, itemCode, itemName, docCur, unPrice, exRate, DocEntry, ObjType, mLinenum);
                    }
                    else
                    {
                        string formTypeEx = activeForm.TypeEx;
                        string objectTypeCode = DBCalls.GetObjectTypeCodeByFormType(formTypeEx);

                        SAPbouiCOM.DBDataSource oDBDS = activeForm.DataSources.DBDataSources.Item(objectTypeCode);
                        string DocEntry = oDBDS.GetValue("DocEntry", 0);
                        string ObjType = oDBDS.GetValue("ObjType", 0);

                        frmQCalc = new QuoteCalculator();

                        frmQCalc.ParentFormUID = activeForm.UniqueID;
                        if (activeForm.Mode == BoFormMode.fm_UPDATE_MODE)
                        {
                            activeForm.Items.Item("1").Click();
                        }
                        frmQCalc.LoadFrmByKey(qcid, itemCode, itemName, docCur, unPrice, exRate, DocEntry, ObjType, mLinenum);
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