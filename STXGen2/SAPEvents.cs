using System;
using System.Collections.Generic;
using System.Globalization;
using SAPbouiCOM;

namespace STXGen2
{
    internal class SAPEvents
    {
        

        public static List<int> deletedTexturesList = new List<int>();

        private static SAPbouiCOM.Form oForm;

        private static float exRate = 1;

        public static string itemCode { get; private set; }
        public static string itemName { get; private set; }
        public static string mLinenum { get; private set; }
        public static string qcid { get; private set; }

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

                    //else
                    //{
                    //    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Please add the document before procceding.");
                    //}

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
                    oForm.EnableMenu("1292", true);
                    oForm.EnableMenu("1293", eventInfo.Row > 0);
                }

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        internal static void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.BeforeAction && pVal.MenuUID == "QCalc")
            {
                Form activeForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
                if (activeForm.TypeEx == "149" || activeForm.TypeEx == "139" || activeForm.TypeEx == "140" || activeForm.TypeEx == "133" || activeForm.TypeEx == "179") // Sales Quotation form.
                {
                    Matrix itemMatrix = (Matrix)activeForm.Items.Item("38").Specific; // The item matrix in the Sales Quotation form.
                    int selectedRow = itemMatrix.GetNextSelectedRow();

                    if (selectedRow > -1)
                    {
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

                            string HeaderTable = "";

                            switch (activeForm.TypeEx)
                            {
                                case "149":
                                    //    nQCalc.ObjectType = SAPbobsCOM.BoObjectTypes.oQuotations;
                                    HeaderTable = "OQUT";
                                    break;
                                case "139":
                                    //    nQCalc.ObjectType = SAPbobsCOM.BoObjectTypes.oOrders;
                                    HeaderTable = "ORDR";
                                    break;


                                default:
                                    break;
                            }
                        }

                        QuoteCalculator frmQCalc = new QuoteCalculator();
                        frmQCalc.LoadFrmByKey(qcid, itemCode, itemName, docCur, unPrice, exRate);
                       
                    }
                    else
                    {
                        Program.SBO_Application.SetStatusBarMessage("Please select a row in the item matrix.", BoMessageTime.bmt_Short, false);
                    }
                }
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