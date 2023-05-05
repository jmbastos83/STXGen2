using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;

namespace STXGen2
{
    [FormAttribute("STXGen2.QuoteCalculator", "QuoteCalculator.b1f")]
    class QuoteCalculator : UserFormBase
    {
        private const int FixedMatrixHeight = 110;
        bool loadingForm = true;

        public static string selectedUOM { get; set; } = "";
        public static string oldLengthValue { get; set; } = "";
        public static string oldWidthValue { get; set; } = "";
        public static int selectedMatrixRow { get; set; } = 0;
        public static object mtxMaxLineID { get; set; } = "";
        public static string previousUOM { get; set; } = "";
        public static string currentPrice { get; private set; } = "";
        public static double DocExRate { get; private set; } = 0;
        public static string currentLength { get; private set; } = "";
        public static string currentWidth { get; private set; } = "";
        public static string fileName { get; private set; } = "";
        public static string ToolfileName { get; private set; } = "";
        public static string newImagePath { get; private set; } = "";
        public string parttDescr { get; private set; } = "";
        public int lastBtnOpselection { get; private set; } = 0;
        

        private SAPbouiCOM.EditText QCDocEntry;
        private SAPbouiCOM.EditText QCItemCode;
        private SAPbouiCOM.EditText QCItemName;

        private SAPbouiCOM.EditText QCToolNum;
        private SAPbouiCOM.EditText QCPartNum;

        private SAPbouiCOM.EditText QCNElem;
        private SAPbouiCOM.EditText QCPartName;


        private SAPbouiCOM.Folder FGeneral;
        private SAPbouiCOM.Matrix mTextures;
        private SAPbouiCOM.Matrix mOCosts;
        private SAPbouiCOM.Matrix mOperations;



        private SAPbouiCOM.Folder FOperations;

        private SAPbouiCOM.Button ButtonOk;
        private SAPbouiCOM.Button ButtonCancel;
        private SAPbouiCOM.ButtonCombo BtnGetOPC;


        private SAPbouiCOM.EditText QCPartDesc;
        private SAPbouiCOM.EditText QCPartType;
        private SAPbouiCOM.EditText QCSubPart;
        private SAPbouiCOM.EditText SPartDescr;
        private SAPbouiCOM.StaticText lItemCode;
        private SAPbouiCOM.StaticText lToolNum;
        private SAPbouiCOM.StaticText lPartNum;
        private SAPbouiCOM.StaticText lItemName;
        private SAPbouiCOM.StaticText lNElem;
        private SAPbouiCOM.StaticText lPartName;
        private SAPbouiCOM.StaticText lPartDesc;
        private SAPbouiCOM.StaticText lPartType;

        private SAPbouiCOM.ComboBox UnMsr;
        private SAPbouiCOM.EditText QCLength;
        private SAPbouiCOM.EditText QCWidth;
        private SAPbouiCOM.EditText QCArea;
        private SAPbouiCOM.EditText EditText4;

        private SAPbouiCOM.EditText UnPrice;
        private SAPbouiCOM.EditText LCPrice;
        private SAPbouiCOM.EditText QCDocCur;
        private SAPbouiCOM.EditText LCCurr;

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.StaticText StaticText4;



        public QuoteCalculator()
        {
            try
            {
                if (Program.SBO_Application != null)
                {
                    Program.SBO_Application.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(QCEvents.SBO_Application_MenuEvent);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.QCToolNum = ((SAPbouiCOM.EditText)(this.GetItem("QCToolNum").Specific));
            this.QCPartNum = ((SAPbouiCOM.EditText)(this.GetItem("QCPartNum").Specific));
            this.QCNElem = ((SAPbouiCOM.EditText)(this.GetItem("QCNElem").Specific));
            this.QCPartName = ((SAPbouiCOM.EditText)(this.GetItem("QCPartN").Specific));
            this.ButtonOk = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            //  this.ButtonOk.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.ButtonOk_ClickBefore);
            this.ButtonOk.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.ButtonOk_ClickAfter);
            this.ButtonOk.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.ButtonOk_PressedAfter);
            this.ButtonOk.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.ButtonOk_PressedBefore);
            this.ButtonCancel = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.ButtonCancel.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.ButtonCancel_ClickAfter);
            this.FOperations = ((SAPbouiCOM.Folder)(this.GetItem("FOper").Specific));
            this.FOperations.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.FOperations_PressedAfter);
            this.FGeneral = ((SAPbouiCOM.Folder)(this.GetItem("FGen").Specific));
            this.mTextures = ((SAPbouiCOM.Matrix)(this.GetItem("mTextures").Specific));
            this.mTextures.MatrixLoadAfter += new SAPbouiCOM._IMatrixEvents_MatrixLoadAfterEventHandler(this.mTextures_MatrixLoadAfter);
            this.mTextures.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.mTextures_ClickAfter);
            this.mTextures.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.mTextures_ChooseFromListAfter);
            this.mOCosts = ((SAPbouiCOM.Matrix)(this.GetItem("mOCosts").Specific));
            this.mOCosts.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.mOCosts_ClickAfter);
            this.QCPartDesc = ((SAPbouiCOM.EditText)(this.GetItem("QCPartDesc").Specific));
            this.QCPartType = ((SAPbouiCOM.EditText)(this.GetItem("QCPartType").Specific));
            this.QCPartType.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.QCPartType_ChooseFromListAfter);
            this.QCPartType.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.QCPartType_ChooseFromListBefore);
            this.QCSubPart = ((SAPbouiCOM.EditText)(this.GetItem("QCSubPart").Specific));
            this.QCSubPart.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.QCSubPart_ChooseFromListBefore);
            this.SPartDescr = ((SAPbouiCOM.EditText)(this.GetItem("SPartDescr").Specific));
            this.lItemCode = ((SAPbouiCOM.StaticText)(this.GetItem("lItemCode").Specific));
            this.lToolNum = ((SAPbouiCOM.StaticText)(this.GetItem("lToolNum").Specific));
            this.lPartNum = ((SAPbouiCOM.StaticText)(this.GetItem("lPartNum").Specific));
            this.lItemName = ((SAPbouiCOM.StaticText)(this.GetItem("lItemName").Specific));
            this.lNElem = ((SAPbouiCOM.StaticText)(this.GetItem("lNElem").Specific));
            this.lPartName = ((SAPbouiCOM.StaticText)(this.GetItem("lPartName").Specific));
            this.lPartDesc = ((SAPbouiCOM.StaticText)(this.GetItem("lPartDesc").Specific));
            this.lPartType = ((SAPbouiCOM.StaticText)(this.GetItem("lPartType").Specific));
            this.UnMsr = ((SAPbouiCOM.ComboBox)(this.GetItem("UnMsr").Specific));
            this.UnMsr.ComboSelectBefore += new SAPbouiCOM._IComboBoxEvents_ComboSelectBeforeEventHandler(this.UnMsr_ComboSelectBefore);
            this.UnMsr.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.UnMsr_ComboSelectAfter);
            this.QCLength = ((SAPbouiCOM.EditText)(this.GetItem("QCLength").Specific));
            this.QCLength.GotFocusAfter += new SAPbouiCOM._IEditTextEvents_GotFocusAfterEventHandler(this.QCLength_GotFocusAfter);
            this.QCLength.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.QCLength_LostFocusAfter);
            this.QCWidth = ((SAPbouiCOM.EditText)(this.GetItem("QCWidth").Specific));
            this.QCWidth.GotFocusAfter += new SAPbouiCOM._IEditTextEvents_GotFocusAfterEventHandler(this.QCWidth_GotFocusAfter);
            this.QCWidth.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.QCWidth_LostFocusAfter);
            this.QCArea = ((SAPbouiCOM.EditText)(this.GetItem("QCArea").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("QCHeight").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("lDocEntry").Specific));
            this.QCDocEntry = ((SAPbouiCOM.EditText)(this.GetItem("QCDocEntry").Specific));
            this.QCItemCode = ((SAPbouiCOM.EditText)(this.GetItem("QCItemCode").Specific));
            this.QCItemName = ((SAPbouiCOM.EditText)(this.GetItem("QCItemN").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("QCOpA").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_8").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("Item_10").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("Item_11").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("Item_12").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("Item_13").Specific));
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("Item_14").Specific));
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("Item_15").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("Item_17").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("Item_18").Specific));
            this.EditText14 = ((SAPbouiCOM.EditText)(this.GetItem("Item_19").Specific));
            this.UnPrice = ((SAPbouiCOM.EditText)(this.GetItem("UnPrice").Specific));
            this.UnPrice.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.UnPrice_LostFocusAfter);
            this.LCPrice = ((SAPbouiCOM.EditText)(this.GetItem("LCPrice").Specific));
            this.LCPrice.Item.Visible = false;
            this.QCDocCur = ((SAPbouiCOM.EditText)(this.GetItem("QCDocCur").Specific));
            this.LCCurr = ((SAPbouiCOM.EditText)(this.GetItem("LCCurr").Specific));
            this.LCCurr.Item.Visible = false;
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_20").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_21").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("QCObs").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_23").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_24").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_25").Specific));
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_26").Specific));
            this.StaticText15 = ((SAPbouiCOM.StaticText)(this.GetItem("lOpA").Specific));
            this.StaticText16 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_29").Specific));
            this.StaticText17 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_30").Specific));
            this.StaticText18 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_31").Specific));
            this.PictureBox0 = ((SAPbouiCOM.PictureBox)(this.GetItem("Item_32").Specific));
            this.ToolImg = ((SAPbouiCOM.PictureBox)(this.GetItem("ToolImg").Specific));
            this.ToolImg.ClickAfter += new SAPbouiCOM._IPictureBoxEvents_ClickAfterEventHandler(this.ToolImg_ClickAfter);
            this.ToolImg.DoubleClickAfter += new SAPbouiCOM._IPictureBoxEvents_DoubleClickAfterEventHandler(this.ToolImg_DoubleClickAfter);
            this.PicBrowse = ((SAPbouiCOM.Button)(this.GetItem("3").Specific));
            this.PicBrowse.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.PicBrowse_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("ToolPicC").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.QCPinfo1 = ((SAPbouiCOM.EditText)(this.GetItem("QCPinfo1").Specific));
            this.QCPinfo2 = ((SAPbouiCOM.EditText)(this.GetItem("QCPinfo2").Specific));
            this.StaticText19 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_39").Specific));
            this.StaticText20 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_40").Specific));
            this.mOperations = ((SAPbouiCOM.Matrix)(this.GetItem("mOper").Specific));
            this.mOperations.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.mOperations_ClickAfter);
            this.mOperations.LostFocusAfter += new SAPbouiCOM._IMatrixEvents_LostFocusAfterEventHandler(this.mOperations_LostFocusAfter);
            this.BtnGetOPC = ((SAPbouiCOM.ButtonCombo)(this.GetItem("btnGetOp").Specific));
            this.BtnGetOPC.PressedAfter += new SAPbouiCOM._IButtonComboEvents_PressedAfterEventHandler(this.BtnGetOPC_PressedAfter);
            this.BtnGetOPC.ComboSelectAfter += new SAPbouiCOM._IButtonComboEvents_ComboSelectAfterEventHandler(this.BtnGetOPC_ComboSelectAfter);
            //              this.btnGetOP = ((SAPbouiCOM.Button)(this.GetItem("btnGetOP").Specific));
            //              this.btnGetOP.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.btnGetOP_PressedAfter);
            //              this.btnGetOP.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.btnGetOP_PressedBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            this.UnloadAfter += new UnloadAfterHandler(this.Form_UnloadAfter);

        }


        internal void LoadFrmByKey(string qcid, string itemCode, string itemName, string docCur, string unPrice, float exRate)
        {
            try
            {
                this.UIAPIRawForm.Freeze(true);

                this.UIAPIRawForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                CultureInfo customCultureInfo = new CultureInfo("en-US");
                customCultureInfo.NumberFormat.NumberDecimalSeparator = Utils.decSep;
                customCultureInfo.NumberFormat.NumberGroupSeparator = Utils.thousSep;
                DocExRate = double.Parse(exRate.ToString(customCultureInfo));

                QCItemCode.Value = itemCode;
                QCItemName.Value = itemName;
                //eItemTech.Value = Utils.GetItemTech(itemCode);

                // Enable QCDocEntry field temporarily
                QCDocEntry.Item.Enabled = true;

                QCDocEntry.Value = qcid;

                ButtonOk.Item.Click();

                FGeneral.Select();


                //eDocEntry.Item.Enabled = false;
                this.Show();



                //// Translate the labels on the form
                //SAPbouiCOM.Form form = Program.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
                //FormTranslations.SetStaticTextTranslations(form);


                ToolImg.Picture = ToolImg.Picture = Path.Combine(!Directory.Exists(Path.Combine(Utils.oCompany.BitMapPath, "Tools Images")) ? Utils.oCompany.BitMapPath : Path.Combine(Utils.oCompany.BitMapPath, "Tools Images"), ToolImg.Picture);

                BtnGetOPC.ValidValues.Add("1", "Get Operations");
                BtnGetOPC.ValidValues.Add("2", "Get Operations (Grouped)");

                selectedUOM = UnMsr.Selected.Value;
                previousUOM = UnMsr.Selected.Value;



                SAPbouiCOM.EditText currentlength = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("QCLength").Specific;
                oldLengthValue = currentlength.Value;

                SAPbouiCOM.EditText currentWidth = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("QCWidth").Specific;
                oldWidthValue = currentWidth.Value;

                SAPbouiCOM.EditText DocCur = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("QCDocCur").Specific;
                DocCur.Value = docCur;


                oUserDataSource = this.UIAPIRawForm.DataSources.UserDataSources.Add("MyUNPrice", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);

                SAPbouiCOM.EditText UNPrice = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("UnPrice").Specific;
                UNPrice.DataBind.SetBound(true, "", "MyUNPrice");
                oUserDataSource.Value = unPrice;

                if (Utils.MainCurrency != DocCur.Value)
                {
                    if (Utils.DirectRate == "Y")
                    {
                        double LCprice = double.Parse(Regex.Replace((string.IsNullOrEmpty(unPrice) ? "0" : unPrice), $@"[^\d{Utils.decSep}{Utils.thousSep}]", "")) * DocExRate;
                        LCPrice.Value = $"{LCprice.ToString("0.00")} {Utils.MainCurrency}";

                    }
                    else
                    {
                        double LCprice = double.Parse(Regex.Replace((string.IsNullOrEmpty(unPrice) ? "0" : unPrice), $@"[^\d{Utils.decSep}{Utils.thousSep}]", "")) / DocExRate;
                        LCPrice.Value = $"{LCprice.ToString("0.00")} {Utils.MainCurrency}";
                    }
                    LCCurr.Value = Utils.MainCurrency;
                    this.LCPrice.Item.Visible = true;
                    this.LCCurr.Item.Visible = true;
                }

                currentPrice = this.UnPrice.Value;
                QCEvents.CalculateArea(this.UIAPIRawForm.UniqueID, selectedUOM);

                // Check if mTextures matrix has 0 rows and add a new row if needed
                if (mTextures.RowCount == 0)
                {
                    mTextures.AddRow();
                }
                if (mOperations.RowCount == 0)
                {
                    mOperations.AddRow();
                    SAPbouiCOM.EditText newAutoRow = (SAPbouiCOM.EditText)mOperations.Columns.Item("#").Cells.Item(1).Specific;
                    newAutoRow.Value = "1";
                }

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            finally
            {
                // Set focus to QCToolNum field
                QCToolNum.Item.Click();

                // Disable the QCDocEntry field
                QCDocEntry.Item.Enabled = false;

                QCItemCode.Item.Enabled = false;
                QCItemName.Item.Enabled = false;
                mTextures.AutoResizeColumns();
                loadingForm = false;
                this.UIAPIRawForm.Freeze(false);
            }
        }



        private void OnCustomInitialize()
        {

            QCEvents.FillTextureClass(this.UIAPIRawForm);
            QCEvents.FillUnitMeasures(this.UIAPIRawForm);
        }


        private void Form_ResizeAfter(SBOItemEventArg pVal)
        {

            Item tabControlItem = this.UIAPIRawForm.Items.Item("Item_9");
            Item docEntryItem = this.UIAPIRawForm.Items.Item("QCDocEntry");
            Item matrix1Item = this.UIAPIRawForm.Items.Item("mTextures");
            Item matrix2Item = this.UIAPIRawForm.Items.Item("mOCosts");

            Item buttonOKItem = this.UIAPIRawForm.Items.Item("1");
            Item buttonCancelItem = this.UIAPIRawForm.Items.Item("2");

            int formWidth = this.UIAPIRawForm.ClientWidth;
            int availableWidth = tabControlItem.Width - matrix1Item.Left - 5;

            int minHeight = matrix2Item.Top + 20;
            TabControl tabControl = (TabControl)tabControlItem.Specific;

            // Set a fixed height for 5 rows (assuming a row height of 21 pixels)
            int fixedMatrixHeight = 21 * 5;

            docEntryItem.Left = formWidth - docEntryItem.Width - 5;

            matrix1Item.Width = tabControlItem.Width - 20;
            matrix1Item.Height = fixedMatrixHeight;


            QCPinfo1.Item.Top = matrix1Item.Top + matrix1Item.Height + 10;
            //QCPinfo1.Item.Left = matrix1Item.Left;
            QCPinfo1.Item.Width = availableWidth / 2;

            QCPinfo2.Item.Top = QCPinfo1.Item.Top + QCPinfo1.Item.Height + 5;
            QCPinfo2.Item.Width = availableWidth / 2;

            matrix2Item.Top = QCPinfo2.Item.Top + QCPinfo2.Item.Height + 10;  //matrix1Item.Top + matrix1Item.Height + 20; // Add an additional space between the matrices
            matrix2Item.Width = availableWidth / 2;
            matrix2Item.Height = fixedMatrixHeight;

            //Resize the tab control while maintaining the minimum width and height
            tabControlItem.Width = this.UIAPIRawForm.ClientWidth - tabControlItem.Left - 5;
            tabControlItem.Height = minHeight;

            buttonOKItem.Top = this.UIAPIRawForm.ClientHeight - 30;
            buttonCancelItem.Top = this.UIAPIRawForm.ClientHeight - 30;

            mTextures.AutoResizeColumns();
            mOCosts.AutoResizeColumns();
            mOperations.AutoResizeColumns();

        }


        private void mTextures_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.DataTable selectedDataTable = null;
            try
            {
                var application = Program.SBO_Application;
                var oForm = application.Forms.Item(pVal.FormUID);

                if (pVal.ItemUID == "mTextures" && pVal.ColUID == "QCTexture")
                {
                    
                    if (isChooseFromListTriggered)
                    {
                        isChooseFromListTriggered = false;
                        return;
                    }
                    SBOChooseFromListEventArg chooseFromListEventArg = (SBOChooseFromListEventArg)pVal;
                    string chooseFromListId = chooseFromListEventArg.ChooseFromListUID;
                    SAPbouiCOM.ChooseFromList chooseFromList = oForm.ChooseFromLists.Item(chooseFromListId);

                    // Get the selected item from the Choose From List
                    selectedDataTable = chooseFromListEventArg.SelectedObjects;
                    if (selectedDataTable != null && selectedDataTable.Rows.Count > 0)
                    {
                        string TextureCode = selectedDataTable.GetValue("Code", 0).ToString();
                        string TClass = selectedDataTable.GetValue("U_complexityIX", 0).ToString();

                        SAPbouiCOM.Matrix mtxTextures = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("mTextures").Specific;

                        isChooseFromListTriggered = true;
                        ((SAPbouiCOM.EditText)mtxTextures.Columns.Item("QCQuantity").Cells.Item(selectedMatrixRow).Specific).Value = "1";
                        ((SAPbouiCOM.EditText)mtxTextures.Columns.Item("QCCovA").Cells.Item(selectedMatrixRow).Specific).Value = "0 " + selectedUOM + "²";
                        ((SAPbouiCOM.ComboBox)mtxTextures.Columns.Item("QCTClass").Cells.Item(selectedMatrixRow).Specific).Select(TClass, BoSearchKey.psk_ByValue);
                        ((SAPbouiCOM.ComboBox)mtxTextures.Columns.Item("QCGComp").Cells.Item(selectedMatrixRow).Specific).Select("2", BoSearchKey.psk_ByValue);
                        ((SAPbouiCOM.EditText)mtxTextures.Columns.Item("QCTexture").Cells.Item(selectedMatrixRow).Specific).Value = TextureCode;
                        //mtxTextures.AddRow();
                    }
                }

            }
            catch (Exception ex)
            {
                // Log or display the exception message
                Program.SBO_Application.MessageBox("Error: " + ex.Message);
            }
        }



        private void mTextures_MatrixLoadAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                mtxMaxLineID = DBCalls.GetMatrixLastLineID(QCDocEntry.Value);

            }
            catch (Exception ex)
            {

                Program.SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, false);
            }
        }

        private void UnMsr_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {

            bool isUoMAreaChanging = false;

            if (pVal.ItemChanged == true && loadingForm == false)
            {

                // Get the selected Unit of Measure
                SAPbouiCOM.ComboBox uomComboBox = (SAPbouiCOM.ComboBox)this.UIAPIRawForm.Items.Item("UnMsr").Specific;
                selectedUOM = uomComboBox.Selected.Value;

                // Get the current Length and Width values
                EditText edtLength = (EditText)this.UIAPIRawForm.Items.Item("QCLength").Specific;
                EditText edtWidth = (EditText)this.UIAPIRawForm.Items.Item("QCWidth").Specific;

                double length = double.Parse(Regex.Replace((string.IsNullOrEmpty(edtLength.Value) ? "0" : edtLength.Value), $@"[^\d{Utils.decSep}{Utils.thousSep}]", ""));
                double width = double.Parse(Regex.Replace((string.IsNullOrEmpty(edtWidth.Value) ? "0" : edtWidth.Value), $@"[^\d{Utils.decSep}{Utils.thousSep}]", ""));

                // Perform the conversion based on the selected Unit of Measure
                double convertedLength = DBCalls.ConvertDimensions(length, selectedUOM, previousUOM);
                double convertedWidth = DBCalls.ConvertDimensions(width, selectedUOM, previousUOM);

                // Update the Length and Width fields with the converted values
                edtLength.Value = $"{Math.Round(convertedLength, Utils.MeasureDec)}";
                edtWidth.Value = $"{Math.Round(convertedWidth, Utils.MeasureDec)}";

                QCEvents.CalculateArea(this.UIAPIRawForm.UniqueID, selectedUOM);

                // Prompt the user to confirm before updating the value of the edtLength control
                if (selectedUOM != previousUOM)
                {
                    bool confirmUpdate = Program.SBO_Application.MessageBox("Do you want to update the lines coverage area with the new calculated area?", 1, "Yes", "No") == 1;
                    if (confirmUpdate)
                    {
                        isUoMAreaChanging = true;
                        QCEvents.UpdateCovArea(this.UIAPIRawForm, previousUOM, selectedUOM, isUoMAreaChanging);
                    }
                    else
                    {
                        QCEvents.UpdateCovArea(this.UIAPIRawForm, previousUOM, selectedUOM, isUoMAreaChanging);
                    }
                    edtLength.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
        }

        private void UnMsr_ComboSelectBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            // Get the selected Unit of Measure
            SAPbouiCOM.IComboBox uomComboBox = (SAPbouiCOM.ComboBox)this.UIAPIRawForm.Items.Item("UnMsr").Specific;
            if (uomComboBox.Selected != null)
            {
                previousUOM = ((SAPbouiCOM.ValidValue)uomComboBox.Selected).Value;
            }
        }

        private EditText EditText1;
        private EditText EditText2;
        private EditText EditText5;
        private EditText EditText6;
        private EditText EditText7;
        private EditText EditText8;
        private EditText EditText9;
        private EditText EditText10;
        private EditText EditText12;
        private EditText EditText13;
        private EditText EditText14;
        private UserDataSource oUserDataSource;
        private bool lostFocusQCLength = false;
        private bool lostFocusQCWidth = false;

        private void UnPrice_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (currentPrice != this.UnPrice.Value)
            {
                double newPrice = double.Parse(this.UnPrice.Value);
                UnPrice.Value = $"{newPrice.ToString("0.00")} {this.QCDocCur.Value}";

                if (Utils.MainCurrency != this.QCDocCur.Value)
                {
                    if (Utils.DirectRate == "Y")
                    {
                        double LCprice = double.Parse(Regex.Replace((string.IsNullOrEmpty(UnPrice.Value) ? "0" : UnPrice.Value), $@"[^\d{Utils.decSep}{Utils.thousSep}]", "")) * DocExRate;
                        LCPrice.Value = $"{LCprice.ToString("0.00")} {Utils.MainCurrency}";

                    }
                    else
                    {
                        double LCprice = double.Parse(Regex.Replace((string.IsNullOrEmpty(UnPrice.Value) ? "0" : UnPrice.Value), $@"[^\d{Utils.decSep}{Utils.thousSep}]", "")) / DocExRate;
                        LCPrice.Value = $"{LCprice.ToString("0.00")} {Utils.MainCurrency}";
                    }
                    LCCurr.Value = Utils.MainCurrency;
                    this.LCPrice.Item.Visible = true;
                    this.LCCurr.Item.Visible = true;

                    PictureBox0.Picture = QCEvents.LoadImageFromResources();
                    //if (LCPrice=0)
                    //{
                    //    PictureBox0.Picture = Properties.Resources.
                    //}
                    //else if (condition2)
                    //{
                    //    imagePath = "C:\\image3.png";
                    //}
                    //else
                    //{
                    //    imagePath = "C:\\image1.png"; // default image path
                    //}
                }
            }
        }

        private void QCLength_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (currentLength != QCLength.Value)
            {
                if (lostFocusQCLength)
                {
                    lostFocusQCLength = false;
                    return;
                }
                this.QCLength.Value = $"{this.QCLength.Value} {selectedUOM}";
                QCEvents.CalculateArea(this.UIAPIRawForm.UniqueID, selectedUOM);
                lostFocusQCLength = true;

            }

        }

        private void QCWidth_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (currentWidth != QCWidth.Value)
            {
                if (lostFocusQCWidth)
                {
                    lostFocusQCWidth = false;
                    return;
                }
                this.QCWidth.Value = $"{this.QCWidth.Value} {selectedUOM}";
                QCEvents.CalculateArea(this.UIAPIRawForm.UniqueID, selectedUOM);
                lostFocusQCWidth = true;
            }


        }

        private void QCWidth_GotFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            currentWidth = QCWidth.Value;

        }

        private void QCLength_GotFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            currentLength = QCLength.Value;

        }

        private StaticText StaticText5;
        private StaticText StaticText6;
        private StaticText StaticText7;
        private StaticText StaticText8;
        private StaticText StaticText9;
        private EditText EditText0;
        private StaticText StaticText10;
        private StaticText StaticText11;
        private StaticText StaticText12;
        private StaticText StaticText13;
        private StaticText StaticText15;
        private StaticText StaticText16;
        private StaticText StaticText17;
        private StaticText StaticText18;
        private SAPbouiCOM.PictureBox PictureBox0;
        private SAPbouiCOM.PictureBox ToolImg;
        private SAPbouiCOM.Button PicBrowse;
        private SAPbouiCOM.Button Button1;

        private async void PicBrowse_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            TaskCompletionSource<string> tcs = new TaskCompletionSource<string>();
            Thread t = new Thread(() => OpenImageFileDialog(tcs))
            {
                IsBackground = true,
            };

            t.SetApartmentState(ApartmentState.STA);
            t.Start();

            try
            {
                ToolfileName = await tcs.Task;

                // Get the SAP installation path
                string sapImageFolderPath = Path.GetDirectoryName(Utils.oCompany.BitMapPath);

                // Construct the default SAP image folder path
                sapToolsFolder = Path.Combine(sapImageFolderPath, "Tools Images");

                // Make sure the Images folder exists
                if (!Directory.Exists(sapToolsFolder))
                {
                    Directory.CreateDirectory(sapToolsFolder);
                }

                // Copy the selected image to the SAP image folder
                if (File.Exists(Path.Combine(sapToolsFolder, Path.GetFileName(ToolfileName))))
                {
                    newImageFilename = Path.Combine(sapToolsFolder, Path.GetFileName(ToolfileName));
                    string newImagePath = Path.Combine(sapToolsFolder, newImageFilename);
                    ToolImg.Picture = newImagePath;
                }
                else
                {
                    newImageFilename = "QCID" + QCDocEntry.Value + "_" + Path.GetFileNameWithoutExtension(Path.GetFileName(ToolfileName)) + Path.GetExtension(Path.GetFileName(ToolfileName));
                    string newImagePath = Path.Combine(sapToolsFolder, newImageFilename);
                    File.Copy(ToolfileName, newImagePath, true);
                    ToolImg.Picture = newImagePath;
                }
                this.UIAPIRawForm.Mode = BoFormMode.fm_UPDATE_MODE;


            }
            catch (OperationCanceledException)
            {
                // The user did not select an image, so do nothing.
            }
            catch (Exception ex)
            {
                Program.SBO_Application.MessageBox(string.Format("{0} Path: {1}", ex.Message, string.IsNullOrEmpty(sapToolsFolder) ? "Path is not defined" : sapToolsFolder));
            }
        }

        private void OpenImageFileDialog(TaskCompletionSource<string> tcs)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Filter = "Image Files(*.BMP;*.JPG;*.JPEG;*.PNG)|*.BMP;*.JPG;*.JPEG;*.PNG",
                Title = "Select an image",
            };

            DialogResult dr = openFileDialog.ShowDialog(new System.Windows.Forms.Form());
            if (dr == DialogResult.OK)
            {
                tcs.SetResult(openFileDialog.FileName);
            }
            else
            {
                tcs.SetCanceled();
            }
        }


        private void Button1_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            bool confirmDelete = Program.SBO_Application.MessageBox("Do you want to delete the image from Tools Image Folder?", 1, "Yes", "No") == 1;
            if (confirmDelete)
            {
                if (File.Exists(ToolImg.Picture))
                {
                    File.Delete(ToolImg.Picture);
                }
            }
            ToolImg.Picture = "";
            this.UIAPIRawForm.Mode = BoFormMode.fm_UPDATE_MODE;

        }

        private EditText QCPinfo1;
        private EditText QCPinfo2;
        private StaticText StaticText19;
        private StaticText StaticText20;
        private string sapToolsFolder;
        private string newImageFilename;
        private bool isChooseFromListTriggered;

        private void ButtonOk_PressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            ToolImg.Picture = Path.GetFileName(ToolImg.Picture);

        }


        private void ButtonOk_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (File.Exists(Path.Combine(!Directory.Exists(Path.Combine(Utils.oCompany.BitMapPath, "Tools Images")) ? Utils.oCompany.BitMapPath : Path.Combine(Utils.oCompany.BitMapPath, "Tools Images"), ToolImg.Picture)))
            {
                ToolImg.Picture = Path.Combine(!Directory.Exists(Path.Combine(Utils.oCompany.BitMapPath, "Tools Images")) ? Utils.oCompany.BitMapPath : Path.Combine(Utils.oCompany.BitMapPath, "Tools Images"), ToolImg.Picture);

            }
            else
            {
                ToolImg.Picture = "";
            }
         }


        private void ToolImg_DoubleClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (File.Exists(ToolImg.Picture))
            {
                Process.Start(ToolImg.Picture);
            }

        }

        private void ToolImg_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (File.Exists(ToolImg.Picture))
            {
                Process.Start(ToolImg.Picture);
            }
        }



        private void FOperations_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                this.UIAPIRawForm.Freeze(true);
                mOperations.AutoResizeColumns();
            }
            catch (Exception ex)
            {

                Program.SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Medium, false);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
            }

        }



        private void btnGetOP_PressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            Matrix matrix1 = (Matrix)this.UIAPIRawForm.Items.Item("mTextures").Specific;
            List<Dictionary<string, string>> matrix1Values = QCEvents.GetAllValuesFromMatrix1(matrix1);

        }

        private void QCPartType_ChooseFromListBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ChooseFromList oCfl = this.UIAPIRawForm.ChooseFromLists.Item("cflPartT");
                SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string strSQL = $"SELECT T0.\"ItemCode\", T0.\"ItemName\" as \"Part Name\" FROM OITM T0 WHERE T0.\"ItemCode\" like 'SPT-%' and T0.\"ItemCode\" like '%00'";
                oRS.DoQuery(strSQL);

                SAPbouiCOM.Conditions oCons = null;
                SAPbouiCOM.Condition oCon = null;

                oCfl.SetConditions(oCons);

                oCons = oCfl.GetConditions();

                if (oRS.RecordCount > 0)
                {
                    do
                    {
                        oCon = oCons.Add();
                        oCon.Alias = "ItemCode";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = oRS.Fields.Item("ItemCode").Value.ToString();

                        oRS.MoveNext();

                        if (!oRS.EoF)
                        {

                            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

                        }
                    } while (!oRS.EoF);

                }
                oCfl.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void QCSubPart_ChooseFromListBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {

                SAPbouiCOM.Conditions oCons = null;
                SAPbouiCOM.Condition oCon = null;

                SAPbouiCOM.ChooseFromList oCfl = this.UIAPIRawForm.ChooseFromLists.Item("cflSPart");
                SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)Utils.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string strSQL = $"SELECT OITM.\"ItemCode\", OITM.\"ItemName\" as \"Part Name\" FROM OITM WHERE left(OITM.\"ItemCode\",5) = left('{this.QCPartType.Value}',5) and OITM.\"ItemCode\" not like 'SPT-%00'";
                oRS.DoQuery(strSQL);

                oCons = null;
                oCon = null;

                oCfl.SetConditions(oCons);

                oCons = oCfl.GetConditions();

                if (oRS.RecordCount > 0)
                {
                    do
                    {
                        oCon = oCons.Add();
                        oCon.Alias = "ItemCode";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = oRS.Fields.Item("ItemCode").Value.ToString();
                        oRS.MoveNext();
                        if (!oRS.EoF)
                        {
                            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                        }
                    } while (!oRS.EoF);
                }
                oCfl.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void QCPartType_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            SBOChooseFromListEventArg chooseFromListEventArg = (SBOChooseFromListEventArg)pVal;
            string chooseFromListId = chooseFromListEventArg.ChooseFromListUID;
            SAPbouiCOM.ChooseFromList chooseFromList = this.UIAPIRawForm.ChooseFromLists.Item(chooseFromListId);

            // Get the selected item from the Choose From List
            SAPbouiCOM.DataTable selectedDataTable = chooseFromListEventArg.SelectedObjects;
            if (selectedDataTable != null && selectedDataTable.Rows.Count > 0)
            {
                string sptCode = selectedDataTable.GetValue("ItemCode", 0).ToString();
                parttDescr = selectedDataTable.GetValue("ItemName", 0).ToString();

                this.SPartDescr.Value = parttDescr;
            }
        }

        private void BtnGetOPC_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            lastBtnOpselection = pVal.PopUpIndicator;
            switch (pVal.PopUpIndicator)
            {
                case 0:
                    QCEvents.GetOperations(this.UIAPIRawForm);
                    break;
                case 1:
                    QCEvents.GetOperationsGrp(this.UIAPIRawForm);
                    break;
                default:
                    QCEvents.GetOperations(this.UIAPIRawForm);
                    break;
            }
        }

        private void BtnGetOPC_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            switch (lastBtnOpselection)
            {
                case 0:
                    QCEvents.GetOperations(this.UIAPIRawForm);
                    break;
                case 1:
                    QCEvents.GetOperationsGrp(this.UIAPIRawForm);
                    break;
                default:
                    QCEvents.GetOperations(this.UIAPIRawForm);
                    break;
            }
        }

        private void mOperations_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            //if (pVal.ColUID == "OPQtdT")
            //{
            //    this.UIAPIRawForm.Freeze(true);
            //    mOperations.FlushToDataSource();
            //    mOperations.LoadFromDataSource();
            //    this.UIAPIRawForm.Freeze(false);
            //    this.UIAPIRawForm.Refresh();
            //}

        }

        private void ButtonCancel_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            //selectedMatrixRow = 0;
            //QCEvents.operationsUpdate = false;
            //if (this.UIAPIRawForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && QCEvents.operationsUpdate == true)
            //{
            //    DBCalls.UpdateOperationsDB(QCEvents.operations, this.QCDocEntry);
            //}

        }

        private void ButtonOk_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            
            //if (this.UIAPIRawForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && QCEvents.operationsUpdate == true)
            //{
                
            //    QCEvents.operationsUpdate = false;
            //    SAPbouiCOM.DataTable mOperations = QCEvents.operations;
            //    string qCDocEntry = this.QCDocEntry.Value;

            //    // Convert the SAPbouiCOM.DataTable to a .NET DataTable object
            //    System.Data.DataTable mOperationsConverted = ConvertToDataTable(mOperations);

            //    Thread updateThread = new Thread(() => DBCalls.UpdateOperationsDB(mOperationsConverted, qCDocEntry));
            //    updateThread.Start();
            //}
        }
        private System.Data.DataTable ConvertToDataTable(SAPbouiCOM.DataTable sapDataTable)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            for (int i = 0; i < sapDataTable.Columns.Count; i++)
            {
                dt.Columns.Add(sapDataTable.Columns.Item(i).Name);
            }

            for (int i = 0; i < sapDataTable.Rows.Count; i++)
            {
                System.Data.DataRow newRow = dt.NewRow();
                for (int j = 0; j < sapDataTable.Columns.Count; j++)
                {
                    newRow[j] = sapDataTable.GetValue(j, i);
                }
                dt.Rows.Add(newRow);
            }

            return dt;
        }



        private void mOperations_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            QCEvents.lastClickedMatrixUID = pVal.ItemUID;

        }

        private void mOCosts_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            QCEvents.lastClickedMatrixUID = pVal.ItemUID;

        }

        private void mTextures_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            QCEvents.lastClickedMatrixUID = pVal.ItemUID;
            selectedMatrixRow = pVal.Row;

        }

        private void Form_UnloadAfter(SBOItemEventArg pVal)
        {
            QuoteCalculator.mtxMaxLineID = 0;
            selectedMatrixRow = 0;

            if (this.UIAPIRawForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && QCEvents.operationsUpdate == true)
            {

                QCEvents.operationsUpdate = false;
                SAPbouiCOM.DataTable mOperations = QCEvents.operations;
                string qCDocEntry = this.QCDocEntry.Value;

                // Convert the SAPbouiCOM.DataTable to a .NET DataTable object
                System.Data.DataTable mOperationsConverted = ConvertToDataTable(mOperations);

                Thread updateThread = new Thread(() => DBCalls.UpdateOperationsDB(mOperationsConverted, qCDocEntry));
                updateThread.Start();

            }
        }
    }
}
