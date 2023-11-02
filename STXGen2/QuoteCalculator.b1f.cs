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
using System.Xml;
using STXGen2.Properties;

namespace STXGen2
{
    [FormAttribute("STXGen2.QuoteCalculator", "QuoteCalculator.b1f")]
    class QuoteCalculator : UserFormBase
    {
        private const int FixedMatrixHeight = 110;
        bool loadingForm = true;
        private int selectedRow = SAPEvents.selectedRow;
        public static string selectedUOM { get; set; } = "";
        public static string oldLengthValue { get; set; } = "";
        public static string oldWidthValue { get; set; } = "";

        public static string previousUOM { get; set; } = "";
        public static string currentPrice { get; private set; } = "";
        public static double DocExRate { get; private set; } = 0;
        public static string currentLength { get; private set; } = "";
        public static string currentWidth { get; private set; } = "";
        public static string currentHeight { get; private set; } = "";

        public static QCResults UnloadResults { get; set; }

        public static string ToolfileName { get; private set; } = "";
        public static string parttDescr { get; set; } = "";
        public int lastBtnOpselection { get; private set; } = 0;
        public string oldPartType { get; private set; }
        public string previousLineTotal { get; private set; }
        public string newCost { get; private set; }
        public string previousQty { get; private set; }
        public string previousResc { get; private set; }
        public static bool recalcConfirm { get; set; }
        public bool lostFocusCovA { get; private set; }
        public string subparttDescr { get; private set; }
        public static string mOperatinsListXML { get; set; }

        private SAPbouiCOM.EditText QCDocEntry;
        private SAPbouiCOM.EditText BaseLine;
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
        private SAPbouiCOM.EditText QCHeight;

        private SAPbouiCOM.EditText UnPrice;
        private SAPbouiCOM.EditText LCPrice;
        private SAPbouiCOM.EditText QCDocCur;
        private SAPbouiCOM.EditText LCCurr;

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.StaticText StaticText4;

        private StaticText lLCPrice;
        private SAPbouiCOM.Button Button0;
        private bool cellchecked;


        private SAPbouiCOM.CheckBox DefBOM;
        private SAPbouiCOM.ComboBox OPFilter;
        //private EditText EditText3;
        //private LinkedButton LinkedButton0;
        //private StaticText StaticText9;
        //private EditText EditText4;

        private StaticText StaticText5;
        private StaticText StaticText6;
        private StaticText StaticText7;
        private StaticText StaticText8;

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

        private EditText QCOpA;
        private EditText QCOPTot;
        private EditText QCOTCost;
        private EditText QCTotalHF;
        private EditText QCTotalH;
        private EditText QCTotalSCF;
        private EditText QCTEst;
        private EditText QCTotalSC;
        private EditText EditText12;
        private EditText QCLeadTime;
        private EditText EditText14;
        private UserDataSource oUserDataSource;
        private bool lostFocusQCLength = false;
        private bool lostFocusQCWidth = false;
        private bool lostFocusQCHeight = false;
        private LinkedButton LinkedButton2;
        private StaticText StaticText19;
        private EditText EditText16;
        private bool formUpdateTrigger;
        private string sapDocEntry;
        private string sapObjType;
        private string sapDocLineNum;

        private EditText EditText1;
        private SAPbouiCOM.Button reCalc;


        public QuoteCalculator()
        {

        }
        public class QCResults
        {
            public string QCID { get; set; }
            public string QCLineN { get; set; }
            public string QCuPrice { get; set; }
            public string QCcPrice { get; set; }
            public string QClTime { get; set; }
            public string QCptNum { get; set; }
            public string QCtNum { get; set; }
            public string QCprtName { get; set; }

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
            this.ButtonOk.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.ButtonOk_PressedAfter);
            this.ButtonOk.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.ButtonOk_PressedBefore);
            this.ButtonCancel = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.FOperations = ((SAPbouiCOM.Folder)(this.GetItem("FOper").Specific));
            this.FOperations.ClickAfter += new SAPbouiCOM._IFolderEvents_ClickAfterEventHandler(this.FOperations_ClickAfter);
            this.FOperations.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.FOperations_PressedAfter);
            this.FGeneral = ((SAPbouiCOM.Folder)(this.GetItem("FGen").Specific));
            this.FGeneral.ClickAfter += new SAPbouiCOM._IFolderEvents_ClickAfterEventHandler(this.FGeneral_ClickAfter);
            this.mTextures = ((SAPbouiCOM.Matrix)(this.GetItem("mTextures").Specific));
            this.mTextures.LostFocusAfter += new SAPbouiCOM._IMatrixEvents_LostFocusAfterEventHandler(this.mTextures_LostFocusAfter);
            this.mTextures.ClickBefore += new SAPbouiCOM._IMatrixEvents_ClickBeforeEventHandler(this.mTextures_ClickBefore);
            this.mTextures.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.mTextures_ClickAfter);
            this.mTextures.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.mTextures_ChooseFromListAfter);
            this.mOCosts = ((SAPbouiCOM.Matrix)(this.GetItem("mOCosts").Specific));
            this.mOCosts.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.mOCosts_ClickAfter);
            this.QCPartDesc = ((SAPbouiCOM.EditText)(this.GetItem("QCPartDesc").Specific));
            this.QCPartType = ((SAPbouiCOM.EditText)(this.GetItem("QCPartType").Specific));
            this.QCPartType.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.QCPartType_LostFocusAfter);
            this.QCPartType.GotFocusAfter += new SAPbouiCOM._IEditTextEvents_GotFocusAfterEventHandler(this.QCPartType_GotFocusAfter);
            this.QCPartType.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.QCPartType_ChooseFromListAfter);
            this.QCPartType.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.QCPartType_ChooseFromListBefore);
            this.QCSubPart = ((SAPbouiCOM.EditText)(this.GetItem("QCSubPart").Specific));
            this.QCSubPart.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.QCSubPart_ChooseFromListAfter);
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
            this.QCHeight = ((SAPbouiCOM.EditText)(this.GetItem("QCHeight").Specific));
            this.QCHeight.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.QCHeight_LostFocusAfter);
            this.QCHeight.GotFocusAfter += new SAPbouiCOM._IEditTextEvents_GotFocusAfterEventHandler(this.QCHeight_GotFocusAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("lDocEntry").Specific));
            this.QCDocEntry = ((SAPbouiCOM.EditText)(this.GetItem("QCDocEntry").Specific));
            this.QCItemCode = ((SAPbouiCOM.EditText)(this.GetItem("QCItemCode").Specific));
            this.QCItemName = ((SAPbouiCOM.EditText)(this.GetItem("QCItemN").Specific));
            this.QCOpA = ((SAPbouiCOM.EditText)(this.GetItem("QCOpA").Specific));
            this.QCOPTot = ((SAPbouiCOM.EditText)(this.GetItem("QCOPTot").Specific));
            this.QCOTCost = ((SAPbouiCOM.EditText)(this.GetItem("QCOTCost").Specific));
            this.QCTEst = ((SAPbouiCOM.EditText)(this.GetItem("QCTEst").Specific));
            this.QCTotalHF = ((SAPbouiCOM.EditText)(this.GetItem("QCTotalHF").Specific));
            this.QCTotalH = ((SAPbouiCOM.EditText)(this.GetItem("QCTotalH").Specific));
            this.QCTotalSCF = ((SAPbouiCOM.EditText)(this.GetItem("QCTotalSCF").Specific));
            this.QCTotalSC = ((SAPbouiCOM.EditText)(this.GetItem("QCTotalSC").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("Item_17").Specific));
            this.QCLeadTime = ((SAPbouiCOM.EditText)(this.GetItem("QCLeadTime").Specific));
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
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("lUnPrice").Specific));
            this.lLCPrice = ((SAPbouiCOM.StaticText)(this.GetItem("lLCPrice").Specific));
            this.lLCPrice.Item.Visible = false;
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
            this.lPinfo1 = ((SAPbouiCOM.StaticText)(this.GetItem("lPinfo1").Specific));
            this.lPinfo2 = ((SAPbouiCOM.StaticText)(this.GetItem("lPinfo2").Specific));
            this.mOperations = ((SAPbouiCOM.Matrix)(this.GetItem("mOper").Specific));
            this.mOperations.DoubleClickAfter += new SAPbouiCOM._IMatrixEvents_DoubleClickAfterEventHandler(this.mOperations_DoubleClickAfter);
            this.mOperations.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.mOperations_ChooseFromListAfter);
            this.mOperations.GotFocusAfter += new SAPbouiCOM._IMatrixEvents_GotFocusAfterEventHandler(this.mOperations_GotFocusAfter);
            this.mOperations.LostFocusAfter += new SAPbouiCOM._IMatrixEvents_LostFocusAfterEventHandler(this.mOperations_LostFocusAfter);
            this.mOperations.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.mOperations_ClickAfter);
            this.BtnGetOPC = ((SAPbouiCOM.ButtonCombo)(this.GetItem("btnGetOp").Specific));
            this.BtnGetOPC.PressedAfter += new SAPbouiCOM._IButtonComboEvents_PressedAfterEventHandler(this.BtnGetOPC_PressedAfter);
            this.BtnGetOPC.ComboSelectAfter += new SAPbouiCOM._IButtonComboEvents_ComboSelectAfterEventHandler(this.BtnGetOPC_ComboSelectAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("OpRem").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.DefBOM = ((SAPbouiCOM.CheckBox)(this.GetItem("DefBOM").Specific));
            this.DefBOM.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.DefBOM_PressedAfter);
            this.OPFilter = ((SAPbouiCOM.ComboBox)(this.GetItem("OPFilter").Specific));
            this.OPFilter.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.OPFilter_ComboSelectAfter);
            this.LinkedButton2 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_21").Specific));
            this.StaticText19 = ((SAPbouiCOM.StaticText)(this.GetItem("lWOrder").Specific));
            this.EditText16 = ((SAPbouiCOM.EditText)(this.GetItem("Item_27").Specific));
            this.BaseLine = ((SAPbouiCOM.EditText)(this.GetItem("BaseLine").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("QCWOrder").Specific));
            this.reCalc = ((SAPbouiCOM.Button)(this.GetItem("Recalc").Specific));
            this.OnCustomInitialize();

        }


        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            this.UnloadBefore += new SAPbouiCOM.Framework.FormBase.UnloadBeforeHandler(this.Form_UnloadBefore);

        }


        private void OnCustomInitialize()
        {
            QCEvents.FillTextureClass(this.UIAPIRawForm);
            QCEvents.FillUnitMeasures(this.UIAPIRawForm);
        }


        internal string AddUDO(string docEntry, string objType, string mLinenum)
        {
            BoObjectTypes objectType = DBCalls.GetSAPObjectType(objType);
            SAPbobsCOM.Documents doc = (SAPbobsCOM.Documents)Utils.oCompany.GetBusinessObject(objectType);

            SAPbobsCOM.CompanyService oCompanyService = Utils.oCompany.GetCompanyService();
            SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("STXQC19");
            SAPbobsCOM.GeneralData oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
            SAPbobsCOM.GeneralDataParams oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

            #region Add new QcID

            oGeneralData.SetProperty("U_bsObjType", doc.DocObjectCodeEx);
            oGeneralData.SetProperty("U_bsDocEntry", docEntry);
            oGeneralData.SetProperty("U_bsLineNum", mLinenum);

            var response = oGeneralService.Add(oGeneralData);
            string DocEntry = response.GetProperty("DocEntry").ToString();

            #endregion 
            return DocEntry;
        }

        internal void LoadFrmByKey(string qcid, string itemCode, string itemName, string docCur, string unPrice, float exRate,string DocEntry, string ObjType, string mLinenum)
        {
            try
            {
                sapDocEntry = DocEntry;
                sapObjType = ObjType;
                sapDocLineNum = mLinenum;

                this.UIAPIRawForm.Freeze(true);
                formUpdateTrigger = false;
                SetFormModeToFind();
                LoadDocumentAndBindMatrix(qcid);
                ParseExchangeRate(exRate);
                SetFieldValues(itemCode, itemName);
                SetButtonValidValues();
                BindFieldsAndCalculateArea(docCur, unPrice);
                AddRowIfMatrixEmpty();
                MatrixSorting();
                DisableFormWO();
                DisableGTOperCC1(itemCode);
                
                this.Show();
            }
            catch (Exception)
            {
                // Log error here and show a user-friendly message
                System.Windows.Forms.MessageBox.Show("An error occurred while loading the form. Please try again.");
            }
            finally
            {
                SetFinalFormProperties();
                this.UIAPIRawForm.Freeze(false);
            }
        }

        private void DisableGTOperCC1(string itemCode)
        {
            string tech1 = DBCalls.GetItemTech(itemCode);
            if (tech1 != "MT" && tech1 != "LA" && tech1 != "CT")
            {
                this.BtnGetOPC.Item.Enabled = false;
            }

        }

        private void DisableFormWO()
        {
            if (!string.IsNullOrEmpty(EditText1.Value))
            {
                this.mTextures.Item.Enabled = false;
                this.mOperations.Item.Enabled = false;
                this.QCPartType.Item.Enabled = false;
                this.QCSubPart.Item.Enabled = false;
                this.BtnGetOPC.Item.Enabled = false;
                this.DefBOM.Item.Enabled = false;
                this.Button0.Item.Enabled = false;
                this.reCalc.Item.Enabled = false;
            }
        }

        private void MatrixSorting()
        {
            mTextures.Columns.Item("#").TitleObject.Sortable = true;
            mTextures.Columns.Item("#").TitleObject.Sort(BoGridSortType.gst_Ascending);
            mTextures.Columns.Item("#").TitleObject.Sortable = false;

            mOperations.Columns.Item("#").TitleObject.Sortable = true;
            mOperations.Columns.Item("#").TitleObject.Sort(BoGridSortType.gst_Ascending);
            mOperations.Columns.Item("#").TitleObject.Sortable = false;
        }

        private void LoadDocumentAndBindMatrix(string qcid)
        {
            // Enable QCDocEntry field temporarily
            QCDocEntry.Item.Enabled = true;
            QCDocEntry.Value = qcid;

            QCEvents.BindMatrixCheckboxes(this.UIAPIRawForm, mOperations, mOperations.RowCount);
            ButtonOk.Item.Click();

            FGeneral.Select();
        }

        private void SetFormModeToFind()
        {
            this.UIAPIRawForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        }

        private void ParseExchangeRate(float exRate)
        {
            System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();
            DocExRate = double.Parse(exRate.ToString("F"), NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, sapNumberFormat);
        }

        private void SetFieldValues(string itemCode, string itemName)
        {
            QCItemCode.Value = itemCode;
            QCItemName.Value = itemName;

            ToolImg.Picture = Path.Combine(!Directory.Exists(Path.Combine(Utils.oCompany.BitMapPath, "Tools Images")) ? Utils.oCompany.BitMapPath : Path.Combine(Utils.oCompany.BitMapPath, "Tools Images"), ToolImg.Picture);
        }

        private void SetButtonValidValues()
        {
            BtnGetOPC.ValidValues.Add("1", "Get Operations");
            BtnGetOPC.ValidValues.Add("2", "Get Operations (Grouped)");
        }

        private void BindFieldsAndCalculateArea(string docCur, string unPrice)
        {
            selectedUOM = UnMsr.Selected.Value;
            previousUOM = UnMsr.Selected.Value;

            System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();

            SAPbouiCOM.EditText DocCur = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("QCDocCur").Specific;
            DocCur.Value = docCur;

            oUserDataSource = this.UIAPIRawForm.DataSources.UserDataSources.Add("MyUNPrice", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);

            SAPbouiCOM.EditText UNPrice = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("UnPrice").Specific;
            UNPrice.DataBind.SetBound(true, "", "MyUNPrice");
            oUserDataSource.Value = unPrice;

            double LCprice = HelperMethods.ParseDoubleWCur(unPrice, sapNumberFormat);

            if (Utils.MainCurrency != DocCur.Value)
            {
                if (Utils.DirectRate == "Y")
                {
                    LCPrice.Value = $"{(LCprice * DocExRate).ToString("#,0.00", sapNumberFormat)} {Utils.MainCurrency}";
                }
                else
                {
                    LCPrice.Value = $"{(LCprice / DocExRate).ToString("#,0.00", sapNumberFormat)} {Utils.MainCurrency}";
                }
                LCCurr.Value = Utils.MainCurrency;
                this.LCPrice.Item.Visible = true;
                this.LCCurr.Item.Visible = true;
                this.lLCPrice.Item.Visible = true;
            }

            currentPrice = this.UnPrice.Value;

            QCEvents.GetSubPartType(UIAPIRawForm, this.QCSubPart);
            QCEvents.CalculateArea(this.UIAPIRawForm.UniqueID, selectedUOM);
            QCEvents.GetFiltersOperations(this.UIAPIRawForm, this.QCDocEntry);
            FormDataRecalculation();

            PictureBox0.Picture = QCEvents.SellMarginImage(this.UIAPIRawForm);
        }

        private void AddRowIfMatrixEmpty()
        {
            if (mTextures.RowCount == 0 || !string.IsNullOrWhiteSpace(((SAPbouiCOM.EditText)mTextures.Columns.Item("QCTexture").Cells.Item(mTextures.RowCount).Specific).Value))
            {
                mTextures.AddRow();
                mTextures.ClearRowData(mTextures.RowCount);
                SAPbouiCOM.EditText newAutoRow = (SAPbouiCOM.EditText)mTextures.Columns.Item("#").Cells.Item(mTextures.RowCount).Specific;
                newAutoRow.Value = mTextures.RowCount.ToString();
            }

            if (mOperations.RowCount == 0)
            {
                mOperations.AddRow();
                SAPbouiCOM.EditText newAutoRow = (SAPbouiCOM.EditText)mOperations.Columns.Item("#").Cells.Item(1).Specific;
                newAutoRow.Value = "1";


            }
        }

        private void SetFinalFormProperties()
        {
            QCToolNum.Item.Click();

            QCDocEntry.Item.Enabled = false;
            QCItemCode.Item.Enabled = false;
            QCItemName.Item.Enabled = false;
            mTextures.AutoResizeColumns();
            loadingForm = false;
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
                        mtxTextures.SetCellWithoutValidation(SAPEvents.selectedRow, "QCQuantity", "1");
                        mtxTextures.SetCellWithoutValidation(SAPEvents.selectedRow, "QCCovA", QCArea.Value);
                        ((SAPbouiCOM.ComboBox)mtxTextures.Columns.Item("QCTClass").Cells.Item(SAPEvents.selectedRow).Specific).Select(TClass, BoSearchKey.psk_ByValue);
                        ((SAPbouiCOM.ComboBox)mtxTextures.Columns.Item("QCGComp").Cells.Item(SAPEvents.selectedRow).Specific).Select("2", BoSearchKey.psk_ByValue);
                        ((SAPbouiCOM.EditText)mtxTextures.Columns.Item("QCTexture").Cells.Item(SAPEvents.selectedRow).Specific).Value = TextureCode;
                    }
                }

            }
            catch (Exception ex)
            {
                // Log or display the exception message
                Program.SBO_Application.MessageBox("Error: " + ex.Message);
            }
        }



        private void UnMsr_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {

            bool isUoMAreaChanging = false;

            if (pVal.ItemChanged == true && loadingForm == false)
            {
                try
                {
                    this.UIAPIRawForm.Freeze(true);
                    System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();

                    // Get the selected Unit of Measure
                    SAPbouiCOM.ComboBox uomComboBox = (SAPbouiCOM.ComboBox)this.UIAPIRawForm.Items.Item("UnMsr").Specific;
                    selectedUOM = uomComboBox.Selected.Value;

                    // Get the current Length and Width values
                    EditText edtLength = (EditText)this.UIAPIRawForm.Items.Item("QCLength").Specific;
                    EditText edtWidth = (EditText)this.UIAPIRawForm.Items.Item("QCWidth").Specific;
                    EditText edtHeight = (EditText)this.UIAPIRawForm.Items.Item("QCHeight").Specific;

                    double length = HelperMethods.ParseDoubleWUOM(edtLength.Value, sapNumberFormat); 
                    double width = HelperMethods.ParseDoubleWUOM(edtWidth.Value, sapNumberFormat); 
                    double height = HelperMethods.ParseDoubleWUOM(edtHeight.Value, sapNumberFormat); 

                    // Perform the conversion based on the selected Unit of Measure
                    double convertedLength = DBCalls.ConvertDimensions(length, selectedUOM, previousUOM);
                    double convertedWidth = DBCalls.ConvertDimensions(width, selectedUOM, previousUOM);
                    double convertedHeight = DBCalls.ConvertDimensions(height, selectedUOM, previousUOM);

                    // Update the Length and Width fields with the converted values
                    edtLength.Value = $"{Math.Round(convertedLength, Utils.MeasureDec).ToString("N", sapNumberFormat)}";
                    edtWidth.Value = $"{Math.Round(convertedWidth, Utils.MeasureDec).ToString("N", sapNumberFormat)}";
                    edtHeight.Value = $"{Math.Round(convertedHeight, Utils.MeasureDec).ToString("N", sapNumberFormat)}";

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
                    uomComboBox.Active = true;
                }
                finally
                {
                    this.UIAPIRawForm.Freeze(false);
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




        private void UnPrice_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();

            if (currentPrice != this.UnPrice.Value)
            {
                double newPrice = HelperMethods.ParseDoubleWCur(this.UnPrice.Value,sapNumberFormat);
                UnPrice.Value = HelperMethods.FormatValueCur(newPrice, this.QCDocCur.Value);

                if (Utils.MainCurrency != this.QCDocCur.Value)
                {
                    double LCprice = Utils.DirectRate == "Y" ? newPrice * DocExRate : newPrice / DocExRate;
                    LCPrice.Value = HelperMethods.FormatValueCur(LCprice, Utils.MainCurrency);

                    LCCurr.Value = Utils.MainCurrency;
                    this.LCPrice.Item.Visible = true;
                    this.LCCurr.Item.Visible = true;

                    
                }
            }
            PictureBox0.Picture = QCEvents.SellMarginImage(this.UIAPIRawForm);
        }

        private void QCLength_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();
            double qcLength = 0;

            if (currentLength != QCLength.Value)
            {
                if (lostFocusQCLength)
                {
                    lostFocusQCLength = false;
                    return;
                }
                try
                {
                    qcLength = double.Parse(this.QCLength.Value, System.Globalization.NumberStyles.AllowDecimalPoint | System.Globalization.NumberStyles.AllowThousands, sapNumberFormat);

                }
                catch (Exception)
                {
                    qcLength = 0;
                    Program.SBO_Application.SetStatusBarMessage("Please, place a numeric value.", BoMessageTime.bmt_Short, true);
                }
                string formattedQCLength = qcLength.ToString("#,0.00", sapNumberFormat);

                this.QCLength.Value = $"{formattedQCLength} {selectedUOM}";
                QCEvents.CalculateArea(this.UIAPIRawForm.UniqueID, selectedUOM);
                lostFocusQCLength = true;
            }

        }

        private void QCWidth_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();
            double qcWidth = 0;

            if (currentWidth != QCWidth.Value)
            {
                if (lostFocusQCWidth)
                {
                    lostFocusQCWidth = false;
                    return;
                }
                try
                {
                    qcWidth = double.Parse(this.QCWidth.Value, System.Globalization.NumberStyles.AllowDecimalPoint | System.Globalization.NumberStyles.AllowThousands, sapNumberFormat);
                }
                catch (Exception)
                {
                    qcWidth = 0;
                    Program.SBO_Application.SetStatusBarMessage("Please, place a numeric value.", BoMessageTime.bmt_Short, true);
                }

                string formattedQCWidth = qcWidth.ToString("#,0.00", sapNumberFormat);

                this.QCWidth.Value = $"{formattedQCWidth} {selectedUOM}";
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
        private StaticText lPinfo1;
        private StaticText lPinfo2;
        private string sapToolsFolder;
        private string newImageFilename;
        private bool isChooseFromListTriggered;

        private void ButtonOk_PressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.FormMode == 2)
            {
                for (int i = 1; i <= mOperations.RowCount; i++)
                {

                    SAPbouiCOM.EditText cellvalue = (SAPbouiCOM.EditText)mOperations.Columns.Item("OPErrMsg").Cells.Item(i).Specific;
                    if (!string.IsNullOrEmpty(cellvalue.Value))
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Please correct all the errors on the operations matrix before proceding...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false; // This will prevent the event from propagating further
                        return;
                    }
                }
                formUpdateTrigger = true;

            }

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
                // string strSQL = $"SELECT OITM.\"ItemCode\", OITM.\"ItemName\" as \"Part Name\" FROM OITM WHERE left(OITM.\"ItemCode\",6) = left('{this.QCPartType.Value}',6) and OITM.\"ItemCode\" not like 'SPT-%00'";
                string strSQL = string.IsNullOrEmpty(this.QCPartType.Value)
                 ? $"SELECT OITM.\"ItemCode\", OITM.\"ItemName\" as \"Part Name\" FROM OITM WHERE OITM.\"ItemCode\" like 'SPT-%' and right(OITM.\"ItemCode\",2) != '00'"
                 : $"SELECT OITM.\"ItemCode\", OITM.\"ItemName\" as \"Part Name\" FROM OITM WHERE left(OITM.\"ItemCode\",6) = left('{this.QCPartType.Value}',6) and OITM.\"ItemCode\" like 'SPT-%' and right(OITM.\"ItemCode\",2) != '00'";

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
            string sptCode = "";
            SBOChooseFromListEventArg chooseFromListEventArg = (SBOChooseFromListEventArg)pVal;
            string chooseFromListId = chooseFromListEventArg.ChooseFromListUID;
            SAPbouiCOM.ChooseFromList chooseFromList = this.UIAPIRawForm.ChooseFromLists.Item(chooseFromListId);

            // Get the selected item from the Choose From List
            SAPbouiCOM.DataTable selectedDataTable = chooseFromListEventArg.SelectedObjects;
            if (selectedDataTable != null && selectedDataTable.Rows.Count > 0)
            {
                sptCode = selectedDataTable.GetValue("ItemCode", 0).ToString();
                parttDescr = selectedDataTable.GetValue("ItemName", 0).ToString();

                //this.SPartDescr.Value = parttDescr;
            }
        }

        private void BtnGetOPC_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            lastBtnOpselection = pVal.PopUpIndicator;
            int noperations = mOperations.RowCount;
            switch (pVal.PopUpIndicator)
            {
                case 0:
                    if (noperations > 0)
                    {
                        bool confirmGetOper = Program.SBO_Application.MessageBox("This operation will clear the current operations. Do you want to continue?", 1, "Yes", "No") == 1;
                        if (confirmGetOper)
                        {
                            if (this.DefBOM.Checked == true)
                            {
                                DefBOM.Checked = false;
                            }

                            this.mOperations.Clear();
                            QCEvents.GetOperations(this.UIAPIRawForm, selectedRow);
                            QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                        }
                    }
                    else
                    {
                        if (this.DefBOM.Checked == true)
                        {
                            DefBOM.Checked = false;
                        }
                        this.mOperations.Clear();
                        QCEvents.GetOperations(this.UIAPIRawForm, selectedRow);
                        QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                    }
                    break;
                case 1:
                    if (noperations > 0)
                    {
                        bool confirmGetOper = Program.SBO_Application.MessageBox("This operation will clear the current operations. Do you want to continue?", 1, "Yes", "No") == 1;
                        if (confirmGetOper)
                        {
                            if (this.DefBOM.Checked == true)
                            {
                                DefBOM.Checked = false;
                            }
                            this.mOperations.Clear();
                            QCEvents.GetOperationsGrp(this.UIAPIRawForm);
                            QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                        }
                    }
                    else
                    {
                        if (this.DefBOM.Checked == true)
                        {
                            DefBOM.Checked = false;
                        }
                        this.mOperations.Clear();
                        QCEvents.GetOperationsGrp(this.UIAPIRawForm);
                        QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                    }
                    break;
                default:
                    if (noperations > 0)
                    {
                        bool confirmGetOper = Program.SBO_Application.MessageBox("This operation will clear the current operations. Do you want to continue?", 1, "Yes", "No") == 1;
                        if (confirmGetOper)
                        {
                            if (this.DefBOM.Checked == true)
                            {
                                DefBOM.Checked = false;
                            }
                            this.mOperations.Clear();
                            QCEvents.GetOperations(this.UIAPIRawForm, selectedRow);
                            QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                        }
                    }
                    else
                    {
                        if (this.DefBOM.Checked == true)
                        {
                            DefBOM.Checked = false;
                        }
                        this.mOperations.Clear();
                        QCEvents.GetOperations(this.UIAPIRawForm, selectedRow);
                        QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                    }
                    break;
            }
        }

        private void BtnGetOPC_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                this.UIAPIRawForm.Freeze(true);
                int noperations = mOperations.RowCount;
                switch (lastBtnOpselection)
                {
                    case 0:
                        if (noperations > 0)
                        {
                            bool confirmGetOper = Program.SBO_Application.MessageBox("This operation will clear the current operations. Do you want to continue?", 1, "Yes", "No") == 1;
                            if (confirmGetOper)
                            {
                                if (this.DefBOM.Checked == true)
                                {
                                    DefBOM.Checked = false;
                                }

                                this.mOperations.Clear();
                                QCEvents.GetOperations(this.UIAPIRawForm, selectedRow);
                                QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                            }
                        }
                        else
                        {
                            if (this.DefBOM.Checked == true)
                            {
                                DefBOM.Checked = false;
                            }
                            this.mOperations.Clear();
                            QCEvents.GetOperations(this.UIAPIRawForm, selectedRow);
                            QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                        }
                        break;
                    case 1:
                        if (noperations > 0)
                        {
                            bool confirmGetOper = Program.SBO_Application.MessageBox("This operation will clear the current operations. Do you want to continue?", 1, "Yes", "No") == 1;
                            if (confirmGetOper)
                            {
                                if (this.DefBOM.Checked == true)
                                {
                                    DefBOM.Checked = false;
                                }
                                this.mOperations.Clear();
                                QCEvents.GetOperationsGrp(this.UIAPIRawForm);
                                QCEvents.OperationsCalcTotal(this.UIAPIRawForm);

                            }
                        }
                        else
                        {
                            if (this.DefBOM.Checked == true)
                            {
                                DefBOM.Checked = false;
                            }
                            this.mOperations.Clear();
                            QCEvents.GetOperationsGrp(this.UIAPIRawForm);
                            QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                        }
                        break;
                    default:
                        if (noperations > 0)
                        {
                            bool confirmGetOper = Program.SBO_Application.MessageBox("This operation will clear the current operations. Do you want to continue?", 1, "Yes", "No") == 1;
                            if (confirmGetOper)
                            {
                                if (this.DefBOM.Checked == true)
                                {
                                    DefBOM.Checked = false;
                                }
                                this.mOperations.Clear();
                                QCEvents.GetOperations(this.UIAPIRawForm, selectedRow);
                                QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                            }
                        }
                        else
                        {
                            if (this.DefBOM.Checked == true)
                            {
                                DefBOM.Checked = false;
                            }
                            this.mOperations.Clear();
                            QCEvents.GetOperations(this.UIAPIRawForm, selectedRow);
                            QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                            
                        }
                        break;
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                PictureBox0.Picture = QCEvents.SellMarginImage(this.UIAPIRawForm);
                this.UIAPIRawForm.Freeze(false);
            }
          
        }

        private void mOperations_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPEvents.lastClickedMatrixUID = pVal.ItemUID;

        }

        private void mOCosts_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPEvents.lastClickedMatrixUID = pVal.ItemUID;

        }

        private void mTextures_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPEvents.lastClickedMatrixUID = pVal.ItemUID;
            SAPEvents.selectedRow = pVal.Row;

        }


        private void mTextures_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            mTextures.FlushToDataSource();

        }



        private void FOperations_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            resizeOperationsFormUI();

        }
        private void FGeneral_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            resizeTextureFormUI();

        }


        #region Form Scale Resize

        private void Form_ResizeAfter(SBOItemEventArg pVal)
        {

            Item tabControlItem = this.UIAPIRawForm.Items.Item("Item_9");
            Item docEntryItem = this.UIAPIRawForm.Items.Item("QCDocEntry");

            Folder operationsFolder = (Folder)this.UIAPIRawForm.Items.Item("FOper").Specific;
            Folder texturesFolder = (Folder)this.UIAPIRawForm.Items.Item("FGen").Specific;

            SAPbouiCOM.Item obsText = (SAPbouiCOM.Item)this.UIAPIRawForm.Items.Item("QCObs");

            Item buttonOKItem = this.UIAPIRawForm.Items.Item("1");
            Item buttonCancelItem = this.UIAPIRawForm.Items.Item("2");

            //Resize the tab control while maintaining the minimum width and height
            tabControlItem.Width = this.UIAPIRawForm.ClientWidth - tabControlItem.Left - 5;
            if (texturesFolder.Selected)
            {
                resizeTextureFormUI();

            }
            else if (operationsFolder.Selected)
            {
                resizeOperationsFormUI();
            }

            int marginBetweenObsTextAndButtons = 10; // Adjust this value for your preferred spacing

            buttonOKItem.Top = obsText.Top + obsText.Height + marginBetweenObsTextAndButtons;
            buttonCancelItem.Top = obsText.Top + obsText.Height + marginBetweenObsTextAndButtons;

            int totalHeightNeeded = buttonOKItem.Top + buttonOKItem.Height + 10;  // adding a margin of 10
            if (this.UIAPIRawForm.ClientHeight < totalHeightNeeded)
            {
                this.UIAPIRawForm.ClientHeight = totalHeightNeeded;
            }

            mOperations.AutoResizeColumns();
            mTextures.AutoResizeColumns();
            mOCosts.AutoResizeColumns();
            this.UIAPIRawForm.Refresh();


        }

        private void resizeTextureFormUI()
        {
            this.UIAPIRawForm.Freeze(true);
            Item tabControlItem = this.UIAPIRawForm.Items.Item("Item_9");
            Item matrix1Item = this.UIAPIRawForm.Items.Item("mTextures");
            Item matrix2Item = this.UIAPIRawForm.Items.Item("mOCosts");

            int availableWidth = tabControlItem.Width - matrix1Item.Left - 5;

            int minHeight = matrix2Item.Top + 20;

            tabControlItem.Height = minHeight;

            // Set a fixed height for 5 rows (assuming a row height of 21 pixels)
            int fixedMatrixHeight = 21 * 5;

            matrix1Item.Width = tabControlItem.Width - 20;
            matrix1Item.Height = fixedMatrixHeight;

            QCPinfo1.Item.Top = matrix1Item.Top + matrix1Item.Height + 10;
            QCPinfo1.Item.Width = availableWidth / 2;
            lPinfo1.Item.Top = matrix1Item.Top + matrix1Item.Height + 27;

            QCPinfo2.Item.Top = QCPinfo1.Item.Top + QCPinfo1.Item.Height + 5;
            QCPinfo2.Item.Width = availableWidth / 2;
            lPinfo2.Item.Top = QCPinfo1.Item.Top + QCPinfo1.Item.Height + 22;

            matrix2Item.Top = QCPinfo2.Item.Top + QCPinfo2.Item.Height + 10;  //matrix1Item.Top + matrix1Item.Height + 20; // Add an additional space between the matrices
            matrix2Item.Width = availableWidth / 2;
            matrix2Item.Height = fixedMatrixHeight;

            SAPbouiCOM.Item obsText = this.UIAPIRawForm.Items.Item("QCObs");
            obsText.Top = this.UIAPIRawForm.Height - obsText.Height - 80; // 50 pixel gap from bottom of form

            int minDistanceFromTab = 30;
            if (obsText.Top < tabControlItem.Top + tabControlItem.Height + minDistanceFromTab)
            {
                obsText.Top = tabControlItem.Top + tabControlItem.Height + minDistanceFromTab;
            }

            SAPbouiCOM.Item cancelButton = this.UIAPIRawForm.Items.Item("2"); // Replace "OkButton" with the actual ID of your Ok button
            if (obsText.Top + obsText.Height > cancelButton.Top)
            {
                obsText.Top = cancelButton.Top - obsText.Height - 10; // Keeps a 10 pixel gap from the Ok button
            }

            this.UIAPIRawForm.Freeze(false);
        }

        private void resizeOperationsFormUI()
        {
            this.UIAPIRawForm.Freeze(true);
            Item tabControlItem = this.UIAPIRawForm.Items.Item("Item_9");
            Item matrix1Item = this.UIAPIRawForm.Items.Item("mOper");
            Item remarkBoxItem = this.UIAPIRawForm.Items.Item("QCObs");
            Item remOperationsBtn = this.UIAPIRawForm.Items.Item("OpRem");

            // Resize the tab control while maintaining the minimum width and height
            tabControlItem.Width = this.UIAPIRawForm.ClientWidth - tabControlItem.Left - 5;

            int remarkBoxHeight = remarkBoxItem.Height;

            int availableHeight = this.UIAPIRawForm.ClientHeight - tabControlItem.Top - remarkBoxHeight;
            tabControlItem.Height = (int)(availableHeight * 0.90);

            // Get tab control dimensions
            int tabControlWidth = tabControlItem.Width;
            int tabControlHeight = tabControlItem.Height;

            // Set the matrix width and height based on the tab control dimensions
            matrix1Item.Width = tabControlWidth - 15; // Adjust the value as needed to fit within the tab control
            matrix1Item.Height = tabControlHeight - 80;

            // Calculate the top position of the matrix inside the tab control
            int matrixTopPosition = matrix1Item.Top - tabControlItem.Top;
            int matrixHeight = matrix1Item.Height;

            // Calculate the new top position for the OpRem button based on the matrix height and desired spacing
            remOperationsBtn.Top = tabControlItem.Top + tabControlHeight - 25;

            SAPbouiCOM.Item obsText = this.UIAPIRawForm.Items.Item("QCObs");
            obsText.Top = this.UIAPIRawForm.Height - obsText.Height - 80; // 50 pixel gap from bottom of form

            int minDistanceFromTab = 30;
            if (obsText.Top < tabControlItem.Top + tabControlItem.Height + minDistanceFromTab)
            {
                obsText.Top = tabControlItem.Top + tabControlItem.Height + minDistanceFromTab;
            }

            SAPbouiCOM.Item cancelButton = this.UIAPIRawForm.Items.Item("2"); // Replace "OkButton" with the actual ID of your Ok button
            if (obsText.Top + obsText.Height > cancelButton.Top)
            {
                obsText.Top = cancelButton.Top - obsText.Height - 10; // Keeps a 10 pixel gap from the Ok button
            }

            this.UIAPIRawForm.Freeze(false);
        }

        #endregion

        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            RemoveCheckedRowsFromMatrix(mOperations);
        }

        private void RemoveCheckedRowsFromMatrix(Matrix mOperations)
        {
            try
            {
                this.UIAPIRawForm.Freeze(true);
                SAPbouiCOM.DBDataSource oDBDataSource = (SAPbouiCOM.DBDataSource)this.UIAPIRawForm.DataSources.DBDataSources.Item("@STXQC19O");

                if (oDBDataSource.Size == 0)
                {
                    for (int rowIndex = mOperations.RowCount; rowIndex >= 1; rowIndex--)
                    {
                        // Get the value of the "OPcheck" column for the current row
                        SAPbouiCOM.CheckBox checkBox = (SAPbouiCOM.CheckBox)mOperations.Columns.Item("OPcheck").Cells.Item(rowIndex).Specific;

                        // Check if the checkbox is checked
                        if (checkBox.Checked)
                        {
                            checkBox.Checked = false;
                        }
                    }
                }
                else
                {
                    // Iterate through the rows in reverse order
                    for (int rowIndex = mOperations.RowCount; rowIndex >= 1; rowIndex--)
                    {
                        // Get the value of the "OPcheck" column for the current row
                        SAPbouiCOM.CheckBox checkBox = (SAPbouiCOM.CheckBox)mOperations.Columns.Item("OPcheck").Cells.Item(rowIndex).Specific;

                        // Check if the checkbox is checked
                        if (checkBox.Checked)
                        {
                            // Remove the row from the data source
                            oDBDataSource.RemoveRecord(rowIndex - 1);

                            if (rowIndex <= mOperations.RowCount)
                            {
                                mOperations.CommonSetting.SetRowBackColor(rowIndex, -1);
                            }
                        }
                    }

                    // Update the # aka LineID column
                    for (int i = 0; i < oDBDataSource.Size; i++)
                    {
                        oDBDataSource.SetValue("VisOrder", i, (i + 1).ToString());
                    }

                    mOperations.LoadFromDataSource();
                    if (mOperations.RowCount == 0)
                    {
                        AddRowIfMatrixEmpty();
                        int a = oDBDataSource.Size;
                    }
                }
                   
            }
            finally
            {
                QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                this.UIAPIRawForm.Freeze(false);
                this.UIAPIRawForm.Update();
            }
        }

        private void mOperations_DoubleClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ItemUID == "mOper" && pVal.ColUID == "OPcheck")
            {
                this.UIAPIRawForm.Freeze(true);
                if (!cellchecked)
                {
                    for (int i = 0; i < mOperations.RowCount; i++)
                    {
                        ((SAPbouiCOM.CheckBox)mOperations.Columns.Item("OPcheck").Cells.Item(i + 1).Specific).Checked = true;
                        cellchecked = true;
                    }
                }
                else
                {
                    for (int i = 0; i < mOperations.RowCount; i++)
                    {
                        ((SAPbouiCOM.CheckBox)mOperations.Columns.Item("OPcheck").Cells.Item(i + 1).Specific).Checked = false;
                        cellchecked = false;
                    }
                }
                this.UIAPIRawForm.Freeze(false);
            }

        }

        private void OPFilter_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            this.UIAPIRawForm.Freeze(true);
            try
            {
                // Get the selected value from the ComboBox
                string selectedValue = OPFilter.Selected.Value;

                string dataSourceId = mOperations.Columns.Item("OPSeq").DataBind.Alias.ToString();
                QCEvents.GetResultsfromFilter(this.UIAPIRawForm, mOperations, selectedValue);
                QCEvents.OperationsTotalFilter(this.UIAPIRawForm, selectedValue);
            }
            catch (Exception ex)
            {

                Program.SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, false);
            }
            finally
            {
                DisableFormWO();
                this.UIAPIRawForm.Freeze(false);
            }

        }

        private void QCSubPart_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            SBOChooseFromListEventArg chooseFromListEventArg = (SBOChooseFromListEventArg)pVal;
            string chooseFromListId = chooseFromListEventArg.ChooseFromListUID;
            SAPbouiCOM.ChooseFromList chooseFromList = this.UIAPIRawForm.ChooseFromLists.Item(chooseFromListId);

            // Get the selected item from the Choose From List
            SAPbouiCOM.DataTable selectedDataTable = chooseFromListEventArg.SelectedObjects;
            if (selectedDataTable != null && selectedDataTable.Rows.Count > 0)
            {
                string sptCode = selectedDataTable.GetValue("ItemCode", 0).ToString();
                subparttDescr = selectedDataTable.GetValue("ItemName", 0).ToString();

                this.SPartDescr.Value = parttDescr + ": " + subparttDescr; //this.SPartDescr.Value + ": " + parttDescr;
            }

        }

        private void QCPartType_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (oldPartType != this.QCPartType.Value)
            {
                this.QCSubPart.Value = "";
                this.SPartDescr.Value = parttDescr;

            }

        }

        private void QCPartType_GotFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            oldPartType = this.QCPartType.Value;
        }

        private void DefBOM_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            int noperations = mOperations.RowCount;
            
            try
            {
                this.UIAPIRawForm.Freeze(true);
                if (this.DefBOM.Checked == true)
                {
                    //if (noperations > 0)
                    //{
                        bool confirmDefBom = Program.SBO_Application.MessageBox("This operation will clear the current operations. Do you want to continue?", 1, "Yes", "No") == 1;
                        if (confirmDefBom)
                        {
                            this.mOperations.Clear();
                            QCEvents.GetDefOperations(this.UIAPIRawForm,selectedRow);
                            QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                        }
                    //}
                    //else
                    //{
                    //    this.mOperations.Clear();
                    //    QCEvents.GetDefOperations(this.UIAPIRawForm);
                    //    QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                    //}

                }
                else
                {
                    this.mOperations.Clear();
                    QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
                }

            }
            finally
            {
                DisableGTOperCC1(QCItemCode.Value);
                this.UIAPIRawForm.Freeze(false);
                PictureBox0.Picture = QCEvents.SellMarginImage(this.UIAPIRawForm);
            }
            
        }


        private void FormDataRecalculation()
        {
            var uIAPIRawForm = this.UIAPIRawForm;
            QCEvents.OperationsCalcTotal(this.UIAPIRawForm);
        }

        private void mOperations_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ItemUID != "mOper" && (pVal.ColUID != "OPResc" || pVal.ColUID != "OPcode" || !pVal.ActionSuccess))
                    return;

                if (pVal.ColUID == "OPResc" && pVal.ActionSuccess)
                {
                    SAPbouiCOM.EditText resQty = (SAPbouiCOM.EditText)mOperations.Columns.Item("OPQtdT").Cells.Item(pVal.Row).Specific;
                    SAPbouiCOM.EditText totalResc = (SAPbouiCOM.EditText)mOperations.Columns.Item("OPTotal").Cells.Item(pVal.Row).Specific;

                    var selectedData = GetSelectedResourceFromChooseFromList(pVal);
                    if (selectedData == null)
                        return;

                    string resCode = selectedData.Item1;
                    string resCodeN = selectedData.Item2;

                    string subCost = DBCalls.ResCost(resCode);
                    double subCostValue;

                    if (!double.TryParse(subCost, NumberStyles.Any, CultureInfo.InvariantCulture, out subCostValue))
                        return;

                    this.UIAPIRawForm.Freeze(true);

                    UpdateOperationControls(pVal.Row, resCode, resCodeN, subCostValue);

                    if (!recalcConfirm)
                    {
                        RecalculateLineData(pVal.Row, resCode, resQty, resQty.Value, totalResc.Value);
                        recalcConfirm = true;
                    }
                }
                if (pVal.ColUID == "OPcode" && pVal.ActionSuccess)
                {
                    var selectedData = GetSelectedOperationFromChooseFromList(pVal);
                    if (selectedData == null)
                        return;

                    string operCode = selectedData.Item1;
                    string oprName = selectedData.Item2;
                    string oprLocalName = selectedData.Item3;


                    mOperations.SetCellWithoutValidation(pVal.Row, "OPcode", operCode);
                    mOperations.SetCellWithoutValidation(pVal.Row, "OPName", oprName);
                    mOperations.SetCellWithoutValidation(pVal.Row, "OPNameL", oprLocalName);

                    mOperations.FlushToDataSource();
                    mOperations.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                Program.SBO_Application.SetStatusBarMessage(ex.Message, BoMessageTime.bmt_Short, true);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
            }
            
        }

        private void RecalculateLineData(int row, string resCode, SAPbouiCOM.EditText resQty, string resQtyValue, string totalRescValue)
        {
            if (recalcConfirm == false)
            {
                QCEvents.mtxLineDataRecalculation(this.UIAPIRawForm,resCode,resQty,resQtyValue,totalRescValue,previousLineTotal,"mOper",previousResc);
                recalcConfirm = true;
            }
        }

        private void UpdateOperationControls(int row, string resCode, string resCodeN, double subCostValue)
        {

            SAPbouiCOM.EditText resQty = (SAPbouiCOM.EditText)mOperations.Columns.Item("OPQtdT").Cells.Item(row).Specific;
            SAPbouiCOM.EditText errMsg = (SAPbouiCOM.EditText)mOperations.Columns.Item("OPErrMsg").Cells.Item(row).Specific;

            mOperations.SetCellWithoutValidation(row, "OPResc", resCode);
            mOperations.SetCellWithoutValidation(row, "OPResN", resCodeN);
            mOperations.SetCellWithoutValidation(row, "OPCost", subCostValue.ToString(CultureInfo.InvariantCulture));

            HandleResourceCostErrors(subCostValue, errMsg, row);

            double resQtyValue;

            if (double.TryParse(resQty.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out resQtyValue))
            {

                mOperations.SetCellWithoutValidation(row, "OPTotal", (subCostValue * resQtyValue).ToString(CultureInfo.InvariantCulture));

            }

            mOperations.FlushToDataSource();
            mOperations.LoadFromDataSource();
        }

        private void HandleResourceCostErrors(double subCostValue, EditText errMsg, int row)
        {
            Color orangeColor = Color.FromArgb(0xFF, 0xD1, 0x55);
            int warning = (orangeColor.R) + (orangeColor.G << 8) + (orangeColor.B << 16);

            if (subCostValue == 0)
            {
                errMsg.Value = Resources.mOperErr4;
                mOperations.CommonSetting.SetRowBackColor(row, warning);
            }
            else if (string.IsNullOrEmpty(errMsg.Value) || errMsg.Value == Resources.mOperErr1)
            {
                errMsg.Value = "";
                mOperations.CommonSetting.SetRowBackColor(row, -1);
            }
        }

        private Tuple<string, string> GetSelectedResourceFromChooseFromList(SBOItemEventArg pVal)
        {
            SBOChooseFromListEventArg chooseFromListEventArg = (SBOChooseFromListEventArg)pVal;
            SAPbouiCOM.ChooseFromList chooseFromList = this.UIAPIRawForm.ChooseFromLists.Item(chooseFromListEventArg.ChooseFromListUID);

            SAPbouiCOM.DataTable selectedDataTable = chooseFromListEventArg.SelectedObjects;
            if (selectedDataTable != null && selectedDataTable.Rows.Count > 0)
            {
                string resCode = selectedDataTable.GetValue("VisResCode", 0).ToString();
                string resCodeN = selectedDataTable.GetValue("ResName", 0).ToString();
                return new Tuple<string, string>(resCode, resCodeN);
            }
            return null;
        }

        private Tuple<string, string, string> GetSelectedOperationFromChooseFromList(SBOItemEventArg pVal)
        {
            SBOChooseFromListEventArg chooseFromListEventArg = (SBOChooseFromListEventArg)pVal;
            SAPbouiCOM.ChooseFromList chooseFromList = this.UIAPIRawForm.ChooseFromLists.Item(chooseFromListEventArg.ChooseFromListUID);

            SAPbouiCOM.DataTable selectedDataTable = chooseFromListEventArg.SelectedObjects;
            if (selectedDataTable != null && selectedDataTable.Rows.Count > 0)
            {
                string operCode = selectedDataTable.GetValue("Code", 0).ToString();
                string operName = selectedDataTable.GetValue("U_STXOPDes", 0).ToString();
                string operLocalName = selectedDataTable.GetValue("U_STXOPDesLocal", 0).ToString();
                return new Tuple<string, string, string>(operCode, operName, operLocalName);
            }
            return null;
        }

        private void mOperations_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ItemUID == "mOper" && pVal.ColUID == "OPQtdT" && pVal.ActionSuccess == true)
            {
                SAPbouiCOM.EditText opResc = (SAPbouiCOM.EditText)mOperations.Columns.Item("OPResc").Cells.Item(pVal.Row).Specific;
                SAPbouiCOM.EditText opNewQty = (SAPbouiCOM.EditText)mOperations.Columns.Item("OPQtdT").Cells.Item(pVal.Row).Specific;
                SAPbouiCOM.EditText opRescCost = (SAPbouiCOM.EditText)mOperations.Columns.Item("OPCost").Cells.Item(pVal.Row).Specific;

                double opNewQtyValue, opRescCostValue;

                if (double.TryParse(opNewQty.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out opNewQtyValue) && double.TryParse(opRescCost.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out opRescCostValue))
                {
                    newCost = (opNewQtyValue * opRescCostValue).ToString(CultureInfo.InvariantCulture);
                    mOperations.SetCellWithoutValidation(pVal.Row, "OPTotal", newCost);
                }
                else
                {
                    newCost = "0";
                    mOperations.SetCellWithoutValidation(pVal.Row, "OPTotal", newCost);
                }
                this.UIAPIRawForm.Freeze(true);
                mOperations.FlushToDataSource();
                mOperations.LoadFromDataSource();

                if (recalcConfirm == false)
                {
                    QCEvents.mtxLineDataRecalculation(this.UIAPIRawForm, opResc.Value, opNewQty, previousQty, newCost, previousLineTotal, pVal.ItemUID, previousResc);
                    PictureBox0.Picture = QCEvents.SellMarginImage(this.UIAPIRawForm);
                    recalcConfirm = true;
                }

                this.UIAPIRawForm.Freeze(false);
            }
        }

        private void mOperations_GotFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ItemUID == "mOper")
            {
                SAPbouiCOM.EditText opQty = (SAPbouiCOM.EditText)mOperations.Columns.Item("OPQtdT").Cells.Item(pVal.Row).Specific;
                previousQty = opQty.Value;
                SAPbouiCOM.EditText opTotal = (SAPbouiCOM.EditText)mOperations.Columns.Item("OPTotal").Cells.Item(pVal.Row).Specific;
                previousLineTotal = opTotal.Value;
                SAPbouiCOM.EditText OPResc = (SAPbouiCOM.EditText)mOperations.Columns.Item("OPResc").Cells.Item(pVal.Row).Specific;
                previousResc = OPResc.Value;
                recalcConfirm = false;
            }


        }




        private void mTextures_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ItemUID == "mTextures" && pVal.ColUID == "QCCovA" && pVal.ActionSuccess == true)
            {
                System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();
                SAPbouiCOM.EditText cov = (SAPbouiCOM.EditText)mTextures.Columns.Item("QCCovA").Cells.Item(pVal.Row).Specific;

                if (lostFocusCovA)
                {
                    lostFocusCovA = false;
                    return;
                }

                double covA = 0;

                try
                {

                    covA = HelperMethods.ParseSAPValueToDouble(Regex.Replace((string.IsNullOrEmpty(cov.Value) ? "0" : cov.Value), $@"[^\d{Utils.decSep}{Utils.thousSep}]", ""));

                }
                catch (Exception)
                {
                    covA = 0;
                    Program.SBO_Application.SetStatusBarMessage("Please, place a numeric value.", BoMessageTime.bmt_Short, true);
                }
                string formattedQCLength = covA.ToString("N", sapNumberFormat);

                cov.Value = $"{formattedQCLength} {selectedUOM}²";
                lostFocusCovA = true;
            }
        }

        private void Form_UnloadBefore(SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            UnloadResults = new QCResults
            {
                QCID = this.QCDocEntry.Value,
                QCLineN = this.BaseLine.Value,
                QCuPrice = this.UnPrice.Value,  
                QClTime = this.QCLeadTime.Value,  
                QCcPrice = this.QCTEst.Value,
                QCtNum = this.QCToolNum.Value,
                QCptNum = this.QCPartNum.Value,
                QCprtName = this.QCPartName.Value
            };
            if (formUpdateTrigger == true)
            {
                DBCalls.UpdateSAPDocument(UnloadResults,sapDocEntry,sapObjType,sapDocLineNum);

                if (!string.IsNullOrEmpty(Utils.ParentFormUID))
                {
                    SAPbouiCOM.Form parentForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(Utils.ParentFormUID);
                    parentForm.Select();
                    Program.SBO_Application.ActivateMenuItem("1304");
                }
            }
            
        }

        private void QCHeight_GotFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            currentHeight = QCHeight.Value;

        }

        private void QCHeight_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();
            double qcHeight = 0;

            if (currentHeight != QCHeight.Value)
            {
                if (lostFocusQCHeight)
                {
                    lostFocusQCHeight = false;
                    return;
                }
                try
                {
                    qcHeight = double.Parse(this.QCHeight.Value, System.Globalization.NumberStyles.AllowDecimalPoint | System.Globalization.NumberStyles.AllowThousands, sapNumberFormat);
                }
                catch (Exception)
                {
                    qcHeight = 0;
                    Program.SBO_Application.SetStatusBarMessage("Please, place a numeric value.", BoMessageTime.bmt_Short, true);
                }

                string formattedQCHeight = qcHeight.ToString("N", sapNumberFormat);

                this.QCHeight.Value = $"{formattedQCHeight} {selectedUOM}";
                QCEvents.CalculateArea(this.UIAPIRawForm.UniqueID, selectedUOM);
                lostFocusQCHeight = true;
            }

        }
    }
}
