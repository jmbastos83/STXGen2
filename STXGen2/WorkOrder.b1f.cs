using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace STXGen2
{
    [FormAttribute("65211", "WorkOrder.b1f")]
    class WorkOrder : SystemFormBase
    {
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.EditText EditText0;

        public WorkOrder()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("SONum").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_2").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("Item_4").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("Item_5").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("Item_6").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("WOType").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_8").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("Item_10").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_9").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_11").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_12").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_13").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_14").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_15").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_16").Specific));
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

            //this.UIAPIRawForm.DataSources.DBDataSources.Add("OWOR");

            //var etSONum = (SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("SONum").Specific;
            //;
            //etSONum.DataBind.SetBound(true, "OWOR", "U_STXSONum");

        }

        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.StaticText StaticText7;
    }
}
