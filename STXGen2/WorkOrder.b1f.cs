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
    }
}
