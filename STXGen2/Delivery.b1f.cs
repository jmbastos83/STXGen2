
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace STXGen2
{

    [FormAttribute("140", "Delivery.b1f")]
    class Delivery : SystemFormBase
    {
        public static bool ddWizard { get; set; }

        public Delivery()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("MKSeg1").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("MKSEG2").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("STXBrand").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("NBOID").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("OEMPgm").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("OEM").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("GKAM").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_9").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_11").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_12").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_13").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ActivateAfter += new SAPbouiCOM.Framework.FormBase.ActivateAfterHandler(this.Form_ActivateAfter);
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private void OnCustomInitialize()
        {

        }

        private void Form_ActivateAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (ddWizard)
            {
                try
                {
                    this.UIAPIRawForm.Freeze(true);
                    SAPbouiCOM.Matrix itemMatrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("38").Specific;
                    for (int i = itemMatrix.RowCount; i >= 1; i--)
                    {
                        string itemCode = (string)((SAPbouiCOM.EditText)itemMatrix.Columns.Item("1").Cells.Item(i).Specific).Value;

                        if (itemCode.StartsWith("FG"))
                        {
                            itemMatrix.DeleteRow(i);
                        }
                    }
                    ddWizard = false;
                }
                finally
                {
                    this.UIAPIRawForm.Freeze(false);
                }
               
            }
        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.StaticText StaticText6;

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (ddWizard)
            {
                try
                {
                    this.UIAPIRawForm.Freeze(true);
                    SAPbouiCOM.Matrix itemMatrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("38").Specific;
                    for (int i = itemMatrix.RowCount; i >= 1; i--)
                    {
                        string itemCode = (string)((SAPbouiCOM.EditText)itemMatrix.Columns.Item("1").Cells.Item(i).Specific).Value;

                        if (itemCode.StartsWith("FG"))
                        {
                            itemMatrix.DeleteRow(i);
                        }
                    }
                    ddWizard = false;
                }
                finally
                {
                    this.UIAPIRawForm.Freeze(false);
                }

            }

        }
    }
}
