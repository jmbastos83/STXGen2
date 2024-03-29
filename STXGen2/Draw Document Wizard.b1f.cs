
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace STXGen2
{

    [FormAttribute("425", "Draw Document Wizard.b1f")]
    class Draw_Document_Wizard : SystemFormBase
    {
        public Draw_Document_Wizard()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("49").Specific));
            this.Button0.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.Button0_PressedBefore);
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

        }

        private SAPbouiCOM.Button Button0;

        private void Button0_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.OptionBtn drawAll = (SAPbouiCOM.OptionBtn)this.UIAPIRawForm.Items.Item("47").Specific;
            SAPbouiCOM.OptionBtn customDraw = (SAPbouiCOM.OptionBtn)this.UIAPIRawForm.Items.Item("48").Specific;

            SAPbouiCOM.Matrix itemMatrix = (SAPbouiCOM.Matrix)this.UIAPIRawForm.Items.Item("3").Specific;
            int rowCount = itemMatrix.RowCount;

            if (drawAll.Selected)
            {
                Delivery.ddWizard = true;
            }
            if (customDraw.Selected)
            {
                for (int rowIndex = 1; rowIndex <= itemMatrix.RowCount; rowIndex++)
                {
                    if (itemMatrix.IsRowSelected(rowIndex))
                    {
                        string itemCode = (string)((SAPbouiCOM.EditText)itemMatrix.Columns.Item("2").Cells.Item(rowIndex).Specific).Value;

                        if (itemCode.StartsWith("FG"))
                        {
                            Program.SBO_Application.SetStatusBarMessage("Please deselect FG Items.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            BubbleEvent = false;
                        }
                    }
                }
            }
            

        }
    }
}
