using System;
using SAPbouiCOM;

namespace STXGen2
{
    internal class SAPForms
    {
        internal static void updateSystemMatrix(Form activeForm, int SysFormLine)
        {
            SAPbouiCOM.Matrix sysFormMatrix = (SAPbouiCOM.Matrix)activeForm.Items.Item("matrixItemId").Specific;

            EditText grossPrice = (EditText)sysFormMatrix.Columns.Item("124").Cells.Item(SysFormLine).Specific;
            grossPrice.Value = "0";
        }
    }
}