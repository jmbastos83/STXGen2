﻿using SAPbouiCOM;
using System.Globalization;

namespace STXGen2
{
    internal static class HelperMethods
    {
        private static readonly System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();

        internal static SAPbouiCOM.Matrix GetMatrix(IForm uIAPIRawForm, string itemId)
            => (SAPbouiCOM.Matrix)uIAPIRawForm.Items.Item(itemId).Specific;

        internal static double ParseValueToDouble(string value)
            => double.Parse(value, NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);

        internal static double ParseSAPValueToDouble(string value)
            => double.Parse(value, sapNumberFormat);

        internal static void UpdateEditText(IForm uIAPIRawForm, string itemId, double total)
        {
            SAPbouiCOM.EditText myTotalEditText = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item(itemId).Specific;
            myTotalEditText.Value = total.ToString("N", sapNumberFormat);
        }
    }
}