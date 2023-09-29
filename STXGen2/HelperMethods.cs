using SAPbouiCOM;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace STXGen2
{
    internal static class HelperMethods
    {
        private static readonly System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();

        internal static SAPbouiCOM.Matrix GetMatrix(IForm uIAPIRawForm, string itemId)
            => (SAPbouiCOM.Matrix)uIAPIRawForm.Items.Item(itemId).Specific;

        internal static double ParseValueToDouble(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return 0;
            }

            return double.Parse(value, NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
        }

        internal static double ParseSAPValueToDouble(string value)
            => double.Parse(value, sapNumberFormat);

        internal static void UpdateEditText(IForm uIAPIRawForm, string itemId, double total)
        {
            SAPbouiCOM.EditText myTotalEditText = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item(itemId).Specific;
            myTotalEditText.Value = total.ToString("N", sapNumberFormat);
        }

        internal static double ParseDoubleWCur(string value, NumberFormatInfo numberFormat)
        {
            return double.Parse(Regex.Replace((string.IsNullOrEmpty(value) ? "0" : value), $@"[^\d{Utils.decSep}{Utils.thousSep}]", ""), NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, numberFormat);

        }
        internal static double ParseDoubleWUOM(string value, NumberFormatInfo numberFormat)
        {
            return double.Parse(Regex.Replace((string.IsNullOrEmpty(value) ? "0" : value), $@"[^\d{Utils.decSep}{Utils.thousSep}]", ""), NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, numberFormat);

        }
        internal static string FormatValueCur(double value, string currency)
        {
            System.Globalization.NumberFormatInfo sapNumberFormat = Utils.GetSAPNumberFormatInfo();
            return $"{value.ToString("#,0.00", sapNumberFormat)} {currency}";
        }

        internal static string GetAndSaveImage(string imageName)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string prefix = "STXGen2.Properties.";
            string resourceName = prefix + imageName;

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    Image image = Image.FromStream(stream);
                    string imagePath = Path.Combine(Path.GetTempPath(), imageName);
                    image.Save(imagePath);
                    return imagePath;
                }
            }

            return string.Empty;
        }
    }
}