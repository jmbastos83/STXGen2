using SAPbobsCOM;
using System;
using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace STXGen2
{
    public static class Utils
    {
        public static SAPbobsCOM.Company oCompany;

        public static string ParentFormUID { get; set; }
        public static string decSep { get; private set; }
        public static string thousSep { get; private set; }
        public static int MeasureDec { get; private set; }
        public static int PriceDec { get; private set; }
        public static int SumDec { get; private set; }
        public static int QtyDec { get; private set; }
        public static string MainCurrency { get; private set; }
        public static string SystemCurrency { get; private set; }
        public static string DirectRate { get; private set; }

        internal static void CompSettings()
        {
            string sSql = $"select \"DecSep\",\"ThousSep\",\"MeasureDec\",\"PriceDec\",\"SumDec\",\"QtyDec\",\"MainCurncy\",\"SysCurrncy\",\"DirectRate\" from \"OADM\"";
            Recordset rs = Utils.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            rs.DoQuery(sSql);
            if (!rs.EoF)
            {
                decSep = string.IsNullOrEmpty((string)rs.Fields.Item("DecSep").Value) ? "" : (string)rs.Fields.Item("DecSep").Value;
                thousSep = string.IsNullOrEmpty((string)rs.Fields.Item("ThousSep").Value) ? "" : (string)rs.Fields.Item("ThousSep").Value;
                MeasureDec = (int)rs.Fields.Item("MeasureDec").Value;
                PriceDec = (int)rs.Fields.Item("PriceDec").Value;
                SumDec = (int)rs.Fields.Item("SumDec").Value;
                QtyDec = (int)rs.Fields.Item("QtyDec").Value;
                MainCurrency = (string)rs.Fields.Item("MainCurncy").Value;
                SystemCurrency = (string)rs.Fields.Item("SysCurrncy").Value;
                DirectRate = (string)rs.Fields.Item("DirectRate").Value;
            }
        }

        internal static void InitialSetup()
        {
            DBStructure.VerifyTables();
            DBStructure.VerifyUDF();
        }

        public static System.Globalization.NumberFormatInfo GetSAPNumberFormatInfo()
        {
            System.Globalization.NumberFormatInfo sapNumberFormat = new System.Globalization.NumberFormatInfo();
            sapNumberFormat.NumberDecimalSeparator = Utils.decSep;
            sapNumberFormat.NumberGroupSeparator = Utils.thousSep;
            return sapNumberFormat;
        }

        public static System.Globalization.CultureInfo GetCompanyCulture()
        {
            SAPbouiCOM.BoLanguages language = Program.SBO_Application.Language;
            string cultureCode = "en-US";

            switch (Program.SBO_Application.Language)
            {
                case SAPbouiCOM.BoLanguages.ln_Hebrew:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Spanish_Ar:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Polish:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Spanish_Pa:
                    break;
                case SAPbouiCOM.BoLanguages.ln_English:
                case SAPbouiCOM.BoLanguages.ln_English_Gb:
                case SAPbouiCOM.BoLanguages.ln_English_Sg:
                    System.Globalization.CultureInfo.DefaultThreadCurrentUICulture = new System.Globalization.CultureInfo("en-GB");
                    break;
                case SAPbouiCOM.BoLanguages.ln_German:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Serbian:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Danish:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Norwegian:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Italian:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Hungarian:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Chinese:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Dutch:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Finnish:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Greek:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Portuguese:
                    System.Globalization.CultureInfo.DefaultThreadCurrentUICulture = new System.Globalization.CultureInfo("pt-PT");
                    break;
                case SAPbouiCOM.BoLanguages.ln_Swedish:
                    break;
                case SAPbouiCOM.BoLanguages.ln_English_Cy:
                    break;
                case SAPbouiCOM.BoLanguages.ln_French:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Spanish:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Russian:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Spanish_La:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Czech_Cz:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Slovak_Sk:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Korean_Kr:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Portuguese_Br:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Japanese_Jp:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Turkish_Tr:
                    break;
                case SAPbouiCOM.BoLanguages.ln_Ukrainian:
                    break;
                case SAPbouiCOM.BoLanguages.ln_TrdtnlChinese_Hk:
                    break;
                default:
                    break;
            }
            return new System.Globalization.CultureInfo(cultureCode);
        }

        public static string XmlSerializeToString(this object objectInstance)
        {
            XmlSerializer serializer = new XmlSerializer(objectInstance.GetType());
            StringBuilder sb = new StringBuilder();
            using (TextWriter writer = new StringWriter(sb))
            {
                serializer.Serialize(writer, objectInstance);
            }
            return sb.ToString();
        }

        public static T XmlDeserializeFromString<T>(this string objectData)
        {
            return (T)XmlDeserializeFromString(objectData, typeof(T));
        }

        public static object XmlDeserializeFromString(this string objectData, Type type)
        {
            if (string.IsNullOrEmpty(objectData))
            {
                throw new ArgumentException("The input string cannot be null or empty.", nameof(objectData));
            }

            try
            {
                XmlSerializer serializer = new XmlSerializer(type);
                using (TextReader reader = new StringReader(objectData))
                {
                    return serializer.Deserialize(reader);
                }
            }
            catch (InvalidOperationException ex) // Handles XML format issues or type mismatch issues
            {
                throw new InvalidOperationException($"Failed to deserialize the XML string into the type {type.FullName}.", ex);
            }
        }

        internal static void OpenImage(string imagePath)
        {
            try
            {
                imagePath = Path.Combine(!Directory.Exists(Path.Combine(Utils.oCompany.BitMapPath, "Tools Images")) ? Utils.oCompany.BitMapPath : Path.Combine(Utils.oCompany.BitMapPath, "Tools Images"), imagePath);

                // Ensure that the path exists
                if (File.Exists(imagePath))
                {
                    

                    // Option 1: Start the associated program to open the image file
                    System.Diagnostics.Process.Start(imagePath);

                    // Option 2: Open a form with a PictureBox to display the image
                    // You would need to create a form with a PictureBox control and load the image into it
                    // ShowImageForm(imagePath);
                }
                else
                {
                    // Handle the case where the image path does not exist
                    Program.SBO_Application.SetStatusBarMessage("Image file not found.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions that occur
                Program.SBO_Application.SetStatusBarMessage("Error opening image: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }
    }
}