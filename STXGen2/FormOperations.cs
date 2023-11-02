using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;

namespace STXGen2
{
    class FormOperations
    {
        internal static void CleanNBOInfo(IForm uIAPIRawForm)
        {
            try
            {
                ItemData.DisableNBOinfo(uIAPIRawForm);

                ((SAPbouiCOM.EditText)(uIAPIRawForm.Items.Item("GKAM").Specific)).Value = string.Empty;
                ((SAPbouiCOM.EditText)(uIAPIRawForm.Items.Item("OEM").Specific)).Value = string.Empty;
                ((SAPbouiCOM.EditText)(uIAPIRawForm.Items.Item("OEMPgm").Specific)).Value = string.Empty;
                ((SAPbouiCOM.EditText)(uIAPIRawForm.Items.Item("NBOID").Specific)).Value = string.Empty;
                ((SAPbouiCOM.EditText)(uIAPIRawForm.Items.Item("STXBrand").Specific)).Value = string.Empty;
                ((SAPbouiCOM.EditText)(uIAPIRawForm.Items.Item("MKSEG2").Specific)).Value = string.Empty;
                ((SAPbouiCOM.EditText)(uIAPIRawForm.Items.Item("MKSeg1").Specific)).Value = string.Empty;
                ((SAPbouiCOM.EditText)(uIAPIRawForm.Items.Item("BrandID").Specific)).Value = string.Empty;
                ((SAPbouiCOM.EditText)(uIAPIRawForm.Items.Item("MK2ID").Specific)).Value = string.Empty;
                ((SAPbouiCOM.EditText)(uIAPIRawForm.Items.Item("MK1ID").Specific)).Value = string.Empty;

                ItemData.EnableNBOinfo(uIAPIRawForm);

                ((SAPbouiCOM.Button)(uIAPIRawForm.Items.Item("ClrNBO").Specific)).Item.Enabled = false;
                ((SAPbouiCOM.EditText)(uIAPIRawForm.Items.Item("MKSeg1").Specific)).Active = true;

            }
            catch (Exception)
            {
                Program.SBO_Application.SetStatusBarMessage("Failed do Clear NBO information.", BoMessageTime.bmt_Short, true);
            }
        }

        internal static void GetNBOInfo(IForm uIAPIRawForm, (string sMkSeg1Name, string sMkseg1ID, string sBrandName, string sBrandID, string sOEM, string sOEMProgram, string sGKAM)? result)
        {
            try
            {
                ItemData.DisableNBOinfo(uIAPIRawForm);
                SAPbouiCOM.EditText eMkseg1 = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("MKSeg1").Specific;
                SAPbouiCOM.EditText eMk1ID = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("MK1ID").Specific;
                SAPbouiCOM.EditText eBrand = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("STXBrand").Specific;
                SAPbouiCOM.EditText eBrandID = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("BrandID").Specific;
                SAPbouiCOM.EditText eOEMPgm = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("OEMPgm").Specific;
                SAPbouiCOM.EditText eOEM = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("OEM").Specific;
                SAPbouiCOM.EditText eGKam = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("GKAM").Specific;
                (string sMkSeg1Name, string sMkseg1ID, string sBrandName, string sBrandID, string sOEM, string sOEMProgram, string sGKAM) = result.Value;

                if ((sOEMProgram.StartsWith("PH-") || sOEMProgram.StartsWith("PH_")) && eMkseg1.Value.ToString() != "")
                {
                    eOEMPgm.Value = sOEMProgram;
                    //eOEM.Value = sOEM;
                    //eGKam.Value = sGKAM;
                }
                else
                {
                    eMkseg1.Value = sMkSeg1Name;
                    eMk1ID.Value = sMkseg1ID;
                    eBrand.Value = sBrandName;
                    eBrandID.Value = sBrandID;
                    eOEMPgm.Value = sOEMProgram;
                    eOEM.Value = sOEM;
                    eGKam.Value = sGKAM;
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                ItemData.EnableNBOinfo(uIAPIRawForm);
            }
        }

        internal static void GetBrandInfo(IForm uIAPIRawForm, (string sMkSeg1Name, string sMkseg1ID, string sBrandID, string sOEM, string sGKAM)? result)
        {
            try
            {
                ItemData.DisableBrandInfo(uIAPIRawForm);
                SAPbouiCOM.EditText eMkseg1 = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("MKSeg1").Specific;
                SAPbouiCOM.EditText eMk1ID = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("MK1ID").Specific;
                SAPbouiCOM.EditText eBrand = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("STXBrand").Specific;
                SAPbouiCOM.EditText eBrandID = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("BrandID").Specific;
                SAPbouiCOM.EditText eOEM = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("OEM").Specific;
                SAPbouiCOM.EditText eGKam = (SAPbouiCOM.EditText)uIAPIRawForm.Items.Item("GKAM").Specific;
                (string sMkSeg1Name, string sMkseg1ID, string sBrandID, string sOEM, string sGKAM) = result.Value;


                eMkseg1.Value = sMkSeg1Name;
                eMk1ID.Value = sMkseg1ID;
                eBrandID.Value = sBrandID;
                eOEM.Value = sOEM;
                eGKam.Value = sGKAM;
            }
            catch (Exception)
            {

                throw;
            }
           
            finally
            {
                ItemData.EnableBrandInfo(uIAPIRawForm);
            }
        }
    }
}
