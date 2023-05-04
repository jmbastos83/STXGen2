using SAPbobsCOM;
using SAPbouiCOM;
using STXGen2.Properties;
using System;
using System.Collections;
using System.Globalization;
using System.Resources;
using System.Threading;

namespace STXGen2
{
    internal class FormTranslations
    {
        public static void SetStaticTextTranslations(Form form)
        {
            //CompanyService companyService = (CompanyService)Utils.oCompany.GetCompanyService();
            //CompanyInfo companyInfo = companyService.GetCompanyInfo();

            string langCodeStr = DBCalls.GetUserLanguage();
           
            BoSuppLangs langCode = (BoSuppLangs)Enum.Parse(typeof(BoSuppLangs), langCodeStr);

            switch (langCode)
            {
                case BoSuppLangs.ln_English:
                    // set static text translations for English
                    SetEnglishTranslations(form);
                    break;
                case BoSuppLangs.ln_Chinese:
                    // set static text translations for Chinese
                    SetChineseTranslations(form);
                    break;
                case BoSuppLangs.ln_French:
                    // set static text translations for Chinese
                    SetFrenchTranslations(form);
                    break;
                // add cases for other languages as needed
                default:
                    // handle unknown language code
                    SetEnglishTranslations(form);
                    break;
            }
        }



        private static void SetEnglishTranslations(Form form)
        {
            ((StaticText)form.Items.Item("lItemCode").Specific).Caption = Resources.ItemCodeLabel;
            ((StaticText)form.Items.Item("lToolNum").Specific).Caption = "Tool Number";


            // set static text translations for English
            ((StaticText)form.Items.Item("lItemCode").Specific).Caption = "ItemCode";
            ((StaticText)form.Items.Item("lToolNum").Specific).Caption = "Tool Number";
            // set other translations for English
        }

        private static void SetChineseTranslations(Form form)
        {
            // set static text translations for Chinese
            ((StaticText)form.Items.Item("lItemCode").Specific).Caption = "物料编码";
            ((StaticText)form.Items.Item("lToolNum").Specific).Caption = "工具号码";
            // set other translations for Chinese
        }

        private static void SetFrenchTranslations(Form form)
        {
            // set static text translations for English
            ((StaticText)form.Items.Item("lItemCode").Specific).Caption = "ItemCode";
            ((StaticText)form.Items.Item("lToolNum").Specific).Caption = "Numéro d’outil";
            // set other translations for English
        }
        

    }
}