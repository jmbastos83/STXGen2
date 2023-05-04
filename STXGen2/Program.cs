using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using SAPbouiCOM.Framework;
using STXGen2.Properties;

namespace STXGen2
{
    class Program
    {
        public static SAPbouiCOM.Application SBO_Application;
        public static SqlConnection connection;
        //private static SAPbobsCOM.Company oCompany;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    //If you want to use an add-on identifier for the development license, you can specify an add-on identifier string as the second parameter.
                    //oApp = new Application(args[0], "XXXXX");
                    oApp = new Application(args[0]);
                }

                SBO_Application = Application.SBO_Application;
                // Register the Application event handler
                SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SAPEvents.SBO_Application_AppEvent);
                // Register the Menu event handler
                SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SAPEvents.SBO_Application_MenuEvent);
                // Register the RightClickEvent event handler
                SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SAPEvents.SBO_Application_RightClickEvent);

                //// Register the ItemEvent event handler for the Sales Quotation form
                //SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SAPEvents.SBO_Application_ItemEvent);

                Utils.oCompany = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
                
                Utils.CompSettings();

                switch (Application.SBO_Application.Language)
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

                //System.Globalization.CultureInfo.DefaultThreadCurrentUICulture = new System.Globalization.CultureInfo("en-GB");
                //System.Globalization.CultureInfo.DefaultThreadCurrentUICulture = new System.Globalization.CultureInfo("pt-PT");
                //System.Windows.Forms.MessageBox.Show(Resources.Msg1);

                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
    }
}