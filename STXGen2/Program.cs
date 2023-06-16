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

                Utils.oCompany = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
                
                Utils.CompSettings();

                Utils.GetCompanyCulture();

                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        
    }
}