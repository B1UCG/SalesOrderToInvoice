using System;
using System.Windows.Forms;
using SAPbouiCOM;
using Application = SAPbouiCOM.Framework.Application;

namespace SkyTech.SalesOrderToInvoice {
    internal static class Program {
        /// <summary>The main entry point for the application.</summary>
        [STAThread]
        private static void Main(string[] args) {
            try {
                var oApp = args.Length < 1 ? new Application() : new Application(args[0]);
                Application.SBO_Application.AppEvent += AppEventHandler;
                oApp.Run();
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private static void AppEventHandler(BoAppEventTypes eventType) {
            switch (eventType) {
                case BoAppEventTypes.aet_ShutDown:
                case BoAppEventTypes.aet_CompanyChanged:
                case BoAppEventTypes.aet_ServerTerminition:
                    System.Windows.Forms.Application.Exit();
                    break;
            }
        }
    }
}