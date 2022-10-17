using System;
using System.Linq;
using System.Xml.Linq;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using SkyTech.SalesOrderToInvoice.Data;
using Application = SAPbouiCOM.Framework.Application;

namespace SkyTech.SalesOrderToInvoice.Forms {
    [Form("139")]
    internal class SalesOrderForm : SystemFormBase {
        private static Matrix ItemsMatrix { get; set; }

        /// <summary>Initialize components. Called by framework after form created.</summary>
        public override void OnInitializeComponent() {
            ItemsMatrix = GetSpecificItem<Matrix>("38");
            OnCustomInitialize();
        }

        /// <summary>Initialize form event. Called by framework before form creation.</summary>
        public override void OnInitializeFormEvents() {
            DataAddBefore += DataAddBeforeEventHandler;
            DataAddAfter  += DataAddAfterEventHandler;
        }

        private static void DataAddBeforeEventHandler(ref BusinessObjectInfo pVal, out bool bubbleEvent) {
            bubbleEvent = true;
            try {
                if (CheckOrderLines()) return;
                bubbleEvent = false;
                Application.SBO_Application.MessageBox("Operation not allowed, the Sales Order should contain at least 3 lines.");
            }
            catch (Exception e) {
                Application.SBO_Application.StatusBar.SetText(e.Message);
            }
        }

        private static void DataAddAfterEventHandler(ref BusinessObjectInfo pVal) {
            try {
                string orderKey = XDocument.Parse(pVal.ObjectKey).Descendants().First(p => p.Name == "DocEntry").Value;
                if (string.IsNullOrWhiteSpace(orderKey)) return;
                int invKey = SalesInvoice.CopyFromOrder(int.Parse(orderKey), out string errMessage);
                if (invKey <= 0) {
                    Application.SBO_Application.MessageBox($"Failed to Add Invoice :: {errMessage}");
                    return;
                }

                Application.SBO_Application.MessageBox($"Invoice added successfully :: [DocEntry: {invKey}]");
            }
            catch (Exception e) {
                Application.SBO_Application.StatusBar.SetText(e.Message);
            }
        }

        private static void OnCustomInitialize() { }

        private static bool CheckOrderLines() {
            int validLines = 0;
            for (int i = 0; i < ItemsMatrix.VisualRowCount; i++) {
                string itemCode = ((EditText) ItemsMatrix.GetCellSpecific("1", i + 1)).Value;
                if (!string.IsNullOrWhiteSpace(itemCode)) validLines++;
                if (validLines >= 3) return true;
            }

            return false;
        }
    }
}