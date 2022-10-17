using System;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace SkyTech.SalesOrderToInvoice.Data {
    public static class SalesInvoice {
        public static int CopyFromOrder(int docEntry, out string errMessage) {
            errMessage = string.Empty;
            try {
                var company = (Company) Application.SBO_Application.Company.GetDICompany();
                var order   = (Documents) company.GetBusinessObject(BoObjectTypes.oOrders);
                var invoice = (Documents) company.GetBusinessObject(BoObjectTypes.oInvoices);
                order.GetByKey(docEntry);
                invoice.CardCode   = order.CardCode;
                invoice.DocDate    = order.DocDate;
                invoice.DocDueDate = order.DocDueDate;
                invoice.TaxDate    = order.TaxDate;
                for (int i = 0; i < order.Lines.Count; i++) {
                    order.Lines.SetCurrentLine(i);
                    if (i > 0) invoice.Lines.Add();
                    invoice.Lines.BaseType  = (int) order.DocObjectCode;
                    invoice.Lines.BaseEntry = order.DocEntry;
                    invoice.Lines.BaseLine  = order.Lines.LineNum;
                }

                int retCode = invoice.Add();
                if (retCode == 0) return int.Parse(company.GetNewObjectKey());
                errMessage = company.GetLastErrorDescription();
                return -1;
            }
            catch (Exception e) {
                Application.SBO_Application.SetStatusBarMessage(e.Message);
            }

            return -1;
        }
    }
}