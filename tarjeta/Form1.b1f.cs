using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using SAPbouiCOM;
using SAPbobsCOM;

namespace tarjeta
{
    [FormAttribute("tarjeta.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.txtDoc = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.btnGenerar = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.btnGenerar.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {
            
        }

        #region VARIABLES
        private SAPbouiCOM.EditText txtDoc;
        private SAPbouiCOM.Button btnGenerar;
        private SAPbouiCOM.Form oForm;
        #endregion


        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
           
            //Actualizamos el campo de conciliacion
            string v_pago = null;
            SAPbobsCOM.Documents oFactura;
            oFactura = (SAPbobsCOM.Documents)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            if (oFactura.GetByKey(int.Parse(txtDoc.Value)))
            {
                oFactura.UserFields.Fields.Item("U_U_Destino").Value = "NO";
                int error = oFactura.Update();
                if (error != 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(Menu.sbo.GetLastErrorDescription(), 1, "OK");
                    return;
                }
            }

            //Cancelamos el pago
            SAPbobsCOM.Recordset oConsulta;
            oConsulta = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oConsulta.DoQuery("SELECT \"ReceiptNum\" FROM OINV WHERE \"DocEntry\"="+txtDoc.Value+" ");
            v_pago = oConsulta.Fields.Item(0).Value.ToString();

            SAPbobsCOM.Payments oPago;
            oPago = (SAPbobsCOM.Payments)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            if (oPago.GetByKey(int.Parse(v_pago)))
            {
                string pp = oPago.Invoices.DocEntry.ToString();
                int error = oPago.Cancel();
                if(error != 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(Menu.sbo.GetLastErrorDescription(), 1, "OK");
                    return;
                }
                else
                {
                    //realizamos en contra asiento
                    string v_cuentaDeb = null;
                    string v_cuentaCred = null;
                    double v_monto = 0;
                    int fila = 0;
                    SAPbobsCOM.Recordset oDatos;
                    oDatos = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oDatos.DoQuery("SELECT T0.\"Account\",T0.\"Debit\",T0.\"Credit\" FROM JDT1 T0 INNER JOIN OJDT T1 ON T0.\"TransId\"=T1.\"TransId\" WHERE T1.\"U_CyC_DocNum\"='" + txtDoc.Value + "' ");
                    while (!oDatos.EoF)
                    {
                        if (fila == 0)
                        {
                            v_cuentaDeb = oDatos.Fields.Item(0).Value.ToString();
                        }
                        if (fila == 1)
                        {
                            v_cuentaCred = oDatos.Fields.Item(0).Value.ToString();
                        }
                        v_monto = double.Parse(oDatos.Fields.Item(1).Value.ToString().Replace(".", ","));
                        fila++;
                        oDatos.MoveNext();
                    }

                    //instanciamos el objeto para crear el asiento
                    SAPbobsCOM.JournalEntries oAsiento;
                    oAsiento = (SAPbobsCOM.JournalEntries)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    //mandamos las variables
                    //cabecera
                    oAsiento.Series = 23;
                    oAsiento.TaxDate = DateTime.Now;
                    oAsiento.DueDate = DateTime.Now;
                    oAsiento.ReferenceDate = DateTime.Now;
                    oAsiento.Memo = "";
                    oAsiento.UserFields.Fields.Item("U_CyC_DocNum").Value = txtDoc.Value;
                    //detalle
                    //debito
                    oAsiento.Lines.AccountCode = v_cuentaDeb;
                    oAsiento.Lines.Debit = v_monto;
                    oAsiento.Lines.Add();
                    //credito
                    oAsiento.Lines.AccountCode = v_cuentaCred;
                    oAsiento.Lines.Credit = v_monto;
                    oAsiento.Lines.Add();
                    int errorC = oAsiento.Add();
                    if (errorC != 0)
                    {
                        string errorDes = Menu.sbo.GetLastErrorDescription();
                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(errorDes, 1, "OK");
                    }
                    else
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Asiento creado correctamente!!", 1, "OK");
                    }

                }
            }

        }
    }
}