using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;
using System.Globalization;
using SAPbobsCOM;

namespace tarjeta
{
    [FormAttribute("UDO_FT_TARJETA")]
    class udoTarjeta : UDOFormBase
    {
        public static string v_documento = null;
        public bool v_grabar = false;
        public udoTarjeta()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            //               variables del UDO
            this.oMatrix = ((SAPbouiCOM.Matrix)(this.GetItem("0_U_G").Specific));
            this.oMatrix.DoubleClickAfter += new SAPbouiCOM._IMatrixEvents_DoubleClickAfterEventHandler(this.oMatrix_DoubleClickAfter);
            this.txtCode = ((SAPbouiCOM.EditText)(this.GetItem("0_U_E").Specific));
            this.txtFechaIni = ((SAPbouiCOM.EditText)(this.GetItem("13_U_E").Specific));
            this.txtFechaFin = ((SAPbouiCOM.EditText)(this.GetItem("14_U_E").Specific));
            this.cboTar = ((SAPbouiCOM.ComboBox)(this.GetItem("15_U_Cb").Specific));
            this.txtCuentaMa = ((SAPbouiCOM.EditText)(this.GetItem("16_U_E").Specific));
            this.txtMemo = ((SAPbouiCOM.EditText)(this.GetItem("17_U_E").Specific));
            this.cboSucu = ((SAPbouiCOM.ComboBox)(this.GetItem("18_U_Cb").Specific));
            this.cboSucu.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.cboSucu_ComboSelectAfter);
            this.txtCuentas = ((SAPbouiCOM.EditText)(this.GetItem("19_U_E").Specific));
            this.btnOK = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            //               variables agregados en el form
            this.btnAdd = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.btnAdd.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btnAdd_ClickAfter);
            this.btnListar = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.btnListar.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btnListar_ClickAfter);
            this.lkn1 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_2").Specific));
            this.btnExcel = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.btnExcel.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btnExcel_ClickAfter);
            this.lblProce = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.lblTotal = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.lblToExcel = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.lblProcExcel = ((SAPbouiCOM.StaticText)(this.GetItem("Item_9").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.txtAsiento = ((SAPbouiCOM.EditText)(this.GetItem("Item_7").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.UnloadAfter += new UnloadAfterHandler(this.Form_UnloadAfter);

        }

        #region VARIABLES
        private SAPbouiCOM.Button btnAdd;
        private SAPbouiCOM.Button btnListar;
        private SAPbouiCOM.LinkedButton lkn1;
        private SAPbouiCOM.Button btnExcel;
        private SAPbouiCOM.Button btnProcesar;
        private SAPbouiCOM.EditText txtFechaIni;
        private SAPbouiCOM.EditText txtFechaFin;
        private SAPbouiCOM.EditText txtCuentaMa;
        private SAPbouiCOM.EditText txtMemo;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.ComboBox cboTar;
        private SAPbouiCOM.ComboBox cboSucu;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.EditText txtCuentas;
        private SAPbouiCOM.StaticText lblTotal;
        private SAPbouiCOM.StaticText lblProce;
        private SAPbouiCOM.StaticText lblToExcel;
        private SAPbouiCOM.StaticText lblProcExcel;
        private SAPbouiCOM.EditText txtCode;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText txtAsiento;
        private SAPbouiCOM.StaticText lblcuenta;
        private SAPbouiCOM.Button btnOK;
        public string v_cuenta = null;
        #endregion


        private void OnCustomInitialize()
        {
            //texto del total
            lblTotal.Item.FontSize = 10;
            int colortxt = System.Drawing.Color.Red.ToArgb();
            lblTotal.Item.ForeColor = colortxt;
            lblProce.Item.ForeColor = colortxt;
            //texto de procesados
            lblProce.Item.FontSize = 10;
            int colortxt2 = System.Drawing.Color.Green.ToArgb();
            lblToExcel.Item.ForeColor = colortxt2;
            lblProcExcel.Item.ForeColor = colortxt2;
            //code inicial
            SAPbobsCOM.Recordset oCode;
            oCode = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oCode.DoQuery("SELECT COALESCE(MAX(\"DocEntry\"),0)+1 FROM \"@TARJETA\" ");
            txtCode.Value = oCode.Fields.Item(0).Value.ToString();
            txtFechaIni.Item.Click();
            txtCode.Item.Enabled = false;
            txtMemo.Value = "Compra POS TD -";
            btnOK.Item.Enabled = false;

        }

       
        
        //FUNCION PARA LISTAR LOS DATOS DE SAP Y CARGAR EN LA MATRIX
        private void btnListar_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {            
            try
            {
                v_grabar = false;
                //habilitamos el boton de excel
                btnExcel.Item.Enabled = true;               
                //verificamos la variable de fecha
                string v_fechaINI = txtFechaFin.Value;
                string v_fechFIN = txtFechaFin.Value;
                string v_feAsiento = txtAsiento.Value;
                if (string.IsNullOrEmpty(v_fechaINI) || string.IsNullOrEmpty(v_fechFIN) || string.IsNullOrEmpty(v_feAsiento))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Fecha de Inicio, Fin o Asiento no puede quedar vacío!!", 1, "OK");
                    return;
                }

                //verificamos el tipo de tarjeta seleccionada
                string v_tipo = null;
                string v_tarj = cboTar.Selected.Value;

                if (string.IsNullOrEmpty(txtCuentaMa.Value.ToString()))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("El campo de cuenta de BANCO no puede quedar vacío!!", 1, "OK");
                    return;
                }

                if (cboSucu.Selected.Value.ToString().Equals("Seleccionar"))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe seleccionar una sucursal", 1, "OK");
                    return;
                }

                if (v_tarj.Equals("Débito")) { v_tipo = "3"; }
                else { v_tipo = ""; }

                //realizmaos la consulta para traer los datos de las tarjetas
                SAPbobsCOM.Recordset oConsulta;
                oConsulta = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oConsulta.DoQuery("SELECT T2.\"DocEntry\",T0.\"VoucherNum\",T0.\"CreditSum\",T1.\"CardCode\",T1.\"CardName\" FROM RCT3 T0 " +
                                  "INNER JOIN ORCT T1 ON T0.\"DocNum\"=T1.\"DocNum\" "+
                                  "INNER JOIN OINV T2 ON T1.\"DocNum\"=T2.\"ReceiptNum\" " +
                                  "INNER JOIN OCRH T3 ON T0.\"VoucherNum\"=T3.\"VoucherNum\"  " +
                                  "WHERE T0.\"CrTypeCode\"=" + v_tipo + " AND T1.\"DocDate\" BETWEEN '" + txtFechaIni.Value + "' AND '" + txtFechaFin.Value + "' " +
                                  "AND T0.\"CreditAcct\" IN ('"+ cboSucu.Selected.Value + "') AND T3.\"Deposited\"='N' AND T2.\"CANCELED\"='N'   "+
                                  "GROUP BY T2.\"DocEntry\",T0.\"VoucherNum\",T0.\"CreditSum\",T1.\"CardCode\",T1.\"CardName\" ");

                SAPbouiCOM.DBDataSource source = oForm.DataSources.DBDataSources.Item("@TARJETADET");
                oMatrix.FlushToDataSource();
                source.Clear();
                int v_filaMatrix = 0;
                double v_total = 0;
                int v_rows = oConsulta.RecordCount;
                NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
                //recorremos
                while (!oConsulta.EoF)
                {
                    
                    //insertamos los datos en la matrix
                    string v_docnum = oConsulta.Fields.Item(0).Value.ToString();
                    string v_voucher = oConsulta.Fields.Item(1).Value.ToString();
                    string v_monto = oConsulta.Fields.Item(2).Value.ToString();
                    string v_cardcode = oConsulta.Fields.Item(3).Value.ToString();
                    string v_cardname = oConsulta.Fields.Item(4).Value.ToString();

                    source.InsertRecord(source.Size);
                    source.Offset = source.Size - 1;
                    source.SetValue("U_DocNum", v_filaMatrix, v_docnum);
                    source.SetValue("U_Voucher", v_filaMatrix, v_voucher);
                    source.SetValue("U_Monto", v_filaMatrix, v_monto);
                    source.SetValue("U_CodSN", v_filaMatrix, v_cardcode);
                    source.SetValue("U_NomSN", v_filaMatrix, v_cardname);
                    oMatrix.LoadFromDataSource();

                    oConsulta.MoveNext();
                    v_filaMatrix++;

                    //sumamos el total
                    v_total = v_total + double.Parse(v_monto);
                    lblTotal.Caption = "SAP: " + v_total.ToString("N",nfi);
                    lblProce.Caption = "Proc. SAP: " + v_filaMatrix.ToString() + "/" + v_rows.ToString();
                    

                }
                btnOK.Item.Enabled = true;
            }
            catch (Exception e)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(e.ToString(), 1, "OK");
            }

        }

        //PROCESO PARA ABRIR LA BUSQUEDA DE ARCHIVOS
        private void btnExcel_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //buscamos el archivo EXCEL
            Thread t = new Thread(() =>
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();

                DialogResult dr = openFileDialog.ShowDialog();
                if (dr == DialogResult.OK)
                {
                    string fileName = openFileDialog.FileName;
                    subirExcel(fileName);
                    // FILE.Value = fileName;
                    //SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(fileName);
                }
            });          // Kick off a new thread
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();

        }

        //FUNCION PARA SUBRI EXCEL
        private void subirExcel(string url)
        {
            //instaciamos los servicios de excel
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(url);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;
            xlWorkbook.Activate();
            
            int rowCount = xlRange.Rows.Count;
            int columnsCount = xlRange.Columns.Count;
            int colCount = xlRange.Columns.Count;
            // oProgresbar = SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.CreateProgressBar("Cargando", rowCount, true);
            int f = 1;
            int fin = rowCount;
            int exRows = rowCount;
            int filaExcel = 1;
            double v_totalExcel = 0;
            NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
            //recorremos el excel
            for (int i = 2; i <= rowCount; i++)
            {
                string v_voucherBanco = xlRange.Cells[i, 4].Value2.ToString();
                string v_montoBanco = xlRange.Cells[i, 6].Value2.ToString();
                v_voucherBanco = v_voucherBanco.Remove(0, 4);

                int matrixrow = this.oMatrix.RowCount;
                //recorremos la matrix
                int filamatrix = 1;               
                while (filamatrix <= matrixrow)
                {
                    SAPbouiCOM.EditText m_voucher = (SAPbouiCOM.EditText)this.oMatrix.Columns.Item(5).Cells.Item(filamatrix).Specific;
                    SAPbouiCOM.EditText m_monto = (SAPbouiCOM.EditText)this.oMatrix.Columns.Item(6).Cells.Item(filamatrix).Specific;
                    string v_voucher = m_voucher.Value;
                    string v_monto = m_monto.Value;

                    if (v_voucherBanco.Contains(v_voucher))
                    {
                        SAPbouiCOM.EditText M_VBANCO = (SAPbouiCOM.EditText)this.oMatrix.Columns.Item(7).Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.EditText M_MBANCO = (SAPbouiCOM.EditText)this.oMatrix.Columns.Item(8).Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.EditText M_DIF = (SAPbouiCOM.EditText)this.oMatrix.Columns.Item(9).Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.CheckBox M_CHECK = (SAPbouiCOM.CheckBox)this.oMatrix.Columns.Item(1).Cells.Item(filamatrix).Specific;
                        M_VBANCO.Value = v_voucherBanco;
                        M_MBANCO.Value = v_montoBanco;

                        double v_dife = double.Parse(v_monto.Replace(".", ",")) - double.Parse(v_montoBanco);
                        M_DIF.Value = v_dife.ToString();

                        if (v_dife == 0)
                        {
                            int color = Color.LightGreen.ToArgb();
                            this.oMatrix.CommonSetting.SetRowBackColor(filamatrix, color);
                            M_CHECK.Checked = true;
                            filamatrix = matrixrow + 1;
                        }
                        else
                        {
                            int color = Color.Blue.ToArgb();
                            this.oMatrix.CommonSetting.SetRowBackColor(filamatrix, color);
                            filamatrix = matrixrow + 1;
                        }
                        //actualizamos el contador y el total
                        v_totalExcel = v_totalExcel + double.Parse(v_montoBanco);
                        lblProcExcel.Caption = "Proc. Excel: " + filaExcel.ToString() + "/" + exRows.ToString() ;
                        lblToExcel.Caption = "EXCEL: " + v_totalExcel.ToString("N",nfi);
                        filaExcel++;
                    }
                    else
                    {
                        filamatrix++;
                    }
                }
                f++;
                //oProgresbar.Text = "Cargando Datos " + f.ToString() + "/" + fin.ToString();

            }
            //xlApp.Workbooks.Close();
            xlWorkbook.Close();
            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Procesado correctamente", 1, "OK");
            //oProgresbar.Text = "Proceso terminado...";
            //oProgresbar.Stop();
        }

        //FUNCION PARA AGREGAR CUENTA
        private void btnAdd_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //agarramos la sucursal seleccionada
            string v_cuenta = null;
            if (string.IsNullOrEmpty(txtCuentas.Value.ToString()))
            {
                 v_cuenta = "'" + cboSucu.Selected.Value + "'";
            }
            else
            {
                 v_cuenta = txtCuentas.Value.ToString() + ", " + "'" + cboSucu.Selected.Value + "'";
            }
            
            txtCuentas.Value = v_cuenta;
            txtMemo.Value = "Compra POS TD - " + cboSucu.Selected.Description.ToString();

        }

        //EVENTO PARA DESPLEGAR LA FACTURA
        private void oMatrix_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        { 
            if (pVal.ColUID == "C_0_2")
            {
                SAPbouiCOM.EditText oDoc = (SAPbouiCOM.EditText)oMatrix.Columns.Item(2).Cells.Item(pVal.Row).Specific;
                string v_doc = oDoc.Value;
                SAPbouiCOM.Framework.Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_Invoice, "", v_doc);
            }
        }
        
        //FUNCION PARA CREAR LOS ASIENTOS CONTABLES
        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (v_grabar == false)
            {
                try
                {
                    v_cuenta = cboSucu.Selected.Value.ToString();
                    //instanciamos el objeto para crear el asiento
                    SAPbobsCOM.JournalEntries oAsiento;
                    oAsiento = (SAPbobsCOM.JournalEntries)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                    SAPbobsCOM.SBObob objBridge = (SAPbobsCOM.SBObob)Menu.sbo.GetBusinessObject(BoObjectTypes.BoBridge);
                    DateTime v_feAsiento = Convert.ToDateTime(objBridge.Format_StringToDate(txtAsiento.Value).Fields.Item(0).Value);
                    //recorremos la matrix
                    int v_cant = oMatrix.RowCount;
                    int fila = 1;
                    int v_dep = 0;
                    while (fila <= v_cant)
                    {
                        SAPbouiCOM.CheckBox oCheck = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item(1).Cells.Item(fila).Specific;
                        SAPbouiCOM.EditText oDocnum = (SAPbouiCOM.EditText)oMatrix.Columns.Item(2).Cells.Item(fila).Specific;
                        SAPbouiCOM.EditText oVou = (SAPbouiCOM.EditText)oMatrix.Columns.Item(5).Cells.Item(fila).Specific;
                        SAPbouiCOM.EditText oMonto = (SAPbouiCOM.EditText)oMatrix.Columns.Item(6).Cells.Item(fila).Specific;

                        string v_vou = oVou.Value.ToString();

                        //consultamos si esta checkeado
                        bool v_check = oCheck.Checked;
                        if (v_check == true)
                        {
                            //deposito
                            SAPbobsCOM.CompanyService oService = Menu.sbo.GetCompanyService();
                            SAPbobsCOM.DepositsService dpService = (SAPbobsCOM.DepositsService)oService.GetBusinessService(SAPbobsCOM.ServiceTypes.DepositsService);
                            SAPbobsCOM.Deposit dpsAddMpesa = (SAPbobsCOM.Deposit)dpService.GetDataInterface(SAPbobsCOM.DepositsServiceDataInterfaces.dsDeposit);
                            SAPbobsCOM.DepositParams dpsParamAddMpesa;
                            SAPbobsCOM.CreditLines créditos = dpsAddMpesa.Credits;
                            SAPbobsCOM.CreditLine credit;
                            dpsAddMpesa.DepositType = SAPbobsCOM.BoDepositTypeEnum.dtCredit;
                            dpsAddMpesa.DepositDate = DateTime.Now;
                            //dpsAddMpesa.AllocationAccount = txtCuentaMa.Value;
                            dpsAddMpesa.DepositAccount = txtCuentaMa.Value;
                            dpsAddMpesa.VoucherAccount = txtCuentaMa.Value;
                            dpsAddMpesa.CommissionAccount = "6.01.01.001.018";
                            dpsAddMpesa.Commission = 0;// double.Parse(oMonto.Value.Replace(".",","));
                            dpsAddMpesa.CommissionDate = Convert.ToDateTime(objBridge.Format_StringToDate(txtFechaIni.Value).Fields.Item(0).Value);
                            dpsAddMpesa.DepositCurrency = "GS";
                            dpsAddMpesa.JournalRemarks = txtMemo.Value;
                            //dpsAddMpesa.BPLID = v_dep;
                            dpsAddMpesa.ReconcileAfterDeposit = SAPbobsCOM.BoYesNoEnum.tYES;
                            credit = dpsAddMpesa.Credits.Add();

                            SAPbobsCOM.Recordset oAbs;
                            oAbs = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            oAbs.DoQuery("SELECT * FROM OCRH  WHERE \"VoucherNum\"='" + v_vou + "'");
                            string v_absId = oAbs.Fields.Item(0).Value.ToString();

                            credit.AbsId = int.Parse(v_absId);
                            try
                            {
                                dpsParamAddMpesa = dpService.AddDeposit(dpsAddMpesa);
                                //actualizamos el campo en la factura
                                SAPbobsCOM.Documents oDocumento;
                                oDocumento = (SAPbobsCOM.Documents)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                                if (oDocumento.GetByKey(int.Parse(oDocnum.Value)))
                                {
                                    oDocumento.UserFields.Fields.Item("U_U_Destino").Value = "SI";
                                    oDocumento.Update();
                                }
                            }
                            catch (Exception e)
                            {
                                if (e.ToString().Contains("por otro usuario"))
                                {
                                    oCheck.Checked = false;
                                    int color = Color.Blue.ToArgb();
                                    this.oMatrix.CommonSetting.SetRowBackColor(fila, color);
                                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Depósito ya generado por otro usuario", 1, "OK");
                                }
                                else
                                {
                                    oCheck.Checked = false;
                                    int color = Color.Blue.ToArgb();
                                    this.oMatrix.CommonSetting.SetRowBackColor(fila, color);
                                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(e.ToString(), 1, "OK");

                                }

                            }



                        }
                        fila++;
                    }
                }
                catch (Exception e)
                {

                }
            }
           
            v_grabar = true;
        }

        //FUNCIONA AL SELECCIONAR UNA SUCURSAL
        private void cboSucu_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            txtMemo.Value = "Compra POS TD - " + cboSucu.Selected.Description.ToString();

            string v_cod = txtCode.Value;
            if (string.IsNullOrEmpty(v_cod))
            {
                //code inicial
                SAPbobsCOM.Recordset oCode;
                oCode = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oCode.DoQuery("SELECT COALESCE(MAX(\"DocEntry\"),0)+1 FROM \"@TARJETA\" ");
                txtCode.Value = oCode.Fields.Item(0).Value.ToString();
                txtMemo.Item.Click();
                txtCode.Item.Enabled = false;
            }

        }

        

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            

        }

        private void Form_UnloadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            
           

        }
    }
}
