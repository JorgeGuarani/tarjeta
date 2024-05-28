using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Globalization;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using SAPbobsCOM;
using SAPbouiCOM;

namespace tarjeta
{
    [FormAttribute("UDO_FT_TARJETA")]
    class formTarjeta : UDOFormBase
    {
        public static string v_documento = null;
        public bool v_grabar = false;
        public static double G_totalSAP = 0;
        public static double G_totalExcel = 0;
        public static double G_comision = 0;
        public static string codigo = null;
        public static string cuenta = null;
        SAPbouiCOM.ProgressBar oProgress;
        public formTarjeta()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            this.oMatrixDeb = ((SAPbouiCOM.Matrix)(this.GetItem("0_U_G").Specific));
            this.oMatrixDeb.DoubleClickAfter += new SAPbouiCOM._IMatrixEvents_DoubleClickAfterEventHandler(this.oMatrixDeb_DoubleClickAfter);
            this.oMatrixCred1 = ((SAPbouiCOM.Matrix)(this.GetItem("1_U_G").Specific));
            this.oMatrixCred2 = ((SAPbouiCOM.Matrix)(this.GetItem("2_U_G").Specific));
            this.otxtCod = ((SAPbouiCOM.EditText)(this.GetItem("0_U_E").Specific));
            this.otxtDesde = ((SAPbouiCOM.EditText)(this.GetItem("13_U_E").Specific));
            this.otxtHasta = ((SAPbouiCOM.EditText)(this.GetItem("14_U_E").Specific));
            this.ocboTipo = ((SAPbouiCOM.ComboBox)(this.GetItem("15_U_Cb").Specific));
            this.ocboTipo.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ocboTipo_ComboSelectAfter);
            this.otxtBanco = ((SAPbouiCOM.EditText)(this.GetItem("16_U_E").Specific));
            this.otxtComen = ((SAPbouiCOM.EditText)(this.GetItem("17_U_E").Specific));
            this.ocboSucu = ((SAPbouiCOM.ComboBox)(this.GetItem("18_U_Cb").Specific));
            this.ocboSucu.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ocboSucu_ComboSelectAfter);
            this.otxtAsiento = ((SAPbouiCOM.EditText)(this.GetItem("20_U_E").Specific));
            this.btnGenerar = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.btnGenerar.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btnGenerar_ClickAfter);
            this.btnCancelar = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.btnListar = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.btnListar.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btnListar_ClickAfter);
            this.btnExcel = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.btnExcel.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btnExcel_ClickAfter);
            this.olblProcSap = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.olblProcExc = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.olblSAP = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.olblEXCEL = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.oFolderDeb = ((SAPbouiCOM.Folder)(this.GetItem("0_U_FD").Specific));
            this.oFolderCre = ((SAPbouiCOM.Folder)(this.GetItem("1_U_FD").Specific));
            this.otxtTotalSAP = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.otxtTotalExc = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.otxtComision = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.olkbBanco = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_9").Specific));
            this.oGridCred = ((SAPbouiCOM.Grid)(this.GetItem("Item_10").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("3_U_FD").Specific));
            this.btnControl = ((SAPbouiCOM.Button)(this.GetItem("Item_11").Specific));
            this.btnControl.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btnControl_ClickAfter);
            this.otxtTotalneto = ((SAPbouiCOM.StaticText)(this.GetItem("Item_12").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }


        private void OnCustomInitialize()
        {
            
            oMatrixDeb.AutoResizeColumns();
            oGridCred.AutoResizeColumns();
            oGridCred.DataTable.Rows.Add();
            oGridCred.Columns.Item("check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
            oFolderCre.Item.Enabled = false;
            //texto del total
            olblSAP.Item.FontSize = 10;
            int colortxt = System.Drawing.Color.Red.ToArgb();
            olblSAP.Item.ForeColor = colortxt;
            olblProcSap.Item.ForeColor = colortxt;
            //texto de procesados
            olblProcSap.Item.FontSize = 10;
            int colortxt2 = System.Drawing.Color.Green.ToArgb();
            olblEXCEL.Item.ForeColor = colortxt2;
            olblProcExc.Item.ForeColor = colortxt2;
            //code inicial
            SAPbobsCOM.Recordset oCode;
            oCode = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oCode.DoQuery("SELECT COALESCE(MAX(\"DocEntry\"),0)+1 FROM \"@TARJETA\" ");
            otxtCod.Value = oCode.Fields.Item(0).Value.ToString();
            otxtDesde.Item.Click();
            otxtCod.Item.Enabled = false;
            otxtComen.Value = "Compra POS TD -";
            btnGenerar.Item.Enabled = true;
            Folder0.Item.Visible = false;


        }

        #region VARIABLES
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Matrix oMatrixDeb;
        private SAPbouiCOM.Matrix oMatrixCred1;
        private SAPbouiCOM.Matrix oMatrixCred2;
        private SAPbouiCOM.ComboBox ocboTipo;
        private SAPbouiCOM.ComboBox ocboSucu;
        private SAPbouiCOM.EditText otxtDesde;
        private SAPbouiCOM.EditText otxtHasta;
        private SAPbouiCOM.EditText otxtAsiento;
        private SAPbouiCOM.EditText otxtBanco;
        private SAPbouiCOM.EditText otxtComen;
        private SAPbouiCOM.EditText otxtCod;
        private SAPbouiCOM.Button btnGenerar;
        private SAPbouiCOM.Button btnCancelar;
        private SAPbouiCOM.Button btnListar;
        private SAPbouiCOM.Button btnExcel;
        private SAPbouiCOM.StaticText olblProcSap;
        private SAPbouiCOM.StaticText olblProcExc;
        private SAPbouiCOM.StaticText olblSAP;
        private SAPbouiCOM.StaticText olblEXCEL;
        private SAPbouiCOM.Folder oFolderDeb;
        private SAPbouiCOM.Folder oFolderCre;
        private SAPbouiCOM.StaticText otxtTotalSAP;
        private SAPbouiCOM.StaticText otxtTotalExc;
        private SAPbouiCOM.StaticText otxtComision;
        private SAPbouiCOM.LinkedButton olkbBanco;
        private SAPbouiCOM.Grid oGridCred;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.Button btnControl;
        private SAPbouiCOM.StaticText otxtTotalneto;
        #endregion

        //FUNCION PARA SUBIR EXCEL
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

                int matrixrow = this.oMatrixDeb.RowCount;
                //recorremos la matrix
                int filamatrix = 1;
                while (filamatrix <= matrixrow)
                {
                    SAPbouiCOM.EditText m_voucher = (SAPbouiCOM.EditText)this.oMatrixDeb.Columns.Item(5).Cells.Item(filamatrix).Specific;
                    SAPbouiCOM.EditText m_monto = (SAPbouiCOM.EditText)this.oMatrixDeb.Columns.Item(6).Cells.Item(filamatrix).Specific;
                    string v_voucher = m_voucher.Value;
                    string v_monto = m_monto.Value;

                    if (v_voucherBanco.Contains(v_voucher))
                    {
                        SAPbouiCOM.EditText M_VBANCO = (SAPbouiCOM.EditText)this.oMatrixDeb.Columns.Item(7).Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.EditText M_MBANCO = (SAPbouiCOM.EditText)this.oMatrixDeb.Columns.Item(8).Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.EditText M_DIF = (SAPbouiCOM.EditText)this.oMatrixDeb.Columns.Item(9).Cells.Item(filamatrix).Specific;
                        SAPbouiCOM.CheckBox M_CHECK = (SAPbouiCOM.CheckBox)this.oMatrixDeb.Columns.Item(1).Cells.Item(filamatrix).Specific;
                        M_VBANCO.Value = v_voucherBanco;
                        M_MBANCO.Value = v_montoBanco;

                        double v_dife = double.Parse(v_monto.Replace(",", ".")) - double.Parse(v_montoBanco);
                        M_DIF.Value = v_dife.ToString();

                        if (v_dife == 0)
                        {
                            int color = Color.LightGreen.ToArgb();
                            this.oMatrixDeb.CommonSetting.SetRowBackColor(filamatrix, color);
                            M_CHECK.Checked = true;
                            filamatrix = matrixrow + 1;
                        }
                        else
                        {
                            int color = Color.Blue.ToArgb();
                            this.oMatrixDeb.CommonSetting.SetRowBackColor(filamatrix, color);
                            filamatrix = matrixrow + 1;
                        }
                        //actualizamos el contador y el total
                        v_totalExcel = v_totalExcel + double.Parse(v_montoBanco);
                        olblProcExc.Caption = "Proc. Excel: " + filaExcel.ToString() + "/" + exRows.ToString();
                        olblEXCEL.Caption = "EXCEL: " + v_totalExcel.ToString("N", nfi);
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

        //FUNCION EXCCEL CREDITO
        private void excelCredito(string url)
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
            int filaMa = 0;
            int fin = rowCount;
            int exRows = rowCount;
            NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
            SAPbouiCOM.DBDataSource source = oForm.DataSources.DBDataSources.Item("@TARJETADET3");
            oMatrixCred2.FlushToDataSource();
            source.Clear();
            int v_totalExc = 0;
            //recorremos el excel
            for (int i = 2; i <= rowCount; i++)
            {
                //cargar matrix de crédito
                string v_fecha = xlRange.Cells[i, 1].Value2.ToString();
                DateTime fecha_v = DateTime.Parse(v_fecha);
                string v_voucher = xlRange.Cells[i, 3].Value2.ToString();
                string v_monto = xlRange.Cells[i, 5].Value2.ToString();
                string v_montoNeto = xlRange.Cells[i, 6].Value2.ToString();

                source.InsertRecord(source.Size);
                source.Offset = source.Size - 1;
                source.SetValue("U_voucher", filaMa, v_voucher);
                source.SetValue("U_monto", filaMa, v_monto);
                oMatrixCred2.LoadFromDataSource();
                

                //cargar grilla de crédito
                oGridCred.DataTable.SetValue("voucher", filaMa, v_voucher);
                oGridCred.DataTable.SetValue("monto", filaMa, v_monto);
                oGridCred.DataTable.SetValue("fecha", filaMa, fecha_v.ToString("yyyyMMdd"));
                oGridCred.DataTable.SetValue("Monto neto", filaMa, v_montoNeto);
                oGridCred.DataTable.Rows.Add();
                v_totalExc = v_totalExc + int.Parse(v_monto);
                //otxtTotalExc.Caption = "Total Excel: " + v_totalExc.ToString("N", nfi);
                filaMa++;
            }
            //xlApp.Workbooks.Close();
            xlWorkbook.Close();
            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Procesado correctamente", 1, "OK");
        }

        //FUNCION PARA CREAR LOS ASIENTOS CONTABLES
        private void btnGenerar_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (v_grabar == false)
            {

                Program.codigo = otxtCod.Value;
                Program.cuenta = ocboSucu.Selected.Value.ToString();
                string v_tipo = ocboTipo.Selected.Value.ToString();
                if (v_tipo.Equals("Débito"))
                {
                    #region DEBITO
                    try
                    {
                        //instanciamos el objeto para crear el asiento
                        SAPbobsCOM.JournalEntries oAsiento;
                        oAsiento = (SAPbobsCOM.JournalEntries)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                        SAPbobsCOM.SBObob objBridge = (SAPbobsCOM.SBObob)Menu.sbo.GetBusinessObject(BoObjectTypes.BoBridge);
                        DateTime v_feAsiento = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                        //recorremos la matrix
                        int v_cant = oMatrixDeb.RowCount;
                        int fila = 1;
                        int v_dep = 0;
                        oProgress = SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.CreateProgressBar("Cargando datos...", v_cant, true);
                        while (fila <= v_cant)
                        {
                            SAPbouiCOM.CheckBox oCheck = (SAPbouiCOM.CheckBox)oMatrixDeb.Columns.Item(1).Cells.Item(fila).Specific;
                            SAPbouiCOM.EditText oDocnum = (SAPbouiCOM.EditText)oMatrixDeb.Columns.Item(2).Cells.Item(fila).Specific;
                            SAPbouiCOM.EditText oVou = (SAPbouiCOM.EditText)oMatrixDeb.Columns.Item(5).Cells.Item(fila).Specific;
                            SAPbouiCOM.EditText oMonto = (SAPbouiCOM.EditText)oMatrixDeb.Columns.Item(6).Cells.Item(fila).Specific;

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
                                dpsAddMpesa.DepositDate = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                                //dpsAddMpesa.AllocationAccount = txtCuentaMa.Value;
                                dpsAddMpesa.DepositAccount = otxtBanco.Value;
                                dpsAddMpesa.VoucherAccount = otxtBanco.Value;
                                dpsAddMpesa.CommissionAccount = "6.01.01.001.018";
                                dpsAddMpesa.Commission = 0;// double.Parse(oMonto.Value.Replace(".",","));
                                dpsAddMpesa.CommissionDate = Convert.ToDateTime(objBridge.Format_StringToDate(otxtDesde.Value).Fields.Item(0).Value);
                                dpsAddMpesa.DepositCurrency = "GS";
                                dpsAddMpesa.JournalRemarks = otxtComen.Value;
                                //dpsAddMpesa.BPLID = v_dep;
                                dpsAddMpesa.ReconcileAfterDeposit = SAPbobsCOM.BoYesNoEnum.tYES;
                                credit = dpsAddMpesa.Credits.Add();

                                SAPbobsCOM.Recordset oAbs;
                                oAbs = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                oAbs.DoQuery("SELECT * FROM OCRH  WHERE \"VoucherNum\"='" + v_vou + "' AND \"CreditSum\"=" + oMonto.Value.Replace(",", ".") + " ");
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
                                        this.oMatrixDeb.CommonSetting.SetRowBackColor(fila, color);
                                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Depósito ya generado por otro usuario", 1, "OK");
                                    }
                                    else
                                    {
                                        oCheck.Checked = false;
                                        int color = Color.Blue.ToArgb();
                                        this.oMatrixDeb.CommonSetting.SetRowBackColor(fila, color);
                                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(e.ToString(), 1, "OK");

                                    }

                                }
                            }
                            fila++;
                            oProgress.Value += 1;
                        }
                        oProgress.Stop();
                    }
                    catch (Exception e)
                    {

                    }
                    #endregion
                }
                else
                {
                    #region CREDITO
                    string v_sucursal = ocboSucu.Selected.Description;
                    if (v_sucursal.Contains("Bancard"))
                    {
                        #region BANCARD
                        SAPbobsCOM.SBObob objBridge = (SAPbobsCOM.SBObob)Menu.sbo.GetBusinessObject(BoObjectTypes.BoBridge);
                        //parametro de cabecera
                        string v_sucu = ocboSucu.Selected.Value.ToString();
                        string v_fecha = otxtAsiento.Value;
                        string v_cuentaBanco = otxtBanco.Value;
                        //generar deposito
                        SAPbobsCOM.CompanyService oService = Menu.sbo.GetCompanyService();
                        SAPbobsCOM.DepositsService dpService = (SAPbobsCOM.DepositsService)oService.GetBusinessService(SAPbobsCOM.ServiceTypes.DepositsService);
                        SAPbobsCOM.Deposit dpsAddCash = (SAPbobsCOM.Deposit)dpService.GetDataInterface(SAPbobsCOM.DepositsServiceDataInterfaces.dsDeposit);
                        dpsAddCash.DepositType = BoDepositTypeEnum.dtCash;
                        dpsAddCash.DepositDate = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                        dpsAddCash.DepositCurrency = "GS";
                        dpsAddCash.AllocationAccount = v_sucu;
                        dpsAddCash.DepositAccount = v_cuentaBanco;
                        dpsAddCash.TotalLC = G_totalExcel;
                        dpsAddCash.JournalRemarks = "Crédito a comercio " + ocboSucu.Selected.Description;
                        SAPbobsCOM.DepositParams dpsParamAddCash = dpService.AddDeposit(dpsAddCash);
                        //generamos el asiento para comision
                        SAPbobsCOM.JournalEntries oAsiento;
                        oAsiento = (SAPbobsCOM.JournalEntries)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                        oAsiento.Series = 23;
                        oAsiento.DueDate = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                        oAsiento.ReferenceDate = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                        oAsiento.TaxDate = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                        oAsiento.Lines.ShortName = "6.01.01.001.210";
                        oAsiento.Lines.Debit = G_comision;
                        oAsiento.Lines.Credit = 0;
                        oAsiento.Lines.DueDate = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                        oAsiento.Lines.TaxDate = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                        oAsiento.Lines.ReferenceDate1 = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                        oAsiento.Lines.LineMemo = "BANCARD-CAMP-COMISION.IVA.RENTA.";
                        oAsiento.Lines.Add();
                        oAsiento.Lines.AccountCode = v_sucu;
                        oAsiento.Lines.Debit = 0;
                        oAsiento.Lines.Credit = G_comision;
                        oAsiento.Lines.DueDate = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                        oAsiento.Lines.TaxDate = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                        oAsiento.Lines.ReferenceDate1 = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                        oAsiento.Lines.LineMemo = "BANCARD-CAMP-COMISION.IVA.RENTA.";
                        oAsiento.Lines.Add();
                        int v_error = oAsiento.Add();
                        //abrir el form de conciliacion
                        SAPbouiCOM.Framework.Application.SBO_Application.OpenForm(BoFormObjectEnum.fo_UserDefinedObject, "OFLT", "");
                        SAPbouiCOM.Form form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
                        reconciliacion();
                        #endregion
                    }

                    else
                    {
                        #region CABAL Y PANAL
                        //comentario cabal o panal
                        string v_cueNombre = ocboSucu.Selected.Description;
                        if (v_cueNombre.Contains("Cabal"))
                        {
                            v_cueNombre = "CABAL";
                        }
                        if (v_cueNombre.Contains("Panal"))
                        {
                            v_cueNombre = "PANAL";
                        }

                        SAPbobsCOM.SBObob objBridge = (SAPbobsCOM.SBObob)Menu.sbo.GetBusinessObject(BoObjectTypes.BoBridge);
                        //cantidad de filas
                        int v_cant = oGridCred.Rows.Count - 1;
                        int v_fila = 0;
                        while (v_fila < v_cant)
                        {
                            string v_monto = oGridCred.DataTable.GetValue("Monto neto", v_fila).ToString();
                            string v_montototal = oGridCred.DataTable.GetValue("monto", v_fila).ToString();
                            string v_fechaAsi = oGridCred.DataTable.GetValue("fecha", v_fila).ToString();
                            DateTime fechaasi_v = Convert.ToDateTime(v_fechaAsi);
                            string FECHA = fechaasi_v.ToString("yyyyMMdd");
                            double comi = double.Parse(v_montototal) - double.Parse(v_monto);

                            //parametro de cabecera
                            string v_sucu = ocboSucu.Selected.Value.ToString();
                            string v_fecha = otxtAsiento.Value;
                            string v_cuentaBanco = otxtBanco.Value;
                            //generar deposito
                            SAPbobsCOM.CompanyService oService = Menu.sbo.GetCompanyService();
                            SAPbobsCOM.DepositsService dpService = (SAPbobsCOM.DepositsService)oService.GetBusinessService(SAPbobsCOM.ServiceTypes.DepositsService);
                            SAPbobsCOM.Deposit dpsAddCash = (SAPbobsCOM.Deposit)dpService.GetDataInterface(SAPbobsCOM.DepositsServiceDataInterfaces.dsDeposit);
                            dpsAddCash.DepositType = BoDepositTypeEnum.dtCash;
                            dpsAddCash.DepositDate = Convert.ToDateTime(objBridge.Format_StringToDate(FECHA).Fields.Item(0).Value);
                            dpsAddCash.DepositCurrency = "GS";
                            dpsAddCash.AllocationAccount = v_sucu;
                            dpsAddCash.DepositAccount = v_cuentaBanco;
                            dpsAddCash.TotalLC = double.Parse(v_monto);
                            dpsAddCash.JournalRemarks = "Crédito a comercio " + ocboSucu.Selected.Description;
                            SAPbobsCOM.DepositParams dpsParamAddCash = dpService.AddDeposit(dpsAddCash);
                            //generamos el asiento para comision
                            string v_sucur = ocboSucu.Selected.Value.ToString();
                            SAPbobsCOM.JournalEntries oAsiento;
                            oAsiento = (SAPbobsCOM.JournalEntries)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                            oAsiento.Series = 23;
                            oAsiento.Reference = v_cueNombre + "-" + ocboSucu.Selected.Description + "-COMISION.IVA.RENTA.";
                            oAsiento.Reference2 = v_cueNombre + "-" + ocboSucu.Selected.Description + "-COMISION.IVA.RENTA.";
                            oAsiento.Memo = v_cueNombre + "-" + ocboSucu.Selected.Description + "-COMISION.IVA.RENTA.";
                            oAsiento.DueDate = Convert.ToDateTime(objBridge.Format_StringToDate(FECHA).Fields.Item(0).Value);
                            oAsiento.ReferenceDate = Convert.ToDateTime(objBridge.Format_StringToDate(FECHA).Fields.Item(0).Value);
                            oAsiento.TaxDate = Convert.ToDateTime(objBridge.Format_StringToDate(FECHA).Fields.Item(0).Value);
                            oAsiento.Lines.ShortName = "6.01.01.001.210";
                            oAsiento.Lines.Debit = comi;
                            oAsiento.Lines.Credit = 0;
                            oAsiento.Lines.DueDate = Convert.ToDateTime(objBridge.Format_StringToDate(FECHA).Fields.Item(0).Value);
                            oAsiento.Lines.TaxDate = Convert.ToDateTime(objBridge.Format_StringToDate(FECHA).Fields.Item(0).Value);
                            oAsiento.Lines.ReferenceDate1 = Convert.ToDateTime(objBridge.Format_StringToDate(FECHA).Fields.Item(0).Value);
                            oAsiento.Lines.LineMemo = v_cueNombre + "-" + ocboSucu.Selected.Description + "-COMISION.IVA.RENTA.";
                            oAsiento.Lines.Add();
                            oAsiento.Lines.AccountCode = v_sucur;
                            oAsiento.Lines.Debit = 0;
                            oAsiento.Lines.Credit = comi;
                            oAsiento.Lines.DueDate = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                            oAsiento.Lines.TaxDate = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                            oAsiento.Lines.ReferenceDate1 = Convert.ToDateTime(objBridge.Format_StringToDate(otxtAsiento.Value).Fields.Item(0).Value);
                            oAsiento.Lines.LineMemo = v_cueNombre + "-" + ocboSucu.Selected.Description + "-COMISION.IVA.RENTA.";
                            oAsiento.Lines.Add();
                            int v_error = oAsiento.Add();
                            v_fila++;
                        }
                        SAPbouiCOM.Framework.Application.SBO_Application.OpenForm(BoFormObjectEnum.fo_UserDefinedObject, "OFLT", "");
                        SAPbouiCOM.Form form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
                        reconciliacion();

                        #endregion
                    }

                    #endregion
                }

            }
            //setemos las variables
            G_comision = 0;
            G_totalExcel = 0;
            G_totalSAP = 0;
            v_grabar = true;

        }

        //FUNCION AL SELECCIONAR UNA SUCURSAL
        private void ocboSucu_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string v_tipo = ocboTipo.Selected.Value;
            if (v_tipo.Equals("Débito"))
            {
                otxtComen.Value = "Compra POS TD - " + ocboSucu.Selected.Description.ToString();
            }
            else
            {
                otxtComen.Value = "Compra POS TC - " + ocboSucu.Selected.Description.ToString();
            }
            
            //code inicial
            SAPbobsCOM.Recordset oCode;
            oCode = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oCode.DoQuery("SELECT COALESCE(MAX(\"DocEntry\"),0)+1 FROM \"@TARJETA\" ");
            otxtCod.Item.Enabled = true;
            otxtCod.Value = oCode.Fields.Item(0).Value.ToString();
            otxtComen.Item.Click();
            otxtCod.Item.Enabled = false;

        }

        //EVENTO PARA DESPLEGAR LA FACTURA
        private void oMatrixDeb_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.ColUID == "C_0_2")
            {
                SAPbouiCOM.EditText oDoc = (SAPbouiCOM.EditText)oMatrixDeb.Columns.Item(2).Cells.Item(pVal.Row).Specific;
                string v_doc = oDoc.Value;
                SAPbouiCOM.Framework.Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_Invoice, "", v_doc);
            }

        }

        //SELECCIONAMOS EL FOLDER
        private void ocboTipo_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string v_fol = ocboTipo.Selected.Value.ToString();
            if (v_fol.Equals("Débito"))
            {
                oFolderDeb.Item.Enabled = true;
                oFolderDeb.Item.Click();
                oFolderCre.Item.Enabled = false;
            }
            else
            {
                oFolderCre.Item.Enabled = true;
                oFolderCre.Item.Click();
                oFolderDeb.Item.Enabled = false;
            }
        }

        //FUNCION PARA LISTAR LOS DATOS DE SAP Y CARGAR EN LA MATRIX
        private void btnListar_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                v_grabar = false;
                string v_tarj = null;
                string v_tipoT = ocboTipo.Selected.Value.ToString();
                if (v_tipoT.Equals("Débito"))
                {
                    #region DEBITO
                    v_tarj = "3";
                    //habilitamos el boton de excel
                    btnExcel.Item.Enabled = true;
                    //verificamos la variable de fecha
                    string v_fechaINI = otxtDesde.Value;
                    string v_fechFIN = otxtHasta.Value;
                    string v_feAsiento = otxtAsiento.Value;
                    if (string.IsNullOrEmpty(v_fechaINI) || string.IsNullOrEmpty(v_fechFIN) || string.IsNullOrEmpty(v_feAsiento))
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Fecha de Inicio, Fin o Asiento no puede quedar vacío!!", 1, "OK");
                        return;
                    }

                    if (string.IsNullOrEmpty(otxtBanco.Value.ToString()))
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("El campo de cuenta de BANCO no puede quedar vacío!!", 1, "OK");
                        return;
                    }

                    if (ocboSucu.Selected.Value.ToString().Equals("Seleccionar"))
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe seleccionar una sucursal", 1, "OK");
                        return;
                    }

                    //realizmaos la consulta para traer los datos de las tarjetas
                    SAPbobsCOM.Recordset oConsulta;
                    oConsulta = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oConsulta.DoQuery("call \"CT_NoDepositados\"('" + otxtDesde.Value + "','" + otxtHasta.Value + "','" + ocboSucu.Selected.Value + "','"+ v_tarj + "')");

                    SAPbouiCOM.DBDataSource source = oForm.DataSources.DBDataSources.Item("@TARJETADET");
                    oMatrixDeb.FlushToDataSource();
                    source.Clear();
                    int v_filaMatrix = 0;
                    int filaMa = 1;
                    double v_total = 0;
                    int v_rows = oConsulta.RecordCount;
                    NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
                    oProgress = SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.CreateProgressBar("Cargando datos...", v_rows, true);
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
                        oMatrixDeb.LoadFromDataSource();
                        int color = Color.White.ToArgb();
                        this.oMatrixDeb.CommonSetting.SetRowBackColor(filaMa, color);
                        oConsulta.MoveNext();
                        v_filaMatrix++;
                        filaMa++;

                        //sumamos el total
                        v_total = v_total + double.Parse(v_monto);
                        olblSAP.Caption = "SAP: " + v_total.ToString("N", nfi);
                        olblProcSap.Caption = "Proc. SAP: " + v_filaMatrix.ToString() + "/" + v_rows.ToString();
                        //oProgress.Value += 1;

                    }
                    btnGenerar.Item.Enabled = true;
                    oProgress.Stop();
                    #endregion
                }
                else
                {
                    #region CREDITO
                    v_tarj = "2";
                    //habilitamos el boton de excel
                    btnExcel.Item.Enabled = true;
                    //verificamos la variable de fecha
                    string v_fechaINI = otxtDesde.Value;
                    string v_fechFIN = otxtHasta.Value;
                    string v_feAsiento = otxtAsiento.Value;
                    if (string.IsNullOrEmpty(v_fechaINI) || string.IsNullOrEmpty(v_fechFIN) || string.IsNullOrEmpty(v_feAsiento))
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Fecha de Inicio, Fin o Asiento no puede quedar vacío!!", 1, "OK");
                        return;
                    }

                    if (string.IsNullOrEmpty(otxtBanco.Value.ToString()))
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("El campo de cuenta de BANCO no puede quedar vacío!!", 1, "OK");
                        return;
                    }

                    if (ocboSucu.Selected.Value.ToString().Equals("Seleccionar"))
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe seleccionar una sucursal", 1, "OK");
                        return;
                    }

                    //realizmaos la consulta para traer los datos de las tarjetas
                    SAPbobsCOM.Recordset oConsulta;
                    oConsulta = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oConsulta.DoQuery("call \"CT_NoDepositados\"('" + otxtDesde.Value + "','" + otxtHasta.Value + "','" + ocboSucu.Selected.Value + "','"+ v_tarj + "')");

                    SAPbouiCOM.DBDataSource source = oForm.DataSources.DBDataSources.Item("@TARJETADET2");
                    oMatrixCred1.FlushToDataSource();
                    source.Clear();
                    int v_filaMatrix = 0;
                    int filaMa = 1;
                    double v_total = 0;
                    int v_rows = oConsulta.RecordCount;
                    NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
                    oProgress = SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.CreateProgressBar("Cargando datos...", v_rows, true);
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
                        source.SetValue("U_voucher", v_filaMatrix, v_voucher);
                        source.SetValue("U_monto", v_filaMatrix, v_monto);
                        source.SetValue("U_CardCode", v_filaMatrix, v_cardcode);
                        source.SetValue("U_CardName", v_filaMatrix, v_cardname);
                        oMatrixCred1.LoadFromDataSource();
                        int color = Color.White.ToArgb();
                        this.oMatrixCred1.CommonSetting.SetRowBackColor(filaMa, color);
                        oConsulta.MoveNext();
                        v_filaMatrix++;
                        filaMa++;

                        //sumamos el total
                        //v_total = v_total + double.Parse(v_monto);
                        //olblSAP.Caption = "SAP: " + v_total.ToString("N", nfi);
                        //olblProcSap.Caption = "Proc. SAP: " + v_filaMatrix.ToString() + "/" + v_rows.ToString();
                        //oProgress.Value += 1;

                    }
                    btnGenerar.Item.Enabled = true;
                    oProgress.Stop();
                    #endregion
                }              
                
            }
            catch (Exception e)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(e.ToString(), 1, "OK");
            }

        }

        //PROCESO PARA ABRIR LA BUSQUEDA DE ARCHIVOS
        private void btnExcel_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string v_tipo = ocboTipo.Selected.Value.ToString();
            //buscamos el archivo EXCEL
            try
            {
                Thread t = new Thread(() =>
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    DialogResult dr = openFileDialog.ShowDialog();
                    if (dr == DialogResult.OK)
                    {
                        string fileName = openFileDialog.FileName;
                        if (v_tipo.Equals("Débito"))
                        {
                            subirExcel(fileName);
                        }
                        else
                        {
                            excelCredito(fileName);
                        }
                    }
                });
                // Kick off a new thread
                t.IsBackground = true;
                t.SetApartmentState(ApartmentState.STA);
                t.Start();

                //prueba para el excel por dispatcher
                //string v_usu = Menu.sbo.UserName.ToString();
                //SAPbobsCOM.Recordset ousu;
                //ousu = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
                //ousu.DoQuery("");

                //string v_ruta = "C:\\Users\\jorge.chaparro\\OneDrive - Grupo Pettengill\\Documentos\\tarjeta\\Template-tarjeta.xlsx";
                //string v_rutaCred = "C:\\Users\\jorge.chaparro\\OneDrive - Grupo Pettengill\\Documentos\\tarjeta\\Template-credito.xlsx";
                //if (v_tipo.Equals("Débito"))
                //{
                //    subirExcel(v_ruta);
                //}
                //else
                //{
                //    excelCredito(v_rutaCred);
                //}
            }
            catch (Exception e)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(e.ToString(), 1, "OK");
            }

        }

        //CONTROL DE GRILLAS
        private void btnControl_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
            int gridSap = oMatrixCred1.RowCount;
            int gridExc = oGridCred.Rows.Count - 1;
            int cantidad = 1;
            int filamatrix = 1;
            int conGrilla = 0;
            int totalSAP = 0;
            int totalExcel = 0;
            int totalneto = 0;
            int comision = 0;

            while (conGrilla < gridExc)
            {
                //voucher del banco
                string v_voucherExcel = oGridCred.DataTable.GetValue("voucher", conGrilla).ToString();
                int conMatrix = 1;
                while (conMatrix <= gridSap)
                {
                    SAPbouiCOM.EditText m_voucher = (SAPbouiCOM.EditText)this.oMatrixCred1.Columns.Item(5).Cells.Item(conMatrix).Specific;
                    SAPbouiCOM.EditText m_total = (SAPbouiCOM.EditText)this.oMatrixCred1.Columns.Item(6).Cells.Item(conMatrix).Specific;
                    SAPbouiCOM.CheckBox M_CHECK = (SAPbouiCOM.CheckBox)this.oMatrixCred1.Columns.Item(1).Cells.Item(conMatrix).Specific;
                    string v_voucher = m_voucher.Value;
                    if (v_voucher.Contains(v_voucherExcel))
                    {
                        int color = Color.LightGreen.ToArgb();
                        this.oMatrixCred1.CommonSetting.SetRowBackColor(conMatrix, color);
                        oGridCred.CommonSetting.SetRowBackColor(cantidad, color);
                        M_CHECK.Checked = true;
                        oGridCred.DataTable.SetValue("check", conGrilla, "Y");
                        string v_totalSAP = m_total.Value;
                        char valor = '.';
                        int v_punto = v_totalSAP.IndexOf(valor);
                        v_totalSAP = v_totalSAP.Remove(v_punto);
                        int v_montoSAP = int.Parse(v_totalSAP);
                        totalSAP = totalSAP + v_montoSAP;

                        string v_totalExcel = oGridCred.DataTable.GetValue("monto", conGrilla).ToString();
                        string v_totalExcelNeto = oGridCred.DataTable.GetValue("Monto neto", conGrilla).ToString();

                        totalExcel = totalExcel + int.Parse(v_totalExcel);
                        totalneto = totalneto + int.Parse(v_totalExcelNeto);
                        break;
                    }
                    conMatrix++;
                }

                conGrilla++;
                cantidad++;
            }
            comision = totalExcel - totalneto;
            otxtTotalExc.Caption = "Total Excel: " + totalExcel.ToString("N", nfi);
            otxtTotalSAP.Caption = "Total SAP: " + totalSAP.ToString("N", nfi); ;
            otxtComision.Caption = "Comisión: " + comision.ToString("N", nfi); ;
            otxtTotalneto.Caption = "Total neto: " + totalneto.ToString("N", nfi);
            //mandamos a las variables globales
            G_totalExcel = totalneto;
            G_totalSAP = totalExcel;
            G_comision = comision;

        }

        //funcion para machear las reconcilaciones internas
        public static void reconciliacion()
        {
            SAPbouiCOM.Framework.Application.SBO_Application.ActivateMenuItem("9461");

            SAPbouiCOM.Form form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            SAPbouiCOM.EditText v_cue = (SAPbouiCOM.EditText)form.Items.Item("120000072").Specific;
            SAPbouiCOM.Button v_boton = (SAPbouiCOM.Button)form.Items.Item("120000001").Specific;
            v_cue.Value = cuenta;
            v_boton.Item.Click();
            //query para machear

            SAPbobsCOM.Recordset oquery;
            oquery = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(BoObjectTypes.BoRecordset);
            oquery.DoQuery("SELECT * FROM \"@TARJETADET2\" WHERE \"Code\"='"+codigo+"' AND \"U_check\"='Y'");
            //oquery.DoQuery("call \"CT_NoDepositados\"('20240217','20240217','1.01.01.003.011','2')");
            while (!oquery.EoF)
            {
                //monto
                string monto = oquery.Fields.Item(9).Value.ToString();
                double montoQue = double.Parse(monto);

                SAPbouiCOM.Form formCon = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)formCon.Items.Item("120000039").Specific;
                int filas = oMatrix.RowCount;
                int v_reco = 1;
                while (v_reco <= filas)
                {
                    SAPbouiCOM.EditText oMon = (SAPbouiCOM.EditText)oMatrix.Columns.Item("120000014").Cells.Item(v_reco).Specific;
                    SAPbouiCOM.CheckBox ocheck = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item(1).Cells.Item(v_reco).Specific;
                    string v_monto = oMon.Value.Replace(" GS", string.Empty).Replace("(", string.Empty).Replace(")", string.Empty);
                    double montoGri = double.Parse(v_monto);
                    if (montoGri == montoQue)
                    {
                        ocheck.Checked = true;
                        break;
                    }
                    v_reco++;
                }
                oquery.MoveNext();
            }
            //seteamos las varaibles globales
            cuenta = "";
            codigo = "";
        }
    }
}
