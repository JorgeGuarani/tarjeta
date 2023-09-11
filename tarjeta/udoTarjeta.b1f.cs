using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;

namespace tarjeta
{
    [FormAttribute("UDO_FT_TARJETA")]
    class udoTarjeta : UDOFormBase
    {
        public udoTarjeta()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            //    variables del UDO
            this.oMatrix = ((SAPbouiCOM.Matrix)(this.GetItem("0_U_G").Specific));
            this.txtFechaIni = ((SAPbouiCOM.EditText)(this.GetItem("13_U_E").Specific));
            this.txtFechaFin = ((SAPbouiCOM.EditText)(this.GetItem("14_U_E").Specific));
            this.cboTar = ((SAPbouiCOM.ComboBox)(this.GetItem("15_U_Cb").Specific));
            this.txtCuentaMa = ((SAPbouiCOM.EditText)(this.GetItem("16_U_E").Specific));
            this.txtCuentaPa = ((SAPbouiCOM.EditText)(this.GetItem("17_U_E").Specific));
            this.cboSucu = ((SAPbouiCOM.ComboBox)(this.GetItem("18_U_Cb").Specific));
            this.txtCuentas = ((SAPbouiCOM.EditText)(this.GetItem("19_U_E").Specific));
            //    variables agregados en el form
            this.btnAdd = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.btnAdd.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btnAdd_ClickAfter);
            this.btnListar = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.btnListar.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btnListar_ClickAfter);
            this.lkn1 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_2").Specific));
            this.lkn2 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_3").Specific));
            this.btnExcel = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.btnExcel.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.btnExcel_ClickAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        #region VARIABLES
        private SAPbouiCOM.Button btnAdd;
        private SAPbouiCOM.Button btnListar;
        private SAPbouiCOM.LinkedButton lkn1;
        private SAPbouiCOM.LinkedButton lkn2;
        private SAPbouiCOM.Button btnExcel;
        private SAPbouiCOM.Button btnProcesar;
        private SAPbouiCOM.EditText txtFechaIni;
        private SAPbouiCOM.EditText txtFechaFin;
        private SAPbouiCOM.EditText txtCuentaMa;
        private SAPbouiCOM.EditText txtCuentaPa;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.ComboBox cboTar;
        private SAPbouiCOM.ComboBox cboSucu;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.EditText txtCuentas;
        #endregion


        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;

        private void btnListar_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                string v_fechaINI = txtFechaFin.Value;
                string v_fechFIN = txtFechaFin.Value;
                if (string.IsNullOrEmpty(v_fechaINI) || string.IsNullOrEmpty(v_fechFIN))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Fecha de inicio o fin no pueden quedar vacío!!", 1, "OK");
                    return;
                }

                //verificamos el tipo de tarjeta seleccionada
                string v_tipo = null;
                string v_tarj = cboTar.Selected.Value;
                string v_cuenta = txtCuentas.Value.ToString();

                if (v_tarj.Equals("Débito")) { v_tipo = "3"; }
                else { v_tipo = ""; }

                //realizmaos la consulta para traer los datos de las tarjetas
                SAPbobsCOM.Recordset oConsulta;
                oConsulta = (SAPbobsCOM.Recordset)Menu.sbo.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oConsulta.DoQuery("SELECT T0.\"DocNum\",T0.\"VoucherNum\",T0.\"CreditSum\",T1.\"CardCode\",T1.\"CardName\" FROM \"FG_PROD\".RCT3 T0 " +
                                  "INNER JOIN \"FG_PROD\".ORCT T1 ON T0.\"DocNum\"=T1.\"DocNum\" "+
                                  "WHERE T0.\"CrTypeCode\"=" + v_tipo + " AND T1.\"DocDate\" BETWEEN '" + txtFechaIni.Value + "' AND '" + txtFechaFin.Value + "' " +
                                  "AND T0.\"CreditAcct\" IN ("+ v_cuenta + ") ");

                SAPbouiCOM.DBDataSource source = oForm.DataSources.DBDataSources.Item("@TARJETADET");
                oMatrix.FlushToDataSource();
                source.Clear();
                int v_filaMatrix = 0;
                int v_canInicio = 1;
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

                }
            }
            catch (Exception e)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(e.ToString(), 1, "OK");
            }

        }

        //proceso para abrir el buscador de archivos
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

        //funcion para buscar excel
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
                    }
                    else
                    {
                        filamatrix++;
                    }
                }
                f++;
                //oProgresbar.Text = "Cargando Datos " + f.ToString() + "/" + fin.ToString();
            }
            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Procesado correctamente", 1, "OK");
            //oProgresbar.Text = "Proceso terminado...";
            //oProgresbar.Stop();
        }

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

        }
    }
}
