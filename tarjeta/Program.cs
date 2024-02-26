using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;
using SAPbouiCOM;
using SAPbobsCOM;

namespace tarjeta
{
    class Program
    {
        public static string cuenta = null;
        public static string codigo = null;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                SAPbouiCOM.Framework.Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new SAPbouiCOM.Framework.Application();
                }
                else
                {
                    oApp = new SAPbouiCOM.Framework.Application(args[0]);
                }
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                SAPbouiCOM.Framework.Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                SAPbouiCOM.Framework.Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }

        public static void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if(pVal.FormTypeEx == "UDO_FT_TARJETA")
                {
                    if(pVal.ItemUID == "1" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.FormMode == 3)
                    {
                        reconciliacion();
                    }
                }
            }
            catch (Exception e)
            {

            }
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
            oquery.DoQuery("SELECT * FROM \"@TARJETADET2\" WHERE \"Code\"='" + codigo + "' AND \"U_check\"='Y'");
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
