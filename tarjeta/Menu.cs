using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM.Framework;
using SAPbobsCOM;
using SAPbouiCOM;

namespace tarjeta
{
    class Menu
    {
        public static SAPbobsCOM.Company sbo = null;
        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;
            sbo = (SAPbobsCOM.Company)SAPbouiCOM.Framework.Application.SBO_Application.Company.GetDICompany();           

            oMenus = SAPbouiCOM.Framework.Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SAPbouiCOM.Framework.Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            //oMenuItem = SAPbouiCOM.Framework.Application.SBO_Application.Menus.Item("43537"); // moudles'

            //oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            //oCreationPackage.UniqueID = "tarjeta";
            //oCreationPackage.String = "tarjeta";
            //oCreationPackage.Enabled = true;
            //oCreationPackage.Position = -1;

            //oMenus = oMenuItem.SubMenus;

            //try
            //{
            //    //  If the manu already exists this code will fail
            //    oMenus.AddEx(oCreationPackage);
            //}
            //catch (Exception e)
            //{

            //}

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = SAPbouiCOM.Framework.Application.SBO_Application.Menus.Item("43537");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "ContraAsiento";
                oCreationPackage.String = "Generar Contra Asiento";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception er)
            { //  Menu already exists
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "tarjeta.Form1")
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.OpenForm(BoFormObjectEnum.fo_UserDefinedObject, "TARJETA", "");                                   
                }

                if(pVal.BeforeAction && pVal.MenuUID== "ContraAsiento")
                {
                    Form1 asiento = new Form1();
                    asiento.Show();
                }

            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

    }
}
