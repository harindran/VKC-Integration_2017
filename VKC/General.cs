using System;
using System.Collections.Generic;
using System.Text;

namespace BranchIntegrator
{
    class General
    {
        public static SAPbobsCOM.Company oCompany;
        
        public static SAPbobsCOM.BusinessPartners oBusinessPartnersA;
        public static SAPbobsCOM.Items oItemsA;
        public static SAPbouiCOM.Application SapApplication;
        public static SAPbobsCOM.CompanyService oCompService;
        public string path = "C:\\Integrator\\SAPXML\\";
        public string pathBackup = "C:\\Integrator\\BackUpSAP\\";
        public string row = "row";
        public string row1 = "Data";
        public string pic = "tickMark.GIF";
        public static bool _bFlag = false;


        #region Connect To Ather Company
        internal bool connectOtherCompany(string Server, string LicServer, string CompanyDB, string SAPUser, string SAPPass, string SQLUser, string SQLPass)
        {
            try
            {
                string cookie, sErrorMsg;
                int iErrorCode = 0;
                string connStr;
                Global.oCompny2 = new SAPbobsCOM.Company();
              
                Global.oCompny2.Server = Server;
                Global.oCompny2.LicenseServer = LicServer + ":30000";
        Global.oCompny2.SLDServer = LicServer + ":40000";
                Global.oCompny2.CompanyDB = CompanyDB;
                Global.oCompny2.UserName = SAPUser;
                Global.oCompny2.UseTrusted = false;
                Global.oCompny2.Password = SAPPass;
                Global.oCompny2.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
                Global.oCompny2.DbUserName = SQLUser;
                Global.oCompny2.DbPassword = SQLPass;
                
                iErrorCode = Global.oCompny2.Connect();
                if (iErrorCode != 0)
                {
                    sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                    
                                                   
                }
                if (Global.oCompny2.Connected == true)
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch
            {
                return false;
                //SBO_Application.MessageBox(Global.oCompny2.GetLastErrorDescription().ToString(), 1, "Ok", "", "");
            }
        }
        #endregion
        
    }
}

