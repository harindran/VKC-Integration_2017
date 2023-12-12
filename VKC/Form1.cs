using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Data;
using System.IO;
using System.Threading;
using System.Data.SqlClient;

namespace BranchIntegrator
{
    public partial class Form1 : Form
    {
        DbConnection ConDb = new DbConnection();
        SqlDataReader objReaderSAP = null;
        string msg = "";
        string StrSql = "", LoggedBranch = "";
        string _str_sql = "";
        string _str_update = "";
        string _str_ItemCode = "";
        private SAPbobsCOM.Company oCompany_2;

        void ConnectCompany()
        {
            try
            {
                StreamReader sr = new StreamReader(Application.StartupPath + "\\Profile.ini");
                string[] s = sr.ReadToEnd().Replace("\r\n\t", "\r").Split('\r');
                General.oCompany = new SAPbobsCOM.Company();
                General.oCompany.Server = s[0].Split(':')[1].TrimEnd();
                General.oCompany.LicenseServer = s[1].Split(':')[1].TrimEnd() + ":30000"; ;
                General.oCompany.SLDServer = s[1].Split(':')[1].TrimEnd() + ":40000"; ;
                General.oCompany.CompanyDB = s[2].Split(':')[1].TrimEnd();
                General.oCompany.UserName = s[3].Split(':')[1].TrimEnd();
                General.oCompany.Password = s[4].Split(':')[1].TrimEnd();
                General.oCompany.DbUserName = s[5].Split(':')[1].TrimEnd();
                General.oCompany.DbPassword = s[6].Split(':')[1].TrimEnd();
                General.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
                //General.oCompany.LicenseServer = s[0].Split(':')[1].TrimEnd() + ":40000";


                //General.oCompany.Server = "192.168.1.100"; //objReaderSap[0].ToString();

                ////General.oCompany.Server = "vibin"; //objReaderSap[0].ToString();
                //General.oCompany.CompanyDB = "VKCGROUPSOPKERALA(D)"; //objReaderSap[1].ToString();
                //General.oCompany.UserName = "manager"; //objReaderSap[4].ToString();
                //General.oCompany.Password = "vkc.in"; //objReaderSap[5].ToString();
                //General.oCompany.DbUserName = "sa"; //objReaderSap[2].ToString();
                //General.oCompany.DbPassword = "sapb1";// objReaderSap[3].ToString();
                //General.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
                
                int iErrorCode = General.oCompany.Connect();

                if (iErrorCode != 0)
                {
                    string sErrorMsg = General.oCompany.GetLastErrorDescription();
                    MessageBox.Show("Error in SAP Connection" + sErrorMsg);
                }
                else
                {
                    MessageBox.Show("Company Connected Successfully!!!");
                }
            }
            catch (Exception Ex)
                            {
                                MessageBox.Show("Company Connection Failed"+Ex); 
                                }
        }
        
        public Form1()
        {
            InitializeComponent();
        }

        public Form1(SAPbobsCOM.Company oCompany_2)
        {
            // TODO: Complete member initialization
            this.oCompany_2 = oCompany_2;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                button1.Enabled = false;
                General g = new General();
                FolderManager fm = new FolderManager();
                try
                {
                    ConnectCompany();
                }
                catch (Exception Ex)
                {
                    MessageBox.Show("Error In DB Connection"+Ex);   
                }

                string str = comboBox1.Text.Trim();
                switch (str)
                {
                    case "":
                        MessageBox.Show("Please select your operation");
                        Cursor.Current = Cursors.Default;
                        button1.Enabled = true;
                        break;
                    case "ITEMGROUP":
                        ITEMGROUP();
                        //button1.BackColor = System.Drawing.Color.White;
                        MessageBox.Show("Item Group Data's Imported Sucessfully");
                        Cursor.Current = Cursors.Default;
                        button1.Enabled = true;
                        GC.Collect();
                        break;

                    case "ITEMMASTER":
                        New_ITEMMASTER();
                        //ITEMMASTER();
                        //button1.BackColor = System.Drawing.Color.White;
                        MessageBox.Show("Item Master Data's Imported Sucessfully");
                        Cursor.Current = Cursors.Default;
                        button1.Enabled = true;
                        GC.Collect();
                        break;
                    case "CUSTOMER":
                        New_sCustomer();
                        //sCustomer();
                        //button1.BackColor = System.Drawing.Color.White;
                        MessageBox.Show("Customer Data's Imported Sucessfully");
                        Cursor.Current = Cursors.Default;
                        button1.Enabled = true;
                        GC.Collect();
                        break;
                    case "VENDOR":
                        sVendor();
                        //button1.BackColor = System.Drawing.Color.White;
                        //MessageBox.Show("Vendor Data's Imported Sucessfully");
                        Cursor.Current = Cursors.Default;
                        button1.Enabled = true;
                        GC.Collect();
                        break;

                    case "ALL":
                        ITEMGROUP();
                        //BARCODE();
                        //  WareHouse();
                        ITEMMASTER();
                        // PRICELIST();
                        //     //TAX();
                        ///////////// SALESEMP();
                        sCustomer();

                        sVendor();

                        // Location();

                        // 

                        // ACCOUNT();
                        //stock();
                        //button1.BackColor = System.Drawing.Color.White;
                        MessageBox.Show("All Master Data's Imported Sucessfully");
                        button1.Enabled = true;
                           Cursor.Current = Cursors.Default;
                           GC.Collect();
                        break;

                }
                //  ITEMGROUP();
                //       //BARCODE();
                //     //  WareHouse();
                //       ITEMMASTER();
                //// PRICELIST();
                //  //     //TAX();
                //    ///////////// SALESEMP();
                //  sCustomer();

                //  sVendor();

                //   // Location();

                //    // 

                //      // ACCOUNT();
                //       //stock();
                //       //button1.Enabled = false;
                //button1.BackColor = System.Drawing.Color.White;
                //MessageBox.Show("Data's Imported Sucessfully");        
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Error Shown At End" + Ex);  
            }
            
          
        }

        private void Customer()
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public void Form1_Load(object sender, EventArgs e)
        {
            //DbConnection ConDb = new DbConnection();
           // ConnectCompany();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        public void ExportCustomer()//Customer
        {
            try
            {
                string sPath = "";
                string FileName = "BPDetails.xml";
                General g = new General();
                if (!File.Exists(sPath + FileName))
                { File.Create(sPath + FileName); }
                ConnectCompany();
                SAPbobsCOM.Recordset oRsInv = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                int recCount = 0;
                System.Xml.XmlDocument oXmlDoc = null;
                string sXmlString = null;
                StrSql = "SELECT BPCode,BPName,isnull(Category,''),isnull(Address1,'')[Address1],isnull(Address2,'')[Address2],isnull(Address3,'.')[Address3],isnull(Pin,'')[Pin],isnull(Telephone,'')[Telephone],isnull(Mobile,'')[Mobile],isnull(E_mail,'')[E_mail],isnull(TinNo,'')[TinNo],isnull(CSTNo,'')[CSTNo],isnull(CreditLimit,0)[CreditLimit],isnull(SalEmpNo,'-1')[SalEmpNo] FROM [NOR_BP_MASTER] where  IntegrationStatus='New'";
                DataSet objDataSet = ConDb.DbDataFromBranch(StrSql);
                sXmlString = objDataSet.GetXml();
                oXmlDoc = new System.Xml.XmlDocument();
                oXmlDoc.LoadXml(sXmlString);
                oXmlDoc.Save((sPath + FileName));
                SAPbobsCOM.BusinessPartners oBPMaster;

                XmlDocument reader = new XmlDocument();
                XmlDocument readerlines = new XmlDocument();
                IFormatProvider ifp = new System.Globalization.CultureInfo("en-US", true);
                reader.Load(sPath + FileName);

                XmlNodeList list = reader.GetElementsByTagName(g.row1);

                foreach (XmlNode node in list)
                {
                    XmlElement Element = (XmlElement)node;
                    oBPMaster = (SAPbobsCOM.BusinessPartners)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                    oBPMaster.CardCode = Element.GetElementsByTagName("BPCode")[0].InnerText;
                    oBPMaster.CardName = Element.GetElementsByTagName("BPName")[0].InnerText;
                    oBPMaster.Address = Element.GetElementsByTagName("Address1")[0].InnerText;

                    oBPMaster.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;
                    string strAddress = Element.GetElementsByTagName("Address1")[0].InnerText;
                    if (strAddress == "")
                        oBPMaster.Addresses.AddressName = Element.GetElementsByTagName("BPName")[0].InnerText;
                    else
                        oBPMaster.Addresses.AddressName = Element.GetElementsByTagName("Address1")[0].InnerText;
                    oBPMaster.Addresses.Block = Element.GetElementsByTagName("Address2")[0].InnerText;
                    oBPMaster.Addresses.City = Element.GetElementsByTagName("Address3")[0].InnerText;
                    oBPMaster.Addresses.ZipCode = Element.GetElementsByTagName("Pin")[0].InnerText;
                    oBPMaster.FiscalTaxID.Address = Element.GetElementsByTagName("Address1")[0].InnerText;
                    oBPMaster.FiscalTaxID.TaxId0 = Element.GetElementsByTagName("CSTNo")[0].InnerText;

                    oBPMaster.ShipToBuildingFloorRoom = Element.GetElementsByTagName("Address1")[0].InnerText;
                    oBPMaster.Block = Element.GetElementsByTagName("Address2")[0].InnerText;
                    oBPMaster.City = Element.GetElementsByTagName("Address3")[0].InnerText;
                    oBPMaster.ZipCode = Element.GetElementsByTagName("Pin")[0].InnerText;
                    oBPMaster.Phone1 = Element.GetElementsByTagName("Telephone")[0].InnerText;
                    oBPMaster.Cellular = Element.GetElementsByTagName("Mobile")[0].InnerText;
                    oBPMaster.EmailAddress = Element.GetElementsByTagName("E_mail")[0].InnerText;

                    oBPMaster.CreditLimit = Convert.ToDouble(Element.GetElementsByTagName("CreditLimit")[0].InnerText);
                    oBPMaster.SalesPersonCode = Convert.ToInt32(Element.GetElementsByTagName("SalEmpNo")[0].InnerText);

                    oBPMaster.UserFields.Fields.Item("U_IntegratedStatus").Value = "I";
                    oBPMaster.CardType = SAPbobsCOM.BoCardTypes.cCustomer;
                    int iErrorCode = oBPMaster.Add();
                    if (iErrorCode != 0)
                    {
                        string sErrorMsg = General.oCompany.GetLastErrorDescription();
                        MessageBox.Show("Error in Export BP Master : " + sErrorMsg);
                        if (iErrorCode == -2035)
                        {
                            StrSql = "Update [NOR_BP_MASTER] set IntegrationStatus='I' where BPcode='" + Element.GetElementsByTagName("BPCode")[0].InnerText + "'";
                            ConDb.QueryNonExecuteBranch(StrSql);
                        }
                    }
                    else
                    {
                        StrSql = "Update [NOR_BP_MASTER] set IntegrationStatus='I' where BPcode='" + Element.GetElementsByTagName("BPCode")[0].InnerText + "'";
                        ConDb.QueryNonExecuteBranch(StrSql);

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while Exportring BP Data.", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        public void ITEMGROUP()
        {
            try
            {
                int iError = 0;
                string sPath = "";
                string FileName = "Item.xml";
                string StrSql = "";
                string insertuser = "";
                General g = new General();
                if (!File.Exists(sPath + FileName))
                { File.Create(sPath + FileName); }
                int recCount = 0;
                System.Xml.XmlDocument oXmlDoc = null;
                string sXmlString = null;
                SAPbobsCOM.Recordset rsCompany = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                StrSql = "select ItmsGrpCod,ItmsGrpNam  from OITB where U_isIntegrated = 'N'";
                DataSet objDataSet = ConDb.DbDataFromSAP(StrSql);
                rsCompany.DoQuery(StrSql);
                if (rsCompany.RecordCount > 0)
                {
                    string strsqlUnitGet = "select Code from [@NOR_UNITMASTER] ";
                    rsCompany.DoQuery(strsqlUnitGet);
                    int i = 1;
                    while (!rsCompany.EoF)
                    {
                        string UnitGet = rsCompany.Fields.Item("Code").Value.ToString();

                        string QRY11 = "Select * from [@NOR_BRANCH_DTL] Where U_UnitId ='" + UnitGet + "'";
                        SAPbobsCOM.Recordset rsCompany1 = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                        rsCompany1.DoQuery(QRY11);
                        if (rsCompany1.RecordCount > 0)
                        {
                            string Licserver = rsCompany1.Fields.Item("U_LicServer").Value.ToString();
                            string server = rsCompany1.Fields.Item("U_ServerName").Value.ToString();
                            string DB = rsCompany1.Fields.Item("U_CompanyDB").Value.ToString();
                            string sUser = rsCompany1.Fields.Item("U_SAPUserName").Value.ToString();
                            string sPass = rsCompany1.Fields.Item("U_SAPPassword").Value.ToString();
                            string sqUser = rsCompany1.Fields.Item("U_ServerUser").Value.ToString();
                            string sqPass = rsCompany1.Fields.Item("U_ServerPass").Value.ToString();
                            g.connectOtherCompany(server,Licserver, DB, sUser, sPass, sqUser, sqPass);

                        }
                        sXmlString = objDataSet.GetXml();
                        oXmlDoc = new System.Xml.XmlDocument();
                        oXmlDoc.LoadXml(sXmlString);
                        oXmlDoc.Save((sPath + FileName));


                        XmlDocument reader = new XmlDocument();
                        XmlDocument readerlines = new XmlDocument();
                        IFormatProvider ifp = new System.Globalization.CultureInfo("en-US", true);
                        reader.Load(sPath + FileName);

                        XmlNodeList list = reader.GetElementsByTagName(g.row1);

                        foreach (XmlNode node in list)
                        {
                            XmlElement Element = (XmlElement)node;
                            try
                            {

                                #region ItemGroup (ADD)

                                SAPbobsCOM.ItemGroups oItemGroups = (SAPbobsCOM.ItemGroups)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItemGroups);
                                {
                                    
                                    oItemGroups.GroupName = Element.GetElementsByTagName("ItmsGrpNam")[0].InnerText;
                                    iError = 0;
                                    iError = oItemGroups.Add();
                                    if (iError != 0)
                                    {
                                        //string sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                        //MessageBox.Show(sErrorMsg + "Item master");
                                    }

                                }
                                #endregion
                            }
                            catch
                            {

                            }
                            if (rsCompany.RecordCount == i)
                            {
                                _str_update = @"UPDATE OITB set U_isIntegrated='Y' where OITB.ItmsGrpNam='" + Element.GetElementsByTagName("ItmsGrpNam")[0].InnerText + "'";
                                insertuser = ConDb.ScalarExecuteSAP(_str_update);
                            }



                        }
                        i++;
                        rsCompany.MoveNext();

                        if (iError == 0)
                        {
                            //_str_update = @"UPDATE OITB set U_isIntegrated='Y' where OITB.ItmsGrpNam='" + Element.GetElementsByTagName("ItmsGrpNam")[0].InnerText + "'";
                            //insertuser = ConDb.ScalarExecuteSAP(_str_update);

                        }

                    }

                }
            }
            catch {  }

        }

         #region ITEM MASTER

                public void ITEMMASTER()
                {
                    try
                    {
                        string sPath = "";
                        string FileName = "Item.xml";
                        string StrSql = "";
                        string insertuser = "";
                        General g = new General();
                        if (!File.Exists(sPath + FileName))
                        { File.Create(sPath + FileName); }
                        int recCount = 0;
                        System.Xml.XmlDocument oXmlDoc = null;
                        string sXmlString = null;
                        StrSql = @"SELECT OITM.ItemCode ItemCode,isnull(ItemName,'-')ItemName,oitm.ItmsGrpCod,B.U_UnitCode,
        isnull(CodeBars,0)[CodeBars],isnull(oitw.OnHand,0) OnHand,isnull(oitw.WhsCode ,0)WhsCode ,isnull(InvntryUom,0)[InvntryUom],isnull(OITW.IsCommited,0)[IsCommited],isnull(OITM.CardCode,'-')[CardCode],OITM.FirmCode,
        OITM.SellItem,OITM.InvntItem,OITM.PrchseItem,OITM.AssetItem,isnull(OITM.FrgnName,'-')[FrgnName],OITM.ItemType,
        isnull(OITM.NumInSale,1)[NumInSale],isnull(OITM.SalPackUn,1)SalPackUn,isnull(OITM.U_Brand,'')U_Brand,isnull(OITM.U_Category,'')U_Category,
        isnull(OITM.U_Class,'')U_Class,isnull(OITM.U_Color,'')U_Color,isnull(OITM.U_Export,'N')U_Export,isnull(OITM.U_GrpCode,'')U_GrpCode,isnull(OITM.U_Model,'')U_Model,
        isnull(OITM.U_NofPairs,'')U_NofPairs,isnull(OITM.U_OrdQty,0)U_OrdQty,isnull(OITM.U_PairSize,'')U_PairSize,isnull(OITM.U_Priority,0)U_Priority,isnull(OITM.U_SFGCat,'')U_SFGCat,
        isnull(OITM.U_Size,'')U_Size,isnull(OITM.U_SizeCat,'')U_SizeCat,isnull(OITM.U_StdSize,'')U_StdSize,isnull(OITM.U_Unit,'')U_Unit,OITM.GLMethod,ISNULL( OITM.U_HsnCcode,'')U_HsnCcode,isnull(oitm.U_VATCode,'')U_VATCode,ISNULL(OITM.U_VatRate,'')U_VatRate,ISNULL(OITM.U_JbType,'')U_JbType,isnull(oitm.U_Abtment,'')U_Abtment
        from OITM
        left join OITW on OITW.ItemCode=OITM.ItemCode
        inner join OITB on OITM.ItmsGrpCod = OITB.ItmsGrpCod
        inner join OWHS on OWHS.WhsCode=OITW.WhsCode
        INNER JOIN [@NOR_OITM_UNIT] B on OITM.ItemCode=B.U_ItemCode WHERE B.U_IsIntegrated='N' AND B.U_UnitCode IS NOT NULL and OWHS.U_Unit=OITM.U_IUnits order by OITM.ItemCode";

                        DataSet objDataSet = ConDb.DbDataFromSAP(StrSql);
                        if (objDataSet.Tables[0].Rows.Count > 0)
                        {
                            string strsqlUnitGet = "select Code from [@NOR_UNITMASTER] ";
                            SAPbobsCOM.Recordset rsCompany = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);            
                            rsCompany.DoQuery(strsqlUnitGet);
                            while (!rsCompany.EoF)
                            {
                                string UnitGet = rsCompany.Fields.Item("Code").Value.ToString();

                                ////string QRY11 = "Select * from [@NOR_BRANCH_DTL] Where U_UnitId ='" + objDataSet.Tables[0].Rows[0]["U_UnitCode"].ToString() + "'";
                                ////SAPbobsCOM.Recordset rsCompany1 = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                //////SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                ////rsCompany1.DoQuery(QRY11);
                                ////if (rsCompany1.RecordCount > 0)
                                ////{
                                ////    string server = rsCompany1.Fields.Item("U_ServerName").Value.ToString();
                                ////    string DB = rsCompany1.Fields.Item("U_CompanyDB").Value.ToString();
                                ////    string sUser = rsCompany1.Fields.Item("U_SAPUserName").Value.ToString();
                                ////    string sPass = rsCompany1.Fields.Item("U_SAPPassword").Value.ToString();
                                ////    string sqUser = rsCompany1.Fields.Item("U_ServerUser").Value.ToString();
                                ////    string sqPass = rsCompany1.Fields.Item("U_ServerPass").Value.ToString();
                                ////    g.connectOtherCompany(server, DB, sUser, sPass, sqUser, sqPass);
                                ////}

                                sXmlString = objDataSet.GetXml();
                                // sXmlString = oRsInv.GetAsXML();
                                oXmlDoc = new System.Xml.XmlDocument();
                                oXmlDoc.LoadXml(sXmlString);
                                oXmlDoc.Save((sPath + FileName));



                                XmlDocument reader = new XmlDocument();
                                XmlDocument readerlines = new XmlDocument();
                                IFormatProvider ifp = new System.Globalization.CultureInfo("en-US", true);
                                reader.Load(sPath + FileName);

                                XmlNodeList list = reader.GetElementsByTagName(g.row1);

                                foreach (XmlNode node in list)
                                {

                                    XmlElement Element = (XmlElement)node;
                                    #region Item Master (ADD)
                                    string strUnitCd = Element.GetElementsByTagName("U_UnitCode")[0].InnerText;

                                    if (strUnitCd == UnitGet)
                                    {
                                        string QRY11 = "Select * from [@NOR_BRANCH_DTL] Where U_UnitId ='" + UnitGet + "'";
                                        SAPbobsCOM.Recordset rsCompany1 = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        //SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                        rsCompany1.DoQuery(QRY11);
                                        bool _connection = false;
                                        if (rsCompany1.RecordCount > 0)
                                        {
                                            string Licserver = rsCompany1.Fields.Item("U_Licserver").Value.ToString();
                                            string server = rsCompany1.Fields.Item("U_ServerName").Value.ToString();
                                            string DB = rsCompany1.Fields.Item("U_CompanyDB").Value.ToString();
                                            string sUser = rsCompany1.Fields.Item("U_SAPUserName").Value.ToString();
                                            string sPass = rsCompany1.Fields.Item("U_SAPPassword").Value.ToString();
                                            string sqUser = rsCompany1.Fields.Item("U_ServerUser").Value.ToString();
                                            string sqPass = rsCompany1.Fields.Item("U_ServerPass").Value.ToString();
                                            _connection = g.connectOtherCompany(server,Licserver, DB, sUser, sPass, sqUser, sqPass);
                                        }
                                        //string QRY1 = "Select * from [@NOR_BRANCH_DTL] Where U_UnitId ='" + strUnitCd + "'";
                                        //SAPbobsCOM.Recordset rsCompany = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        ////SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                        //rsCompany.DoQuery(QRY1);
                                        //if (rsCompany.RecordCount > 0)
                                        //{
                                        //    string server = rsCompany.Fields.Item("U_ServerName").Value.ToString();
                                        //    string DB = rsCompany.Fields.Item("U_CompanyDB").Value.ToString();
                                        //    string sUser = rsCompany.Fields.Item("U_SAPUserName").Value.ToString();
                                        //    string sPass = rsCompany.Fields.Item("U_SAPPassword").Value.ToString();
                                        //    string sqUser = rsCompany.Fields.Item("U_ServerUser").Value.ToString();
                                        //    string sqPass = rsCompany.Fields.Item("U_ServerPass").Value.ToString();
                                        //    g.connectOtherCompany(server, DB, sUser, sPass, sqUser, sqPass);

                                        //oSales = (SAPbobsCOM.Documents)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                                        //oBPMaster = (SAPbobsCOM.BusinessPartners)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                                        if (_connection == true)
                                        {
                                            SAPbobsCOM.Items oItem = (SAPbobsCOM.Items)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                                            
                                            string QryyExistChk = "SELECT * FROM OITM WHERE ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'";
                                            SAPbobsCOM.Recordset rsItem = ((SAPbobsCOM.Recordset)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                           rsItem.DoQuery(QryyExistChk);
                                            if (rsItem.RecordCount > 0)
                                            {
                                                oItem.GetByKey(Element.GetElementsByTagName("ItemCode")[0].InnerText);
                                                //oItem.ItemCode = Element.GetElementsByTagName("ItemCode")[0].InnerText;
                                                oItem.ItemName = Element.GetElementsByTagName("ItemName")[0].InnerText;
                                                oItem.BarCode = Element.GetElementsByTagName("CodeBars")[0].InnerText;
                                                oItem.ForeignName = Element.GetElementsByTagName("FrgnName")[0].InnerText;
                                                //oItem.Mainsupplier = Element.GetElementsByTagName("CardCode")[0].InnerText;
                                                // oItem.ItemType = Element.GetElementsByTagName("ItemType")[0].InnerText;
                                               // oItem.GLMethod = Element.GetElementsByTagName("GLMethod")[0].InnerText;

                                                string _strWarehouse = Element.GetElementsByTagName("WhsCode")[0].InnerText;
                                               
                                                oItem.ItemsGroupCode = Convert.ToInt32(Element.GetElementsByTagName("ItmsGrpCod")[0].InnerText);
                                                oItem.GLMethod = SAPbobsCOM.BoGLMethods.glm_ItemClass;

                                                oItem.UserFields.Fields.Item("U_Class").Value = Element.GetElementsByTagName("U_Class")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Brand").Value = Element.GetElementsByTagName("U_Brand")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Category").Value = Element.GetElementsByTagName("U_Category")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Color").Value = Element.GetElementsByTagName("U_Color")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Export").Value = Element.GetElementsByTagName("U_Export")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_GrpCode").Value = Element.GetElementsByTagName("U_GrpCode")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Model").Value = Element.GetElementsByTagName("U_Model")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_NofPairs").Value = Element.GetElementsByTagName("U_NofPairs")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_OrdQty").Value = Element.GetElementsByTagName("U_OrdQty")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_PairSize").Value = Element.GetElementsByTagName("U_PairSize")[0].InnerText;
                                             //   oItem.UserFields.Fields.Item("U_Priority").Value = Element.GetElementsByTagName("U_Priority")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_SFGCat").Value = Element.GetElementsByTagName("U_SFGCat")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Size").Value = Element.GetElementsByTagName("U_Size")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_SizeCat").Value = Element.GetElementsByTagName("U_SizeCat")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_StdSize").Value = Element.GetElementsByTagName("U_StdSize")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Unit").Value = Element.GetElementsByTagName("U_Unit")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_HsnCcode").Value = Element.GetElementsByTagName("U_HsnCcode")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_VATCode").Value = Element.GetElementsByTagName("U_VATCode")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_VatRate").Value = Element.GetElementsByTagName("U_VatRate")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_JbType").Value = Element.GetElementsByTagName("U_JbType")[0].InnerText;

                                                //if (_strWarehouse != "0")
                                                //{
                                                //    string strWrhse = "select OITW.WhsCode from OITW where ItemCode = '" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'";
                                                //    DataSet DatasetWrhse = ConDb.DbDataFromSAP(strWrhse);

                                             
                                                //    for (int i = 0; i < objDataSet.Tables[0].Rows.Count; i++)
                                                //    {
                                                //        string strWrhseUnit = "select OITW.WhsCode from OITW where ItemCode = '" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and OITW.WhsCode = '" + DatasetWrhse.Tables[0].Rows[i][0].ToString() + "'";
                                                //    SAPbobsCOM.Recordset rsWhs = ((SAPbobsCOM.Recordset)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                                //     rsWhs.DoQuery(strWrhseUnit);
                                                //         string strWrhseCount = "select count(OITW.WhsCode)CntWhs from OITW where ItemCode = '" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'";
                                                //        SAPbobsCOM.Recordset rsWhsCount = ((SAPbobsCOM.Recordset)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                                //     rsWhsCount.DoQuery(strWrhseCount);
                                                //     if (rsWhs.Fields.Item(0).Value.ToString() == "")
                                                //     {
                                                        
                                                //        // oItem.WhsInfo.SetCurrentLine(Convert.ToInt32( rsWhsCount.Fields.Item("CntWhs").Value));
                                                //         oItem.WhsInfo.WarehouseCode = objDataSet.Tables[0].Rows[i]["WhsCode"].ToString();
                                                //         oItem.WhsInfo.Add();
                                                //     }

                                                //    }
                                                //}
                                                
                                                //oItem.Manufacturer = Convert.ToInt32(Element.GetElementsByTagName("FirmCode")[0].InnerText);
                                          
                                                oItem.InventoryUOM = Element.GetElementsByTagName("InvntryUom")[0].InnerText;


                                                //if (Element.GetElementsByTagName("InvntItem")[0].InnerText.ToString() == "Y")
                                                //{
                                                //    oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                                //}
                                                //else
                                                //{
                                                //    oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                                //}
                                                //if (Element.GetElementsByTagName("SellItem")[0].InnerText.ToString() == "Y")
                                                //{
                                                //    oItem.SalesItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                                //}
                                                //else
                                                //{
                                                //    oItem.SalesItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                                //}
                                                //if (Element.GetElementsByTagName("PrchseItem")[0].InnerText.ToString() == "Y")
                                                //{
                                                //    oItem.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                                //}
                                                //else
                                                //{
                                                //    oItem.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                                //}
                                                //if (Element.GetElementsByTagName("AssetItem")[0].InnerText.ToString() == "Y")
                                                //{
                                                //    oItem.AssetItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                                //}
                                                //else
                                                //{
                                                //    oItem.AssetItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                                //}

                                                //oItem.CostAccountingMethod = SAPbobsCOM.BoInventorySystem.bis_FIFO;
                                                int iError = 0;
                                                iError = oItem.Update();
                                                if (iError != 0)
                                                {
                                                    string sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                                    //Global.SapApplication.StatusBar.SetText(sErrorMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                                    //return false;
                                                }
                                                else
                                                {

                                                     #region ItemPriceList
                                                       SAPbobsCOM.Recordset rspricelist = (SAPbobsCOM.Recordset)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);            
                                                        string strItemPrclistQry = @"select PriceList,Price from ITM1 
                                                                                 inner join OITM on OITM.ItemCode=ITM1.ItemCode
                                                                                 where OITM.ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and Price is not NULL";

                                                      

                                                        DataSet datasetPrclist = ConDb.DbDataFromSAP(strItemPrclistQry);
                                                        for (int i = 0; i < datasetPrclist.Tables[0].Rows.Count; i++)
                                                        {
                                                            string strUpdate = "update ITM1 set Price='" + datasetPrclist.Tables[0].Rows[i][1].ToString() + "' where ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and PriceList='" + datasetPrclist.Tables[0].Rows[i][0].ToString() + "'";
                                                            rspricelist.DoQuery(strUpdate);

                                                        }


                                                        #endregion


                                                    #region Item Property(UPDATE)
                                                    SAPbobsCOM.Items oItemUpdate = (SAPbobsCOM.Items)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                                                    string QryPropery = @"SELECT OITM.QryGroup1,OITM.QryGroup2,OITM.QryGroup3,
        OITM.QryGroup4,OITM.QryGroup5,OITM.QryGroup6,OITM.QryGroup7,OITM.QryGroup8,OITM.QryGroup9,OITM.QryGroup10,OITM.QryGroup11,OITM.QryGroup12,OITM.QryGroup13,
        OITM.QryGroup14,OITM.QryGroup15,OITM.QryGroup16,OITM.QryGroup17,OITM.QryGroup18,OITM.QryGroup19,OITM.QryGroup20,OITM.QryGroup21,OITM.QryGroup22,OITM.QryGroup23,
        OITM.QryGroup24,OITM.QryGroup25,OITM.QryGroup26,OITM.QryGroup27,OITM.QryGroup28,OITM.QryGroup29,OITM.QryGroup30,OITM.QryGroup31,OITM.QryGroup32,OITM.QryGroup33,
        OITM.QryGroup34,OITM.QryGroup35,OITM.QryGroup36,OITM.QryGroup37,OITM.QryGroup38,OITM.QryGroup39,OITM.QryGroup40,OITM.QryGroup41,OITM.QryGroup42,OITM.QryGroup43,OITM.QryGroup44,
        OITM.QryGroup45,OITM.QryGroup46,OITM.QryGroup47,OITM.QryGroup48,OITM.QryGroup49,OITM.QryGroup50,OITM.QryGroup51,OITM.QryGroup52,OITM.QryGroup53,OITM.QryGroup54,OITM.QryGroup55,
        OITM.QryGroup56,OITM.QryGroup57,OITM.QryGroup58,OITM.QryGroup59,OITM.QryGroup60,OITM.QryGroup61,OITM.QryGroup62,OITM.QryGroup63,OITM.QryGroup64 FROM OITM WHERE ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'";
                                                    SAPbobsCOM.Recordset rspROPERTY = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                    SAPbobsCOM.Recordset rspROPERTYUpdate = (SAPbobsCOM.Recordset)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                    //SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                                    rspROPERTY.DoQuery(QryPropery);
                                                    string QryPrtyUpdate = @"UPDATE OITM  SET QryGroup1='" + rspROPERTY.Fields.Item("QryGroup1").Value.ToString() + "',QryGroup2='" + rspROPERTY.Fields.Item("QryGroup2").Value.ToString() + "',QryGroup3='" + rspROPERTY.Fields.Item("QryGroup3").Value.ToString() + "',QryGroup4='" + rspROPERTY.Fields.Item("QryGroup4").Value.ToString() + "',QryGroup5='" + rspROPERTY.Fields.Item("QryGroup5").Value.ToString() + "',QryGroup6='" + rspROPERTY.Fields.Item("QryGroup6").Value.ToString() + "',QryGroup7='" + rspROPERTY.Fields.Item("QryGroup7").Value.ToString() + "',QryGroup8='" + rspROPERTY.Fields.Item("QryGroup8").Value.ToString() + "',QryGroup9='" + rspROPERTY.Fields.Item("QryGroup9").Value.ToString() + "',QryGroup10='" + rspROPERTY.Fields.Item("QryGroup10").Value.ToString() + "',QryGroup11='" + rspROPERTY.Fields.Item("QryGroup11").Value.ToString() + "',QryGroup12='" + rspROPERTY.Fields.Item("QryGroup12").Value.ToString() + "',QryGroup13='" + rspROPERTY.Fields.Item("QryGroup13").Value.ToString() + "',QryGroup14='" + rspROPERTY.Fields.Item("QryGroup14").Value.ToString() + "',QryGroup15='" + rspROPERTY.Fields.Item("QryGroup15").Value.ToString() + "',QryGroup16='" + rspROPERTY.Fields.Item("QryGroup16").Value.ToString() + "',QryGroup17='" + rspROPERTY.Fields.Item("QryGroup17").Value.ToString() + "',QryGroup18='" + rspROPERTY.Fields.Item("QryGroup18").Value.ToString() + "',QryGroup19='" + rspROPERTY.Fields.Item("QryGroup19").Value.ToString() + "',QryGroup20='" + rspROPERTY.Fields.Item("QryGroup20").Value.ToString() + "',QryGroup21='" + rspROPERTY.Fields.Item("QryGroup21").Value.ToString() + "',QryGroup22='" + rspROPERTY.Fields.Item("QryGroup22").Value.ToString() + "',QryGroup23='" + rspROPERTY.Fields.Item("QryGroup23").Value.ToString() + "',QryGroup24='" + rspROPERTY.Fields.Item("QryGroup24").Value.ToString() + "',QryGroup25='" + rspROPERTY.Fields.Item("QryGroup25").Value.ToString() + "',QryGroup26='" + rspROPERTY.Fields.Item("QryGroup26").Value.ToString() + "',QryGroup27='" + rspROPERTY.Fields.Item("QryGroup27").Value.ToString() + "',QryGroup28='" + rspROPERTY.Fields.Item("QryGroup28").Value.ToString() + "',QryGroup29='" + rspROPERTY.Fields.Item("QryGroup29").Value.ToString() + "',QryGroup30='" + rspROPERTY.Fields.Item("QryGroup30").Value.ToString() + "',QryGroup31='" + rspROPERTY.Fields.Item("QryGroup31").Value.ToString() + "',QryGroup32='" + rspROPERTY.Fields.Item("QryGroup32").Value.ToString() + "',QryGroup33='" + rspROPERTY.Fields.Item("QryGroup33").Value.ToString() + "',QryGroup34='" + rspROPERTY.Fields.Item("QryGroup34").Value.ToString() + "',QryGroup35='" + rspROPERTY.Fields.Item("QryGroup35").Value.ToString() + "',QryGroup36='" + rspROPERTY.Fields.Item("QryGroup36").Value.ToString() + "',QryGroup37='" + rspROPERTY.Fields.Item("QryGroup37").Value.ToString() + "',QryGroup38='" + rspROPERTY.Fields.Item("QryGroup38").Value.ToString() + "',QryGroup39='" + rspROPERTY.Fields.Item("QryGroup39").Value.ToString() + "',QryGroup40='" + rspROPERTY.Fields.Item("QryGroup40").Value.ToString() + "',QryGroup41='" + rspROPERTY.Fields.Item("QryGroup41").Value.ToString() + "',QryGroup42='" + rspROPERTY.Fields.Item("QryGroup42").Value.ToString() + "',QryGroup43='" + rspROPERTY.Fields.Item("QryGroup43").Value.ToString() + "',QryGroup44='" + rspROPERTY.Fields.Item("QryGroup44").Value.ToString() + "',QryGroup45='" + rspROPERTY.Fields.Item("QryGroup45").Value.ToString() + "',QryGroup46='" + rspROPERTY.Fields.Item("QryGroup46").Value.ToString() + "',QryGroup47='" + rspROPERTY.Fields.Item("QryGroup47").Value.ToString() + "',QryGroup48='" + rspROPERTY.Fields.Item("QryGroup48").Value.ToString() + "',QryGroup49='" + rspROPERTY.Fields.Item("QryGroup49").Value.ToString() + "',QryGroup50='" + rspROPERTY.Fields.Item("QryGroup50").Value.ToString() + "',QryGroup51='" + rspROPERTY.Fields.Item("QryGroup51").Value.ToString() + "',QryGroup52='" + rspROPERTY.Fields.Item("QryGroup52").Value.ToString() + "',QryGroup53='" + rspROPERTY.Fields.Item("QryGroup53").Value.ToString() + "',QryGroup54='" + rspROPERTY.Fields.Item("QryGroup54").Value.ToString() + "',QryGroup55='" + rspROPERTY.Fields.Item("QryGroup55").Value.ToString() + "',QryGroup56='" + rspROPERTY.Fields.Item("QryGroup56").Value.ToString() + "',QryGroup57='" + rspROPERTY.Fields.Item("QryGroup57").Value.ToString() + "',QryGroup58='" + rspROPERTY.Fields.Item("QryGroup58").Value.ToString() + "',QryGroup59='" + rspROPERTY.Fields.Item("QryGroup59").Value.ToString() + "',QryGroup60='" + rspROPERTY.Fields.Item("QryGroup60").Value.ToString() + "',QryGroup61='" + rspROPERTY.Fields.Item("QryGroup61").Value.ToString() + "',QryGroup62='" + rspROPERTY.Fields.Item("QryGroup62").Value.ToString() + "',QryGroup63='" + rspROPERTY.Fields.Item("QryGroup63").Value.ToString() + "',QryGroup64='" + rspROPERTY.Fields.Item("QryGroup64").Value.ToString() + "' WHERE ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'";

                                                    rspROPERTYUpdate.DoQuery(QryPrtyUpdate);
                                                    //oItemUpdate.GetByKey(Element.GetElementsByTagName("ItemCode")[0].InnerText);
                                                    //for (int j = 1; j < strPrty.Length; j++)
                                                    //{
                                                    //    if (strPrty[j] != "" & strPrty[j] != null)
                                                    //    {
                                                    //        oItemUpdate.set_Properties(Convert.ToInt32(strPrty[j]), SAPbobsCOM.BoYesNoEnum.tYES);
                                                    //    }
                                                    //}
                                                    //oItemUpdate.Update();

                                                    //string _str_Query = "UPDATE [ITM1]  SET [Price] = '" + frmItem.DataSources.UserDataSources.Item("udsMRP").ValueEx + "' WHERE ItemCode = '" + strItemCod + "'AND PriceList = 2 ";
                                                    //oRecordSet.DoQuery(_str_Query);

                                                    #endregion

                                                    SAPbobsCOM.Recordset rspUpdateStatus = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                    string strUpdateStatus = "Update [@NOR_OITM_UNIT] SET U_IsIntegrated='Y' WHERE U_ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and U_UnitCode = '" + strUnitCd + "'";
                                                    rspUpdateStatus.DoQuery(strUpdateStatus);

                                                    // Global.SapApplication.StatusBar.SetText("Operation Completed Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                                }  // return true;

                                            }
                                            else
                                            {
                                                SAPbobsCOM.Recordset rsCompanyPricelist = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                string _strClass = Element.GetElementsByTagName("U_Class")[0].InnerText;
                                                string _strWarehouse = Element.GetElementsByTagName("WhsCode")[0].InnerText;
                                                if ((_strClass == "2" || _strClass == "3") && _strWarehouse == "0")
                                                {
                                                    MessageBox.Show("Please select warehouse for Item '" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'");
                                                }
                                                else
                                                {
                                                    oItem.ItemCode = Element.GetElementsByTagName("ItemCode")[0].InnerText;
                                                    
                                                    oItem.ItemName = Element.GetElementsByTagName("ItemName")[0].InnerText;
                                                    oItem.BarCode = Element.GetElementsByTagName("CodeBars")[0].InnerText;
                                                    oItem.ForeignName = Element.GetElementsByTagName("FrgnName")[0].InnerText;
                                                    // oItem.Mainsupplier = Element.GetElementsByTagName("CardCode")[0].InnerText;
                                                    // oItem.ItemType=Element.GetElementsByTagName("ItemType")[0].InnerText;


                                                    oItem.ItemsGroupCode = Convert.ToInt32(Element.GetElementsByTagName("ItmsGrpCod")[0].InnerText);
                                                   // oItem.GLMethod = Element.GetElementsByTagName("GLMethod")[0].InnerText.ToString();
                                                    oItem.GLMethod = SAPbobsCOM.BoGLMethods.glm_ItemClass;
                                                    oItem.Manufacturer = Convert.ToInt32(Element.GetElementsByTagName("FirmCode")[0].InnerText);

                                                    oItem.InventoryUOM = Element.GetElementsByTagName("InvntryUom")[0].InnerText;
                                                    string strWrhse = "select OITW.WhsCode from OITW inner join OITM on OITW.ItemCode =OITM.ItemCode inner join OWHS on OITW.WhsCode=OWHS.WhsCode where oitm.ItemCode = '" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and owhs.U_Unit='" + UnitGet + "'";
                                                    DataSet DatasetWrhse=ConDb.DbDataFromSAP(strWrhse);
                                                    
                                             
                                                    if (_strWarehouse != "0")
                                                    {

                                                        for (int i = 0; i < DatasetWrhse.Tables[0].Rows.Count; i++)
                                                        {
                                                            oItem.WhsInfo.WarehouseCode = DatasetWrhse.Tables[0].Rows[i]["WhsCode"].ToString();
                                                            oItem.WhsInfo.Add();
                                                        }
                                                    }
                                                    
                                                    //-------------User Fields-------------------------------------------------------------------//
                                                    oItem.UserFields.Fields.Item("U_Class").Value = Element.GetElementsByTagName("U_Class")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Brand").Value = Element.GetElementsByTagName("U_Brand")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Category").Value = Element.GetElementsByTagName("U_Category")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Color").Value = Element.GetElementsByTagName("U_Color")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Export").Value = Element.GetElementsByTagName("U_Export")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_GrpCode").Value = Element.GetElementsByTagName("U_GrpCode")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Model").Value = Element.GetElementsByTagName("U_Model")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_NofPairs").Value = Element.GetElementsByTagName("U_NofPairs")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_OrdQty").Value = Element.GetElementsByTagName("U_OrdQty")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_PairSize").Value = Element.GetElementsByTagName("U_PairSize")[0].InnerText;
                                                 //   oItem.UserFields.Fields.Item("U_Priority").Value = Element.GetElementsByTagName("U_Priority")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_SFGCat").Value = Element.GetElementsByTagName("U_SFGCat")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Size").Value = Element.GetElementsByTagName("U_Size")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_SizeCat").Value = Element.GetElementsByTagName("U_SizeCat")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_StdSize").Value = Element.GetElementsByTagName("U_StdSize")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Unit").Value = Element.GetElementsByTagName("U_Unit")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_HsnCcode").Value = Element.GetElementsByTagName("U_HsnCcode")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_VATCode").Value = Element.GetElementsByTagName("U_VATCode")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_VatRate").Value = Element.GetElementsByTagName("U_VatRate")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_JbType").Value = Element.GetElementsByTagName("U_JbType")[0].InnerText;
                                                    //--------------------------------------------------------------------------------------------//

                                                    if (Element.GetElementsByTagName("InvntItem")[0].InnerText.ToString() == "Y")
                                                    {
                                                        oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                                    }
                                                    else
                                                    {
                                                        oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                                    }
                                                    if (Element.GetElementsByTagName("SellItem")[0].InnerText.ToString() == "Y")
                                                    {
                                                        oItem.SalesItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                                    }
                                                    else
                                                    {
                                                        oItem.SalesItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                                    }
                                                    if (Element.GetElementsByTagName("PrchseItem")[0].InnerText.ToString() == "Y")
                                                    {
                                                        oItem.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                                    }
                                                    else
                                                    {
                                                        oItem.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                                    }
                                                    if (Element.GetElementsByTagName("AssetItem")[0].InnerText.ToString() == "Y")
                                                    {
                                                        oItem.AssetItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                                    }
                                                    else
                                                    {
                                                        oItem.AssetItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                                    }

                                                    oItem.CostAccountingMethod = SAPbobsCOM.BoInventorySystem.bis_FIFO;

                                                    int iError = 0;
                                                    iError = oItem.Add();
                                                    if (iError != 0)
                                                    {
                                                        string sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                                        MessageBox.Show(sErrorMsg + "Item master");
                                                        // Global.SapApplication.StatusBar.SetText(sErrorMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                                        //return false;
                                                    }
                                    #endregion
                                                    else
                                                    { 
                                                      #region ItemPriceList
                                                       SAPbobsCOM.Recordset rspricelist = (SAPbobsCOM.Recordset)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);            
                                                        string strItemPrclistQry = @"select PriceList,Price from ITM1 
                                                                                 inner join OITM on OITM.ItemCode=ITM1.ItemCode
                                                                                 where OITM.ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and Price is not NULL";

                                                      

                                                        DataSet datasetPrclist = ConDb.DbDataFromSAP(strItemPrclistQry);
                                                        for (int i = 0; i < datasetPrclist.Tables[0].Rows.Count; i++)
                                                        {
                                                            string strUpdate = "update ITM1 set Price='" + datasetPrclist.Tables[0].Rows[i][1].ToString() + "' where ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and PriceList='" + datasetPrclist.Tables[0].Rows[i][0].ToString() + "'";
                                                            rspricelist.DoQuery(strUpdate);

                                                        }


                                                        #endregion


                                                        #region Item Property(UPDATE)
                                                        SAPbobsCOM.Items oItemUpdate = (SAPbobsCOM.Items)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                                                        string QryPropery = @"SELECT OITM.QryGroup1,OITM.QryGroup2,OITM.QryGroup3,
        OITM.QryGroup4,OITM.QryGroup5,OITM.QryGroup6,OITM.QryGroup7,OITM.QryGroup8,OITM.QryGroup9,OITM.QryGroup10,OITM.QryGroup11,OITM.QryGroup12,OITM.QryGroup13,
        OITM.QryGroup14,OITM.QryGroup15,OITM.QryGroup16,OITM.QryGroup17,OITM.QryGroup18,OITM.QryGroup19,OITM.QryGroup20,OITM.QryGroup21,OITM.QryGroup22,OITM.QryGroup23,
        OITM.QryGroup24,OITM.QryGroup25,OITM.QryGroup26,OITM.QryGroup27,OITM.QryGroup28,OITM.QryGroup29,OITM.QryGroup30,OITM.QryGroup31,OITM.QryGroup32,OITM.QryGroup33,
        OITM.QryGroup34,OITM.QryGroup35,OITM.QryGroup36,OITM.QryGroup37,OITM.QryGroup38,OITM.QryGroup39,OITM.QryGroup40,OITM.QryGroup41,OITM.QryGroup42,OITM.QryGroup43,OITM.QryGroup44,
        OITM.QryGroup45,OITM.QryGroup46,OITM.QryGroup47,OITM.QryGroup48,OITM.QryGroup49,OITM.QryGroup50,OITM.QryGroup51,OITM.QryGroup52,OITM.QryGroup53,OITM.QryGroup54,OITM.QryGroup55,
        OITM.QryGroup56,OITM.QryGroup57,OITM.QryGroup58,OITM.QryGroup59,OITM.QryGroup60,OITM.QryGroup61,OITM.QryGroup62,OITM.QryGroup63,OITM.QryGroup64 FROM OITM WHERE ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'";
                                                        SAPbobsCOM.Recordset rspROPERTY = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        SAPbobsCOM.Recordset rspROPERTYUpdate = (SAPbobsCOM.Recordset)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        //SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                                        rspROPERTY.DoQuery(QryPropery);
                                                        string QryPrtyUpdate = @"UPDATE OITM  SET QryGroup1='" + rspROPERTY.Fields.Item("QryGroup1").Value.ToString() + "',QryGroup2='" + rspROPERTY.Fields.Item("QryGroup2").Value.ToString() + "',QryGroup3='" + rspROPERTY.Fields.Item("QryGroup3").Value.ToString() + "',QryGroup4='" + rspROPERTY.Fields.Item("QryGroup4").Value.ToString() + "',QryGroup5='" + rspROPERTY.Fields.Item("QryGroup5").Value.ToString() + "',QryGroup6='" + rspROPERTY.Fields.Item("QryGroup6").Value.ToString() + "',QryGroup7='" + rspROPERTY.Fields.Item("QryGroup7").Value.ToString() + "',QryGroup8='" + rspROPERTY.Fields.Item("QryGroup8").Value.ToString() + "',QryGroup9='" + rspROPERTY.Fields.Item("QryGroup9").Value.ToString() + "',QryGroup10='" + rspROPERTY.Fields.Item("QryGroup10").Value.ToString() + "',QryGroup11='" + rspROPERTY.Fields.Item("QryGroup11").Value.ToString() + "',QryGroup12='" + rspROPERTY.Fields.Item("QryGroup12").Value.ToString() + "',QryGroup13='" + rspROPERTY.Fields.Item("QryGroup13").Value.ToString() + "',QryGroup14='" + rspROPERTY.Fields.Item("QryGroup14").Value.ToString() + "',QryGroup15='" + rspROPERTY.Fields.Item("QryGroup15").Value.ToString() + "',QryGroup16='" + rspROPERTY.Fields.Item("QryGroup16").Value.ToString() + "',QryGroup17='" + rspROPERTY.Fields.Item("QryGroup17").Value.ToString() + "',QryGroup18='" + rspROPERTY.Fields.Item("QryGroup18").Value.ToString() + "',QryGroup19='" + rspROPERTY.Fields.Item("QryGroup19").Value.ToString() + "',QryGroup20='" + rspROPERTY.Fields.Item("QryGroup20").Value.ToString() + "',QryGroup21='" + rspROPERTY.Fields.Item("QryGroup21").Value.ToString() + "',QryGroup22='" + rspROPERTY.Fields.Item("QryGroup22").Value.ToString() + "',QryGroup23='" + rspROPERTY.Fields.Item("QryGroup23").Value.ToString() + "',QryGroup24='" + rspROPERTY.Fields.Item("QryGroup24").Value.ToString() + "',QryGroup25='" + rspROPERTY.Fields.Item("QryGroup25").Value.ToString() + "',QryGroup26='" + rspROPERTY.Fields.Item("QryGroup26").Value.ToString() + "',QryGroup27='" + rspROPERTY.Fields.Item("QryGroup27").Value.ToString() + "',QryGroup28='" + rspROPERTY.Fields.Item("QryGroup28").Value.ToString() + "',QryGroup29='" + rspROPERTY.Fields.Item("QryGroup29").Value.ToString() + "',QryGroup30='" + rspROPERTY.Fields.Item("QryGroup30").Value.ToString() + "',QryGroup31='" + rspROPERTY.Fields.Item("QryGroup31").Value.ToString() + "',QryGroup32='" + rspROPERTY.Fields.Item("QryGroup32").Value.ToString() + "',QryGroup33='" + rspROPERTY.Fields.Item("QryGroup33").Value.ToString() + "',QryGroup34='" + rspROPERTY.Fields.Item("QryGroup34").Value.ToString() + "',QryGroup35='" + rspROPERTY.Fields.Item("QryGroup35").Value.ToString() + "',QryGroup36='" + rspROPERTY.Fields.Item("QryGroup36").Value.ToString() + "',QryGroup37='" + rspROPERTY.Fields.Item("QryGroup37").Value.ToString() + "',QryGroup38='" + rspROPERTY.Fields.Item("QryGroup38").Value.ToString() + "',QryGroup39='" + rspROPERTY.Fields.Item("QryGroup39").Value.ToString() + "',QryGroup40='" + rspROPERTY.Fields.Item("QryGroup40").Value.ToString() + "',QryGroup41='" + rspROPERTY.Fields.Item("QryGroup41").Value.ToString() + "',QryGroup42='" + rspROPERTY.Fields.Item("QryGroup42").Value.ToString() + "',QryGroup43='" + rspROPERTY.Fields.Item("QryGroup43").Value.ToString() + "',QryGroup44='" + rspROPERTY.Fields.Item("QryGroup44").Value.ToString() + "',QryGroup45='" + rspROPERTY.Fields.Item("QryGroup45").Value.ToString() + "',QryGroup46='" + rspROPERTY.Fields.Item("QryGroup46").Value.ToString() + "',QryGroup47='" + rspROPERTY.Fields.Item("QryGroup47").Value.ToString() + "',QryGroup48='" + rspROPERTY.Fields.Item("QryGroup48").Value.ToString() + "',QryGroup49='" + rspROPERTY.Fields.Item("QryGroup49").Value.ToString() + "',QryGroup50='" + rspROPERTY.Fields.Item("QryGroup50").Value.ToString() + "',QryGroup51='" + rspROPERTY.Fields.Item("QryGroup51").Value.ToString() + "',QryGroup52='" + rspROPERTY.Fields.Item("QryGroup52").Value.ToString() + "',QryGroup53='" + rspROPERTY.Fields.Item("QryGroup53").Value.ToString() + "',QryGroup54='" + rspROPERTY.Fields.Item("QryGroup54").Value.ToString() + "',QryGroup55='" + rspROPERTY.Fields.Item("QryGroup55").Value.ToString() + "',QryGroup56='" + rspROPERTY.Fields.Item("QryGroup56").Value.ToString() + "',QryGroup57='" + rspROPERTY.Fields.Item("QryGroup57").Value.ToString() + "',QryGroup58='" + rspROPERTY.Fields.Item("QryGroup58").Value.ToString() + "',QryGroup59='" + rspROPERTY.Fields.Item("QryGroup59").Value.ToString() + "',QryGroup60='" + rspROPERTY.Fields.Item("QryGroup60").Value.ToString() + "',QryGroup61='" + rspROPERTY.Fields.Item("QryGroup61").Value.ToString() + "',QryGroup62='" + rspROPERTY.Fields.Item("QryGroup62").Value.ToString() + "',QryGroup63='" + rspROPERTY.Fields.Item("QryGroup63").Value.ToString() + "',QryGroup64='" + rspROPERTY.Fields.Item("QryGroup64").Value.ToString() + "' WHERE ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'";

                                                        rspROPERTYUpdate.DoQuery(QryPrtyUpdate);
                                                        //oItemUpdate.GetByKey(Element.GetElementsByTagName("ItemCode")[0].InnerText);
                                                        //for (int j = 1; j < strPrty.Length; j++)
                                                        //{
                                                        //    if (strPrty[j] != "" & strPrty[j] != null)
                                                        //    {
                                                        //        oItemUpdate.set_Properties(Convert.ToInt32(strPrty[j]), SAPbobsCOM.BoYesNoEnum.tYES);
                                                        //    }
                                                        //}
                                                        //oItemUpdate.Update();

                                                        //string _str_Query = "UPDATE [ITM1]  SET [Price] = '" + frmItem.DataSources.UserDataSources.Item("udsMRP").ValueEx + "' WHERE ItemCode = '" + strItemCod + "'AND PriceList = 2 ";
                                                        //oRecordSet.DoQuery(_str_Query);

                                                        #endregion


                                                        SAPbobsCOM.Recordset rspUpdateStatus = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        string strUpdateStatus = "Update [@NOR_OITM_UNIT] SET U_IsIntegrated='Y' WHERE U_ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and U_UnitCode = '" + strUnitCd + "'";
                                                        rspUpdateStatus.DoQuery(strUpdateStatus);
                                                        //"Operation Completed Successfully"
                                                        // MessageBox.Show("Operation Completed Successfully");
                                                        //Global.SapApplication.StatusBar.SetText("Operation Completed Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                   
                                }
                                rsCompany.MoveNext();
                            }
                            // return true;

                        }
                        else
                        {
                            MessageBox.Show("No Item Exists To Integrate");
                        }
                        
                    }
                    
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message+"Item master");
                    }

                
                }//ITEM MASTER 

                public void New_ITEMMASTER()
                {
                    try
                    {
                        string sPath = "";
                        string FileName = "Item.xml";
                        string StrSql = "";
                        string insertuser = "";
                        General g = new General();
                        if (!File.Exists(sPath + FileName))
                        { File.Create(sPath + FileName); }
                        int recCount = 0;
                        System.Xml.XmlDocument oXmlDoc = null;
                        string sXmlString = null;

                        string QRY11 = @"select distinct T2.*,T1.code [UnitGet] from [@NOR_OITM_UNIT] T0 
                                        inner join  [@NOR_UNITMASTER] T1  on T0.U_UnitCode=T1.code
                                        inner join [@NOR_BRANCH_DTL] T2 on T2.U_UnitId=T1.U_UnitCode 
                                           where T0.U_isintegrated='N' and T0.U_UnitCode is not null";
                        
                        //string QRY11 = "Select * from [@NOR_BRANCH_DTL] Where U_UnitId ='" + UnitGet + "'";
                        SAPbobsCOM.Recordset rsCompany1 = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                        rsCompany1.DoQuery(QRY11);
                        bool _connection = false;
                        while (!rsCompany1.EoF)
                        
                        {
                            string UnitGet = rsCompany1.Fields.Item("UnitGet").Value.ToString();
                            string Licserver = rsCompany1.Fields.Item("U_Licserver").Value.ToString();
                            string server = rsCompany1.Fields.Item("U_ServerName").Value.ToString();
                            string DB = rsCompany1.Fields.Item("U_CompanyDB").Value.ToString();
                            string sUser = rsCompany1.Fields.Item("U_SAPUserName").Value.ToString();
                            string sPass = rsCompany1.Fields.Item("U_SAPPassword").Value.ToString();
                            string sqUser = rsCompany1.Fields.Item("U_ServerUser").Value.ToString();
                            string sqPass = rsCompany1.Fields.Item("U_ServerPass").Value.ToString();                            
                            _connection = g.connectOtherCompany(server,Licserver, DB, sUser, sPass, sqUser, sqPass);
                            if (_connection == true)
                            {
                                MessageBox.Show("Branch Connected" + server + DB ); 
                                StrSql = @"SELECT OITM.ItemCode ItemCode,isnull(ItemName,'-')ItemName,oitm.ItmsGrpCod,B.U_UnitCode,
        isnull(CodeBars,0)[CodeBars],isnull(InvntryUom,0)[InvntryUom],ISNULL(BuyUnitMsr,'')[PurchaseUOM],ISNULL(SalUnitMsr,'')[SalesUOM],isnull(OITM.CardCode,'-')[CardCode],OITM.FirmCode,
        OITM.SellItem,OITM.InvntItem,OITM.PrchseItem,OITM.AssetItem,isnull(OITM.FrgnName,'-')[FrgnName],OITM.ItemType,
        isnull(OITM.NumInSale,1)[NumInSale],isnull(OITM.SalPackUn,1)SalPackUn,isnull(OITM.U_Brand,'')U_Brand,isnull(OITM.U_Category,'')U_Category,
        isnull(OITM.U_Class,'')U_Class,isnull(OITM.U_Color,'')U_Color,isnull(OITM.U_Export,'N')U_Export,isnull(OITM.U_GrpCode,'')U_GrpCode,isnull(OITM.U_Model,'')U_Model,
        isnull(OITM.U_NofPairs,'')U_NofPairs,isnull(OITM.U_OrdQty,0)U_OrdQty,isnull(OITM.U_PairSize,'')U_PairSize,isnull(OITM.U_Priority,0)U_Priority,isnull(OITM.U_SFGCat,'')U_SFGCat,
        isnull(OITM.U_Size,'')U_Size,isnull(OITM.U_SizeCat,'')U_SizeCat,isnull(OITM.U_StdSize,'')U_StdSize,isnull(OITM.U_Unit,'')U_Unit,OITM.GLMethod,ISNULL( OITM.U_HsnCcode,'')U_HsnCcode,isnull(oitm.U_VATCode,'')U_VATCode,ISNULL(OITM.U_VatRate,'')U_VatRate,ISNULL(OITM.U_JbType,'')U_JbType,ISNULL(OITM.U_JbType,'')U_JbType, isnull(oitm.U_Abtment,'')U_Abtment
        ,ISNULL(OITM.U_PrdCat,'')[ProductCatogry],ISNULL(OITM.U_PrdTyp,'')[ProductType],ISNULL(OITM.U_Procestype,'')[Processtype],ISNULL(OITM.U_OrderCode,'')[OrderCode],
        Isnull(OITM.U_ItmMrp,0)U_ItmMrp,isnull(OITM.GSTRelevnt,'N')[GSTRelevnt],isnull(OITM.GstTaxCtg,'')[GstTaxCtg],isnull(OITM.MatType,0)[MatType],isnull(OITM.ChapterID,-1)[ChapterID],isnull(c.ChapterID,'') [ChapterCode],
isnull(OITM.PrdStdCst,0)[PrdStdCst]
        from OITM
		inner join OITB on OITM.ItmsGrpCod = OITB.ItmsGrpCod
        INNER JOIN [@NOR_OITM_UNIT] B on OITM.ItemCode=B.U_ItemCode 
        left join OCHP C on C.AbsEntry=OITM.ChapterID 
        WHERE B.U_IsIntegrated='N' AND B.U_UnitCode IS NOT NULL
        and B.U_UnitCode='" + UnitGet  + "' order by OITM.ItemCode";
                                //added stdprdcst 21-Jun-2019 by Tamizh
                                DataSet objDataSet = ConDb.DbDataFromSAP(StrSql);
                        if (objDataSet.Tables[0].Rows.Count > 0)
                        {
                                sXmlString = objDataSet.GetXml();
                                // sXmlString = oRsInv.GetAsXML();
                                oXmlDoc = new System.Xml.XmlDocument();
                                oXmlDoc.LoadXml(sXmlString);
                                oXmlDoc.Save((sPath + FileName));



                                XmlDocument reader = new XmlDocument();
                                XmlDocument readerlines = new XmlDocument();
                                IFormatProvider ifp = new System.Globalization.CultureInfo("en-US", true);
                                reader.Load(sPath + FileName);

                                XmlNodeList list = reader.GetElementsByTagName(g.row1);

                                foreach (XmlNode node in list)
                                {
                                    XmlElement Element = (XmlElement)node;
                                    string strUnitCd = Element.GetElementsByTagName("U_UnitCode")[0].InnerText;
                               #region Item Master (ADD)
                                SAPbobsCOM.Items oItem = (SAPbobsCOM.Items)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                                string QryyExistChk = "SELECT 1 FROM OITM WHERE ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'";
                                SAPbobsCOM.Recordset rsItem = ((SAPbobsCOM.Recordset)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                rsItem.DoQuery(QryyExistChk);
                                    if (rsItem.RecordCount > 0)
                                            {
                                                oItem.GetByKey(Element.GetElementsByTagName("ItemCode")[0].InnerText);
                                                //oItem.ItemCode = Element.GetElementsByTagName("ItemCode")[0].InnerText;
                                                oItem.ItemName = Element.GetElementsByTagName("ItemName")[0].InnerText;
                                                oItem.BarCode = Element.GetElementsByTagName("CodeBars")[0].InnerText;
                                                oItem.ForeignName = Element.GetElementsByTagName("FrgnName")[0].InnerText;
                                                //oItem.Mainsupplier = Element.GetElementsByTagName("CardCode")[0].InnerText;
                                                // oItem.ItemType = Element.GetElementsByTagName("ItemType")[0].InnerText;
                                                // oItem.GLMethod = Element.GetElementsByTagName("GLMethod")[0].InnerText;

                                                //string _strWarehouse = Element.GetElementsByTagName("WhsCode")[0].InnerText;
                                        //commented as per imran mail 05-Nov-2019
                                              //  oItem.ProdStdCost = Convert.ToDouble(Element.GetElementsByTagName("PrdStdCst")[0].InnerText);// added by Tamizh 21-Jun-2019
                                                oItem.ItemsGroupCode = Convert.ToInt32(Element.GetElementsByTagName("ItmsGrpCod")[0].InnerText);
                                                oItem.GLMethod = SAPbobsCOM.BoGLMethods.glm_ItemClass;

                                                oItem.UserFields.Fields.Item("U_Class").Value = Element.GetElementsByTagName("U_Class")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Brand").Value = Element.GetElementsByTagName("U_Brand")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Category").Value = Element.GetElementsByTagName("U_Category")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Color").Value = Element.GetElementsByTagName("U_Color")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Export").Value = Element.GetElementsByTagName("U_Export")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_GrpCode").Value = Element.GetElementsByTagName("U_GrpCode")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Model").Value = Element.GetElementsByTagName("U_Model")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_NofPairs").Value = Element.GetElementsByTagName("U_NofPairs")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_OrdQty").Value = Element.GetElementsByTagName("U_OrdQty")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_PairSize").Value = Element.GetElementsByTagName("U_PairSize")[0].InnerText;
                                                //oItem.UserFields.Fields.Item("U_Priority").Value = Element.GetElementsByTagName("U_Priority")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_SFGCat").Value = Element.GetElementsByTagName("U_SFGCat")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Size").Value = Element.GetElementsByTagName("U_Size")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_SizeCat").Value = Element.GetElementsByTagName("U_SizeCat")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_StdSize").Value = Element.GetElementsByTagName("U_StdSize")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Unit").Value = Element.GetElementsByTagName("U_Unit")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_HsnCcode").Value = Element.GetElementsByTagName("U_HsnCcode")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_VATCode").Value = Element.GetElementsByTagName("U_VATCode")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_VatRate").Value = Element.GetElementsByTagName("U_VatRate")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_JbType").Value = Element.GetElementsByTagName("U_JbType")[0].InnerText;

                                                oItem.UserFields.Fields.Item("U_PrdCat").Value = Element.GetElementsByTagName("ProductCatogry")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_PrdTyp").Value = Element.GetElementsByTagName("ProductType")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_Procestype").Value = Element.GetElementsByTagName("Processtype")[0].InnerText;
                                                oItem.UserFields.Fields.Item("U_OrderCode").Value = Element.GetElementsByTagName("OrderCode")[0].InnerText;

                                                oItem.InventoryUOM = Element.GetElementsByTagName("InvntryUom")[0].InnerText;
                                                oItem.PurchaseUnit = Element.GetElementsByTagName("PurchaseUOM")[0].InnerText;
                                                oItem.SalesUnit = Element.GetElementsByTagName("SalesUOM")[0].InnerText;

                                                oItem.UserFields.Fields.Item("U_ItmMrp").Value = Element.GetElementsByTagName("U_ItmMrp")[0].InnerText;
                                                if (Element.GetElementsByTagName("GSTRelevnt")[0].InnerText.ToString().ToUpper() == "Y")
                                                {
                                                    oItem.GSTRelevnt = SAPbobsCOM.BoYesNoEnum.tYES;
                                                    //if (Element.GetElementsByTagName("ChapterID")[0].InnerText.ToString().ToUpper() != "-1") 
                                                    //{
                                                        string qrychapterid = "select AbsEntry  from ochp where ChapterID='" + Element.GetElementsByTagName("ChapterCode")[0].InnerText + "'";
                                                        SAPbobsCOM.Recordset rschapterid = ((SAPbobsCOM.Recordset)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                                        rschapterid.DoQuery(qrychapterid);
                                                        if (rschapterid.RecordCount > 0) { oItem.ChapterID = rschapterid.Fields.Item("AbsEntry").Value;}
                                                    //}

                                                    if (Element.GetElementsByTagName("MatType")[0].InnerText.ToString().ToUpper() == "1") { oItem.MaterialType = SAPbobsCOM.BoMaterialTypes.mt_FinishedGoods; }
                                                    if (Element.GetElementsByTagName("MatType")[0].InnerText.ToString().ToUpper() == "2") { oItem.MaterialType = SAPbobsCOM.BoMaterialTypes.mt_GoodsInProcess; }
                                                    if (Element.GetElementsByTagName("MatType")[0].InnerText.ToString().ToUpper() == "3") { oItem.MaterialType = SAPbobsCOM.BoMaterialTypes.mt_RawMaterial; }

                                                    if (Element.GetElementsByTagName("GstTaxCtg")[0].InnerText.ToString().ToUpper() == "N") { oItem.GSTTaxCategory = SAPbobsCOM.GSTTaxCategoryEnum.gtc_NilRated; }
                                                    if (Element.GetElementsByTagName("GstTaxCtg")[0].InnerText.ToString().ToUpper() == "R") { oItem.GSTTaxCategory = SAPbobsCOM.GSTTaxCategoryEnum.gtc_Regular; }
                                                    if (Element.GetElementsByTagName("GstTaxCtg")[0].InnerText.ToString().ToUpper() == "E") { oItem.GSTTaxCategory = SAPbobsCOM.GSTTaxCategoryEnum.gtc_Exempt; }
                                                }
                                                else {oItem.GSTRelevnt = SAPbobsCOM.BoYesNoEnum.tNO;}                                                
                                                                                
                                                int iError = 0;
                                                iError = oItem.Update();
                                                if (iError != 0)
                                                {
                                                    string sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                                    MessageBox.Show(sErrorMsg);  
                                                    //return false;
                                                }
                                                else
                                                {

                                                    #region ItemPriceList
                                                    SAPbobsCOM.Recordset rspricelist = (SAPbobsCOM.Recordset)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                    string strItemPrclistQry = @"select PriceList,Price from ITM1 
                                                                                 inner join OITM on OITM.ItemCode=ITM1.ItemCode
                                                                                 where OITM.ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and Price is not NULL";



                                                    DataSet datasetPrclist = ConDb.DbDataFromSAP(strItemPrclistQry);
                                                    for (int i = 0; i < datasetPrclist.Tables[0].Rows.Count; i++)
                                                    {
                                                        string strUpdate = "update ITM1 set Price='" + datasetPrclist.Tables[0].Rows[i][1].ToString() + "' where ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and PriceList='" + datasetPrclist.Tables[0].Rows[i][0].ToString() + "'";
                                                        rspricelist.DoQuery(strUpdate);

                                                    }


                                                    #endregion


                                                    #region Item Property(UPDATE)
                                                    SAPbobsCOM.Items oItemUpdate = (SAPbobsCOM.Items)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                                                    string QryPropery = @"SELECT OITM.QryGroup1,OITM.QryGroup2,OITM.QryGroup3,
        OITM.QryGroup4,OITM.QryGroup5,OITM.QryGroup6,OITM.QryGroup7,OITM.QryGroup8,OITM.QryGroup9,OITM.QryGroup10,OITM.QryGroup11,OITM.QryGroup12,OITM.QryGroup13,
        OITM.QryGroup14,OITM.QryGroup15,OITM.QryGroup16,OITM.QryGroup17,OITM.QryGroup18,OITM.QryGroup19,OITM.QryGroup20,OITM.QryGroup21,OITM.QryGroup22,OITM.QryGroup23,
        OITM.QryGroup24,OITM.QryGroup25,OITM.QryGroup26,OITM.QryGroup27,OITM.QryGroup28,OITM.QryGroup29,OITM.QryGroup30,OITM.QryGroup31,OITM.QryGroup32,OITM.QryGroup33,
        OITM.QryGroup34,OITM.QryGroup35,OITM.QryGroup36,OITM.QryGroup37,OITM.QryGroup38,OITM.QryGroup39,OITM.QryGroup40,OITM.QryGroup41,OITM.QryGroup42,OITM.QryGroup43,OITM.QryGroup44,
        OITM.QryGroup45,OITM.QryGroup46,OITM.QryGroup47,OITM.QryGroup48,OITM.QryGroup49,OITM.QryGroup50,OITM.QryGroup51,OITM.QryGroup52,OITM.QryGroup53,OITM.QryGroup54,OITM.QryGroup55,
        OITM.QryGroup56,OITM.QryGroup57,OITM.QryGroup58,OITM.QryGroup59,OITM.QryGroup60,OITM.QryGroup61,OITM.QryGroup62,OITM.QryGroup63,OITM.QryGroup64 FROM OITM WHERE ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'";
                                                    SAPbobsCOM.Recordset rspROPERTY = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                    SAPbobsCOM.Recordset rspROPERTYUpdate = (SAPbobsCOM.Recordset)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                    //SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                                    rspROPERTY.DoQuery(QryPropery);
                                                    string QryPrtyUpdate = @"UPDATE OITM  SET QryGroup1='" + rspROPERTY.Fields.Item("QryGroup1").Value.ToString() + "',QryGroup2='" + rspROPERTY.Fields.Item("QryGroup2").Value.ToString() + "',QryGroup3='" + rspROPERTY.Fields.Item("QryGroup3").Value.ToString() + "',QryGroup4='" + rspROPERTY.Fields.Item("QryGroup4").Value.ToString() + "',QryGroup5='" + rspROPERTY.Fields.Item("QryGroup5").Value.ToString() + "',QryGroup6='" + rspROPERTY.Fields.Item("QryGroup6").Value.ToString() + "',QryGroup7='" + rspROPERTY.Fields.Item("QryGroup7").Value.ToString() + "',QryGroup8='" + rspROPERTY.Fields.Item("QryGroup8").Value.ToString() + "',QryGroup9='" + rspROPERTY.Fields.Item("QryGroup9").Value.ToString() + "',QryGroup10='" + rspROPERTY.Fields.Item("QryGroup10").Value.ToString() + "',QryGroup11='" + rspROPERTY.Fields.Item("QryGroup11").Value.ToString() + "',QryGroup12='" + rspROPERTY.Fields.Item("QryGroup12").Value.ToString() + "',QryGroup13='" + rspROPERTY.Fields.Item("QryGroup13").Value.ToString() + "',QryGroup14='" + rspROPERTY.Fields.Item("QryGroup14").Value.ToString() + "',QryGroup15='" + rspROPERTY.Fields.Item("QryGroup15").Value.ToString() + "',QryGroup16='" + rspROPERTY.Fields.Item("QryGroup16").Value.ToString() + "',QryGroup17='" + rspROPERTY.Fields.Item("QryGroup17").Value.ToString() + "',QryGroup18='" + rspROPERTY.Fields.Item("QryGroup18").Value.ToString() + "',QryGroup19='" + rspROPERTY.Fields.Item("QryGroup19").Value.ToString() + "',QryGroup20='" + rspROPERTY.Fields.Item("QryGroup20").Value.ToString() + "',QryGroup21='" + rspROPERTY.Fields.Item("QryGroup21").Value.ToString() + "',QryGroup22='" + rspROPERTY.Fields.Item("QryGroup22").Value.ToString() + "',QryGroup23='" + rspROPERTY.Fields.Item("QryGroup23").Value.ToString() + "',QryGroup24='" + rspROPERTY.Fields.Item("QryGroup24").Value.ToString() + "',QryGroup25='" + rspROPERTY.Fields.Item("QryGroup25").Value.ToString() + "',QryGroup26='" + rspROPERTY.Fields.Item("QryGroup26").Value.ToString() + "',QryGroup27='" + rspROPERTY.Fields.Item("QryGroup27").Value.ToString() + "',QryGroup28='" + rspROPERTY.Fields.Item("QryGroup28").Value.ToString() + "',QryGroup29='" + rspROPERTY.Fields.Item("QryGroup29").Value.ToString() + "',QryGroup30='" + rspROPERTY.Fields.Item("QryGroup30").Value.ToString() + "',QryGroup31='" + rspROPERTY.Fields.Item("QryGroup31").Value.ToString() + "',QryGroup32='" + rspROPERTY.Fields.Item("QryGroup32").Value.ToString() + "',QryGroup33='" + rspROPERTY.Fields.Item("QryGroup33").Value.ToString() + "',QryGroup34='" + rspROPERTY.Fields.Item("QryGroup34").Value.ToString() + "',QryGroup35='" + rspROPERTY.Fields.Item("QryGroup35").Value.ToString() + "',QryGroup36='" + rspROPERTY.Fields.Item("QryGroup36").Value.ToString() + "',QryGroup37='" + rspROPERTY.Fields.Item("QryGroup37").Value.ToString() + "',QryGroup38='" + rspROPERTY.Fields.Item("QryGroup38").Value.ToString() + "',QryGroup39='" + rspROPERTY.Fields.Item("QryGroup39").Value.ToString() + "',QryGroup40='" + rspROPERTY.Fields.Item("QryGroup40").Value.ToString() + "',QryGroup41='" + rspROPERTY.Fields.Item("QryGroup41").Value.ToString() + "',QryGroup42='" + rspROPERTY.Fields.Item("QryGroup42").Value.ToString() + "',QryGroup43='" + rspROPERTY.Fields.Item("QryGroup43").Value.ToString() + "',QryGroup44='" + rspROPERTY.Fields.Item("QryGroup44").Value.ToString() + "',QryGroup45='" + rspROPERTY.Fields.Item("QryGroup45").Value.ToString() + "',QryGroup46='" + rspROPERTY.Fields.Item("QryGroup46").Value.ToString() + "',QryGroup47='" + rspROPERTY.Fields.Item("QryGroup47").Value.ToString() + "',QryGroup48='" + rspROPERTY.Fields.Item("QryGroup48").Value.ToString() + "',QryGroup49='" + rspROPERTY.Fields.Item("QryGroup49").Value.ToString() + "',QryGroup50='" + rspROPERTY.Fields.Item("QryGroup50").Value.ToString() + "',QryGroup51='" + rspROPERTY.Fields.Item("QryGroup51").Value.ToString() + "',QryGroup52='" + rspROPERTY.Fields.Item("QryGroup52").Value.ToString() + "',QryGroup53='" + rspROPERTY.Fields.Item("QryGroup53").Value.ToString() + "',QryGroup54='" + rspROPERTY.Fields.Item("QryGroup54").Value.ToString() + "',QryGroup55='" + rspROPERTY.Fields.Item("QryGroup55").Value.ToString() + "',QryGroup56='" + rspROPERTY.Fields.Item("QryGroup56").Value.ToString() + "',QryGroup57='" + rspROPERTY.Fields.Item("QryGroup57").Value.ToString() + "',QryGroup58='" + rspROPERTY.Fields.Item("QryGroup58").Value.ToString() + "',QryGroup59='" + rspROPERTY.Fields.Item("QryGroup59").Value.ToString() + "',QryGroup60='" + rspROPERTY.Fields.Item("QryGroup60").Value.ToString() + "',QryGroup61='" + rspROPERTY.Fields.Item("QryGroup61").Value.ToString() + "',QryGroup62='" + rspROPERTY.Fields.Item("QryGroup62").Value.ToString() + "',QryGroup63='" + rspROPERTY.Fields.Item("QryGroup63").Value.ToString() + "',QryGroup64='" + rspROPERTY.Fields.Item("QryGroup64").Value.ToString() + "' WHERE ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'";

                                                    rspROPERTYUpdate.DoQuery(QryPrtyUpdate);
                                                    //oItemUpdate.GetByKey(Element.GetElementsByTagName("ItemCode")[0].InnerText);
                                                    //for (int j = 1; j < strPrty.Length; j++)
                                                    //{
                                                    //    if (strPrty[j] != "" & strPrty[j] != null)
                                                    //    {
                                                    //        oItemUpdate.set_Properties(Convert.ToInt32(strPrty[j]), SAPbobsCOM.BoYesNoEnum.tYES);
                                                    //    }
                                                    //}
                                                    //oItemUpdate.Update();

                                                    //string _str_Query = "UPDATE [ITM1]  SET [Price] = '" + frmItem.DataSources.UserDataSources.Item("udsMRP").ValueEx + "' WHERE ItemCode = '" + strItemCod + "'AND PriceList = 2 ";
                                                    //oRecordSet.DoQuery(_str_Query);

                                                    #endregion

                                                    SAPbobsCOM.Recordset rspUpdateStatus = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                    string strUpdateStatus = "Update [@NOR_OITM_UNIT] SET U_IsIntegrated='Y' WHERE U_ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and U_UnitCode = '" + strUnitCd + "'";
                                                    rspUpdateStatus.DoQuery(strUpdateStatus);

                                                    // Global.SapApplication.StatusBar.SetText("Operation Completed Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                                }  // return true;

                                            }
                                            else
                                            {// Item add
                                                SAPbobsCOM.Recordset rsCompanyPricelist = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                string _strClass = Element.GetElementsByTagName("U_Class")[0].InnerText;
                                                //string _strWarehouse = Element.GetElementsByTagName("WhsCode")[0].InnerText;
                                                //if ((_strClass == "2" || _strClass == "3") && _strWarehouse == "0")
                                                //{
                                                //    MessageBox.Show("Please select warehouse for Item '" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'");
                                                //}
                                                //else
                                                //{
                                                    oItem.ItemCode = Element.GetElementsByTagName("ItemCode")[0].InnerText;

                                                    oItem.ItemName = Element.GetElementsByTagName("ItemName")[0].InnerText;
                                                    oItem.BarCode = Element.GetElementsByTagName("CodeBars")[0].InnerText;
                                                    oItem.ForeignName = Element.GetElementsByTagName("FrgnName")[0].InnerText;
                                                    // oItem.Mainsupplier = Element.GetElementsByTagName("CardCode")[0].InnerText;
                                                    // oItem.ItemType=Element.GetElementsByTagName("ItemType")[0].InnerText;


                                                    oItem.ItemsGroupCode = Convert.ToInt32(Element.GetElementsByTagName("ItmsGrpCod")[0].InnerText);
                                                    // oItem.GLMethod = Element.GetElementsByTagName("GLMethod")[0].InnerText.ToString();
                                                    oItem.GLMethod = SAPbobsCOM.BoGLMethods.glm_ItemClass;
                                                    oItem.Manufacturer = Convert.ToInt32(Element.GetElementsByTagName("FirmCode")[0].InnerText);

                                                    oItem.InventoryUOM = Element.GetElementsByTagName("InvntryUom")[0].InnerText;
                                                    string strWrhse = "select OITW.WhsCode from OITW inner join OITM on OITW.ItemCode =OITM.ItemCode inner join OWHS on OITW.WhsCode=OWHS.WhsCode where oitm.ItemCode = '" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and owhs.U_Unit='" + UnitGet + "'";
                                                    DataSet DatasetWrhse = ConDb.DbDataFromSAP(strWrhse);


                                                    //if (_strWarehouse != "0")
                                                    //{

                                                        for (int i = 0; i < DatasetWrhse.Tables[0].Rows.Count; i++)
                                                        {
                                                            oItem.WhsInfo.WarehouseCode = DatasetWrhse.Tables[0].Rows[i]["WhsCode"].ToString();
                                                            oItem.WhsInfo.Add();
                                                        }
                                                    //}

                                                    //-------------User Fields-------------------------------------------------------------------//
                                                    oItem.UserFields.Fields.Item("U_Class").Value = Element.GetElementsByTagName("U_Class")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Brand").Value = Element.GetElementsByTagName("U_Brand")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Category").Value = Element.GetElementsByTagName("U_Category")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Color").Value = Element.GetElementsByTagName("U_Color")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Export").Value = Element.GetElementsByTagName("U_Export")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_GrpCode").Value = Element.GetElementsByTagName("U_GrpCode")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Model").Value = Element.GetElementsByTagName("U_Model")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_NofPairs").Value = Element.GetElementsByTagName("U_NofPairs")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_OrdQty").Value = Element.GetElementsByTagName("U_OrdQty")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_PairSize").Value = Element.GetElementsByTagName("U_PairSize")[0].InnerText;
                                                   // oItem.UserFields.Fields.Item("U_Priority").Value = Element.GetElementsByTagName("U_Priority")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_SFGCat").Value = Element.GetElementsByTagName("U_SFGCat")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Size").Value = Element.GetElementsByTagName("U_Size")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_SizeCat").Value = Element.GetElementsByTagName("U_SizeCat")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_StdSize").Value = Element.GetElementsByTagName("U_StdSize")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Unit").Value = Element.GetElementsByTagName("U_Unit")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_HsnCcode").Value = Element.GetElementsByTagName("U_HsnCcode")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_VATCode").Value = Element.GetElementsByTagName("U_VATCode")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_VatRate").Value = Element.GetElementsByTagName("U_VatRate")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_JbType").Value = Element.GetElementsByTagName("U_JbType")[0].InnerText;

                                                    oItem.UserFields.Fields.Item("U_PrdCat").Value = Element.GetElementsByTagName("ProductCatogry")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_PrdTyp").Value = Element.GetElementsByTagName("ProductType")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_Procestype").Value = Element.GetElementsByTagName("Processtype")[0].InnerText;
                                                    oItem.UserFields.Fields.Item("U_OrderCode").Value = Element.GetElementsByTagName("OrderCode")[0].InnerText;
                                                    //-------------------------------updated by Tamizh 23/03/2019-----------------------------------------------//
                                                    oItem.InventoryUOM = Element.GetElementsByTagName("InvntryUom")[0].InnerText;
                                                    oItem.PurchaseUnit = Element.GetElementsByTagName("PurchaseUOM")[0].InnerText;
                                                    oItem.SalesUnit = Element.GetElementsByTagName("SalesUOM")[0].InnerText;
                                                 //   oItem.ProdStdCost = Convert.ToDouble(Element.GetElementsByTagName("PrdStdCst")[0].InnerText);// added by Tamizh 21-Jun-2019
                                        // ------------------------------------------------commented by Tamizh as per imran's recommendation----------------------------------

                                                    //if (Element.GetElementsByTagName("InvntItem")[0].InnerText.ToString() == "Y")
                                                    //{
                                                    //    oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                                    //                                                        }
                                                    //else
                                                    //{
                                                    //    oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                                    //}
                                                    //if (Element.GetElementsByTagName("SellItem")[0].InnerText.ToString() == "Y")
                                                    //{
                                                    //    oItem.SalesItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                                    //}
                                                    //else
                                                    //{
                                                    //    oItem.SalesItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                                    //}
                                                    //if (Element.GetElementsByTagName("PrchseItem")[0].InnerText.ToString() == "Y")
                                                    //{
                                                    //    oItem.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                                    //}
                                                    //else
                                                    //{
                                                    //    oItem.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                                    //}
                                                    if (Element.GetElementsByTagName("AssetItem")[0].InnerText.ToString() == "Y")
                                                    {
                                                        oItem.AssetItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                                    }
                                                    else
                                                    {
                                                        oItem.AssetItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                                    }

                                                    oItem.CostAccountingMethod = SAPbobsCOM.BoInventorySystem.bis_FIFO;

                                                    oItem.UserFields.Fields.Item("U_ItmMrp").Value = Element.GetElementsByTagName("U_ItmMrp")[0].InnerText;
                                                    if (Element.GetElementsByTagName("GSTRelevnt")[0].InnerText.ToString().ToUpper() == "Y")
                                                    {
                                                        oItem.GSTRelevnt = SAPbobsCOM.BoYesNoEnum.tYES;
                                                        //if (Element.GetElementsByTagName("ChapterID")[0].InnerText.ToString().ToUpper() != "-1") { oItem.ChapterID = Int32.Parse(Element.GetElementsByTagName("ChapterID")[0].InnerText.ToString()); }
                                                        if (Element.GetElementsByTagName("ChapterID")[0].InnerText.ToString().ToUpper() != "-1")
                                                        {
                                                            string qrychapterid = "select AbsEntry  from ochp where ChapterID='" + Element.GetElementsByTagName("ChapterCode")[0].InnerText + "'";
                                                            SAPbobsCOM.Recordset rschapterid = ((SAPbobsCOM.Recordset)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                                            rschapterid.DoQuery(qrychapterid);
                                                            if (rschapterid.RecordCount > 0) { oItem.ChapterID = rschapterid.Fields.Item("AbsEntry").Value; }
                                                        }

                                                        if (Element.GetElementsByTagName("MatType")[0].InnerText.ToString().ToUpper() == "1") { oItem.MaterialType = SAPbobsCOM.BoMaterialTypes.mt_FinishedGoods; }
                                                        if (Element.GetElementsByTagName("MatType")[0].InnerText.ToString().ToUpper() == "2") { oItem.MaterialType = SAPbobsCOM.BoMaterialTypes.mt_GoodsInProcess; }
                                                        if (Element.GetElementsByTagName("MatType")[0].InnerText.ToString().ToUpper() == "3") { oItem.MaterialType = SAPbobsCOM.BoMaterialTypes.mt_RawMaterial; }

                                                        if (Element.GetElementsByTagName("GstTaxCtg")[0].InnerText.ToString().ToUpper() == "N") { oItem.GSTTaxCategory = SAPbobsCOM.GSTTaxCategoryEnum.gtc_NilRated; }
                                                        if (Element.GetElementsByTagName("GstTaxCtg")[0].InnerText.ToString().ToUpper() == "R") { oItem.GSTTaxCategory = SAPbobsCOM.GSTTaxCategoryEnum.gtc_Regular; }
                                                        if (Element.GetElementsByTagName("GstTaxCtg")[0].InnerText.ToString().ToUpper() == "E") { oItem.GSTTaxCategory = SAPbobsCOM.GSTTaxCategoryEnum.gtc_Exempt; }
                                                    }
                                                    else { oItem.GSTRelevnt = SAPbobsCOM.BoYesNoEnum.tNO; }       


                                                    int iError = 0;
                                                    iError = oItem.Add();
                                                    if (iError != 0)
                                                    {
                                                        string sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                                        MessageBox.Show(sErrorMsg + "Item master");
                                                        // Global.SapApplication.StatusBar.SetText(sErrorMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                                        //return false;
                                                    }
                                    #endregion
                                                    else
                                                    {
                                                        #region ItemPriceList
                                                        SAPbobsCOM.Recordset rspricelist = (SAPbobsCOM.Recordset)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        string strItemPrclistQry = @"select PriceList,Price from ITM1 
                                                                                 inner join OITM on OITM.ItemCode=ITM1.ItemCode
                                                                                 where OITM.ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and Price is not NULL";



                                                        DataSet datasetPrclist = ConDb.DbDataFromSAP(strItemPrclistQry);
                                                        for (int i = 0; i < datasetPrclist.Tables[0].Rows.Count; i++)
                                                        {
                                                            string strUpdate = "update ITM1 set Price='" + datasetPrclist.Tables[0].Rows[i][1].ToString() + "' where ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and PriceList='" + datasetPrclist.Tables[0].Rows[i][0].ToString() + "'";
                                                            rspricelist.DoQuery(strUpdate);

                                                        }
                                                        #endregion

                                                        #region Item Property(UPDATE)
                                                        SAPbobsCOM.Items oItemUpdate = (SAPbobsCOM.Items)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                                                        string QryPropery = @"SELECT OITM.QryGroup1,OITM.QryGroup2,OITM.QryGroup3,
        OITM.QryGroup4,OITM.QryGroup5,OITM.QryGroup6,OITM.QryGroup7,OITM.QryGroup8,OITM.QryGroup9,OITM.QryGroup10,OITM.QryGroup11,OITM.QryGroup12,OITM.QryGroup13,
        OITM.QryGroup14,OITM.QryGroup15,OITM.QryGroup16,OITM.QryGroup17,OITM.QryGroup18,OITM.QryGroup19,OITM.QryGroup20,OITM.QryGroup21,OITM.QryGroup22,OITM.QryGroup23,
        OITM.QryGroup24,OITM.QryGroup25,OITM.QryGroup26,OITM.QryGroup27,OITM.QryGroup28,OITM.QryGroup29,OITM.QryGroup30,OITM.QryGroup31,OITM.QryGroup32,OITM.QryGroup33,
        OITM.QryGroup34,OITM.QryGroup35,OITM.QryGroup36,OITM.QryGroup37,OITM.QryGroup38,OITM.QryGroup39,OITM.QryGroup40,OITM.QryGroup41,OITM.QryGroup42,OITM.QryGroup43,OITM.QryGroup44,
        OITM.QryGroup45,OITM.QryGroup46,OITM.QryGroup47,OITM.QryGroup48,OITM.QryGroup49,OITM.QryGroup50,OITM.QryGroup51,OITM.QryGroup52,OITM.QryGroup53,OITM.QryGroup54,OITM.QryGroup55,
        OITM.QryGroup56,OITM.QryGroup57,OITM.QryGroup58,OITM.QryGroup59,OITM.QryGroup60,OITM.QryGroup61,OITM.QryGroup62,OITM.QryGroup63,OITM.QryGroup64 FROM OITM WHERE ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'";
                                                        SAPbobsCOM.Recordset rspROPERTY = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        SAPbobsCOM.Recordset rspROPERTYUpdate = (SAPbobsCOM.Recordset)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        //SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                                        rspROPERTY.DoQuery(QryPropery);
                                                        string QryPrtyUpdate = @"UPDATE OITM  SET QryGroup1='" + rspROPERTY.Fields.Item("QryGroup1").Value.ToString() + "',QryGroup2='" + rspROPERTY.Fields.Item("QryGroup2").Value.ToString() + "',QryGroup3='" + rspROPERTY.Fields.Item("QryGroup3").Value.ToString() + "',QryGroup4='" + rspROPERTY.Fields.Item("QryGroup4").Value.ToString() + "',QryGroup5='" + rspROPERTY.Fields.Item("QryGroup5").Value.ToString() + "',QryGroup6='" + rspROPERTY.Fields.Item("QryGroup6").Value.ToString() + "',QryGroup7='" + rspROPERTY.Fields.Item("QryGroup7").Value.ToString() + "',QryGroup8='" + rspROPERTY.Fields.Item("QryGroup8").Value.ToString() + "',QryGroup9='" + rspROPERTY.Fields.Item("QryGroup9").Value.ToString() + "',QryGroup10='" + rspROPERTY.Fields.Item("QryGroup10").Value.ToString() + "',QryGroup11='" + rspROPERTY.Fields.Item("QryGroup11").Value.ToString() + "',QryGroup12='" + rspROPERTY.Fields.Item("QryGroup12").Value.ToString() + "',QryGroup13='" + rspROPERTY.Fields.Item("QryGroup13").Value.ToString() + "',QryGroup14='" + rspROPERTY.Fields.Item("QryGroup14").Value.ToString() + "',QryGroup15='" + rspROPERTY.Fields.Item("QryGroup15").Value.ToString() + "',QryGroup16='" + rspROPERTY.Fields.Item("QryGroup16").Value.ToString() + "',QryGroup17='" + rspROPERTY.Fields.Item("QryGroup17").Value.ToString() + "',QryGroup18='" + rspROPERTY.Fields.Item("QryGroup18").Value.ToString() + "',QryGroup19='" + rspROPERTY.Fields.Item("QryGroup19").Value.ToString() + "',QryGroup20='" + rspROPERTY.Fields.Item("QryGroup20").Value.ToString() + "',QryGroup21='" + rspROPERTY.Fields.Item("QryGroup21").Value.ToString() + "',QryGroup22='" + rspROPERTY.Fields.Item("QryGroup22").Value.ToString() + "',QryGroup23='" + rspROPERTY.Fields.Item("QryGroup23").Value.ToString() + "',QryGroup24='" + rspROPERTY.Fields.Item("QryGroup24").Value.ToString() + "',QryGroup25='" + rspROPERTY.Fields.Item("QryGroup25").Value.ToString() + "',QryGroup26='" + rspROPERTY.Fields.Item("QryGroup26").Value.ToString() + "',QryGroup27='" + rspROPERTY.Fields.Item("QryGroup27").Value.ToString() + "',QryGroup28='" + rspROPERTY.Fields.Item("QryGroup28").Value.ToString() + "',QryGroup29='" + rspROPERTY.Fields.Item("QryGroup29").Value.ToString() + "',QryGroup30='" + rspROPERTY.Fields.Item("QryGroup30").Value.ToString() + "',QryGroup31='" + rspROPERTY.Fields.Item("QryGroup31").Value.ToString() + "',QryGroup32='" + rspROPERTY.Fields.Item("QryGroup32").Value.ToString() + "',QryGroup33='" + rspROPERTY.Fields.Item("QryGroup33").Value.ToString() + "',QryGroup34='" + rspROPERTY.Fields.Item("QryGroup34").Value.ToString() + "',QryGroup35='" + rspROPERTY.Fields.Item("QryGroup35").Value.ToString() + "',QryGroup36='" + rspROPERTY.Fields.Item("QryGroup36").Value.ToString() + "',QryGroup37='" + rspROPERTY.Fields.Item("QryGroup37").Value.ToString() + "',QryGroup38='" + rspROPERTY.Fields.Item("QryGroup38").Value.ToString() + "',QryGroup39='" + rspROPERTY.Fields.Item("QryGroup39").Value.ToString() + "',QryGroup40='" + rspROPERTY.Fields.Item("QryGroup40").Value.ToString() + "',QryGroup41='" + rspROPERTY.Fields.Item("QryGroup41").Value.ToString() + "',QryGroup42='" + rspROPERTY.Fields.Item("QryGroup42").Value.ToString() + "',QryGroup43='" + rspROPERTY.Fields.Item("QryGroup43").Value.ToString() + "',QryGroup44='" + rspROPERTY.Fields.Item("QryGroup44").Value.ToString() + "',QryGroup45='" + rspROPERTY.Fields.Item("QryGroup45").Value.ToString() + "',QryGroup46='" + rspROPERTY.Fields.Item("QryGroup46").Value.ToString() + "',QryGroup47='" + rspROPERTY.Fields.Item("QryGroup47").Value.ToString() + "',QryGroup48='" + rspROPERTY.Fields.Item("QryGroup48").Value.ToString() + "',QryGroup49='" + rspROPERTY.Fields.Item("QryGroup49").Value.ToString() + "',QryGroup50='" + rspROPERTY.Fields.Item("QryGroup50").Value.ToString() + "',QryGroup51='" + rspROPERTY.Fields.Item("QryGroup51").Value.ToString() + "',QryGroup52='" + rspROPERTY.Fields.Item("QryGroup52").Value.ToString() + "',QryGroup53='" + rspROPERTY.Fields.Item("QryGroup53").Value.ToString() + "',QryGroup54='" + rspROPERTY.Fields.Item("QryGroup54").Value.ToString() + "',QryGroup55='" + rspROPERTY.Fields.Item("QryGroup55").Value.ToString() + "',QryGroup56='" + rspROPERTY.Fields.Item("QryGroup56").Value.ToString() + "',QryGroup57='" + rspROPERTY.Fields.Item("QryGroup57").Value.ToString() + "',QryGroup58='" + rspROPERTY.Fields.Item("QryGroup58").Value.ToString() + "',QryGroup59='" + rspROPERTY.Fields.Item("QryGroup59").Value.ToString() + "',QryGroup60='" + rspROPERTY.Fields.Item("QryGroup60").Value.ToString() + "',QryGroup61='" + rspROPERTY.Fields.Item("QryGroup61").Value.ToString() + "',QryGroup62='" + rspROPERTY.Fields.Item("QryGroup62").Value.ToString() + "',QryGroup63='" + rspROPERTY.Fields.Item("QryGroup63").Value.ToString() + "',QryGroup64='" + rspROPERTY.Fields.Item("QryGroup64").Value.ToString() + "' WHERE ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'";

                                                        rspROPERTYUpdate.DoQuery(QryPrtyUpdate);
                                                        //oItemUpdate.GetByKey(Element.GetElementsByTagName("ItemCode")[0].InnerText);
                                                        //for (int j = 1; j < strPrty.Length; j++)
                                                        //{
                                                        //    if (strPrty[j] != "" & strPrty[j] != null)
                                                        //    {
                                                        //        oItemUpdate.set_Properties(Convert.ToInt32(strPrty[j]), SAPbobsCOM.BoYesNoEnum.tYES);
                                                        //    }
                                                        //}
                                                        //oItemUpdate.Update();

                                                        //string _str_Query = "UPDATE [ITM1]  SET [Price] = '" + frmItem.DataSources.UserDataSources.Item("udsMRP").ValueEx + "' WHERE ItemCode = '" + strItemCod + "'AND PriceList = 2 ";
                                                        //oRecordSet.DoQuery(_str_Query);

                                                        #endregion
                                                        
                                                        SAPbobsCOM.Recordset rspUpdateStatus = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        string strUpdateStatus = "Update [@NOR_OITM_UNIT] SET U_IsIntegrated='Y' WHERE U_ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and U_UnitCode = '" + strUnitCd + "'";
                                                        rspUpdateStatus.DoQuery(strUpdateStatus);
                                                        //"Operation Completed Successfully"
                                                        // MessageBox.Show("Operation Completed Successfully");
                                                        //Global.SapApplication.StatusBar.SetText("Operation Completed Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                                    }
                                                //}
                                            }
                                }
                                
                           }  
                            }
                            rsCompany1.MoveNext ();
                        }
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show("" + Ex);    
                        throw;
                    }

                    ///End of New_Itemmaster
                    

                }//ITEM MASTER 
         #endregion

        public void PRICELIST()
        {
            try
            {
                string sPath = "";
                string FileName = "ItemPriceList.xml";
                string StrSql = "";
                string insertuser = "";
                General g = new General();
                if (!File.Exists(sPath + FileName))
                { File.Create(sPath + FileName); }
                int recCount = 0;
                System.Xml.XmlDocument oXmlDoc = null;
                string sXmlString = null;
                //StrSql = "select OITM.ItemCode,listname ,isnull(Price,0)Price from ITM1 inner join OITM on OITM.ItemCode=ITM1.ItemCode inner join OPLN on OPLN.ListNum=ITM1.PriceList where OITM.U_IntegratedStatus='N'";
                StrSql = "select OITM.ItemCode,isnull(Price,0)Price,OPLN.ListName,CASE OPLN.ListName when 'Goverment Rate' then isnull(OITM.U_GTQty,0) " +
                " when 'Loose Rate' then isnull(OITM.U_LRQty,0) when 'ProductRate' then isnull(OITM.U_PRQty,0) when 'Retail Price' then isnull(OITM.U_RRQty,0) when 'Shop Keepers Rate' then isnull(OITM.U_SKRQty,0) " +
                " when 'Whole saleRate' then isnull(OITM.U_WSRQty,0) end Qty from ITM1 inner join OITM on OITM.ItemCode=ITM1.ItemCode " +
                " inner join OPLN on OPLN.ListNum=ITM1.PriceList where OITM.U_IntegratedStatus='I'";
                DataSet objDataSet = ConDb.DbDataFromSAP(StrSql);
                sXmlString = objDataSet.GetXml();
                // sXmlString = oRsInv.GetAsXML();
                oXmlDoc = new System.Xml.XmlDocument();
                oXmlDoc.LoadXml(sXmlString);
                oXmlDoc.Save((sPath + FileName));



                XmlDocument reader = new XmlDocument();
                XmlDocument readerlines = new XmlDocument();
                IFormatProvider ifp = new System.Globalization.CultureInfo("en-US", true);
                reader.Load(sPath + FileName);

                XmlNodeList list = reader.GetElementsByTagName(g.row1);

                foreach (XmlNode node in list)
                {
                    XmlElement Element = (XmlElement)node;
                    try
                    {

                        StrSql = @"delete from [NOR_ITM_PRICE] where ItemCode = '" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and  PrcLstCd = '" + Element.GetElementsByTagName("ListName")[0].InnerText + "'";
                        insertuser = ConDb.ScalarExecuteBranch(StrSql);


                        string StrSql2 = @"insert into [NOR_ITM_PRICE] (ItemCode,PrcLstCd,Price,Qty ) values('" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "','" + Element.GetElementsByTagName("ListName")[0].InnerText + "','" + Element.GetElementsByTagName("Price")[0].InnerText + "'," + Convert.ToDouble(Element.GetElementsByTagName("Qty")[0].InnerText) + ")";
                        insertuser = ConDb.ScalarExecuteBranch(StrSql2);

                    }
                    catch (Exception e)
                    {
                        StrSql = @"Update [NOR_ITM_PRICE] set Qty =" + Element.GetElementsByTagName("Qty")[0].InnerText + ",  Price='" + Element.GetElementsByTagName("Price")[0].InnerText + "' where ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "' and PrcLstCd= '" + Element.GetElementsByTagName("listname")[0].InnerText + "'";
                        insertuser = ConDb.ScalarExecuteBranch(StrSql);
                    }
                    if (insertuser == "Success")
                    {
                        _str_update = @"UPDATE OITM set U_IntegratedStatus='I' where OITM.ItemCode='" + Element.GetElementsByTagName("ItemCode")[0].InnerText + "'";
                        insertuser = ConDb.ScalarExecuteSAP(_str_update);

                    }
                }
            }
            catch { MessageBox.Show("Error in PRICELIST Import"); }

        }//PRICE LIST

        #region Old Customr

//        public void sCustomer()//CUSTOMER
//        {
//            try
//            {
//                string sPath = "";
//                string FileName = "ItemPriceList.xml";
//                string StrSql = "";
//                string insertuser = "";
//                General gen = new General();
//                if (!File.Exists(sPath + FileName))
//                { File.Create(sPath + FileName); }
//                int recCount = 0;
//                // SAPbobsCOM.BusinessPartners oBPMaster;
//                System.Xml.XmlDocument oXmlDoc = null;
//                string sXmlString = null;
//                StrSql = @"SELECT OCRD.CardCode,ISNULL(OCRD.CardName,'-')CardName,isnull(convert(varchar(100),CRD1.Building),'-')+''+ isnull(CRD1.Street,'-') [Address1],isnull(unit.U_unit,'')[U_unit],isnull(unit.U_CrediLmt,'')[U_CrediLmt],isnull(unit.U_status,'')[U_status],
//                        isnull(CRD1.Block,'-') [Address2],isnull(CRD1.City,'-')[Address3],isnull(CRD1.ZipCode,0)[ZipCode],isnull(OCRD.Phone1,0)[Phone1],
//                        isnull(OCRD.Cellular,0)[Cellular],OCRD.SlpCode,OCRD.CardType,CAST(OCRD.CreditLine AS INT)CreditLine,OCRD.Balance,isnull(OPLN.ListName,'')[ListName],
//                        ISNULL((SELECT top 1 CRD7.TaxId1 FROM CRD7 WHERE CRD7.CardCode=OCRD.CardCode AND AddrType='S'),0)[CST],
//                        ISNULL((SELECT top 1 CRD7.TaxId11 FROM CRD7 WHERE CRD7.CardCode=OCRD.CardCode AND AddrType='S'),0)[Tin],
//                        ISNULL((SELECT top 1 CRD7.TaxId0 FROM CRD7 WHERE CRD7.CardCode=OCRD.CardCode  AND AddrType='S'),0)[Pan],
//                         unit.U_unitcode,OCRD.GroupCode
//                        FROM OCRD  INNER JOIN  CRD1 ON OCRD.CardCode  = CRD1.CardCode 
//                        INNER JOIN OPLN ON OCRD.ListNum = OPLN.ListNum  left join [@NOR_UNITALLOC] unit on OCRD.CardCode=unit.U_cuscode  WHERE OCRD.U_IntegratedStatus='N' AND CRD1.AdresType ='B'";
//                //                StrSql = @"select OCRD.CardCode,OCRD.CardName, isnull(convert(varchar(100),CRD1.Building),'')+''+ isnull(CRD1.Street,'') [Address1],
//                //                            isnull(CRD1.Block,'') [Address2],isnull(CRD1.City,'')[Address3],isnull(CRD1.ZipCode,0)[ZipCode],isnull(OCRD.Phone1,0)[Phone1],
//                //                            isnull(OCRD.Cellular,0)[Cellular],OCRD.SlpCode,OCRD.CardType,CAST(OCRD.CreditLine AS INT)CreditLine,OCRD.Balance,isnull(OPLN.ListName,'')[ListName],
//                //                            isnull(CRD7.TaxId1,'')[CST],isnull(TaxId11,'')[Tin] from ocrd left join CRD1 on CRD1.CardCode=OCRD.CardCode
//                //                            left join CRD7 on CRD7.CardCode=OCRD.CardCode and CRD7.Address='Ship to' left join OPLN on OPLN.ListNum=OCRD.ListNum where 
//                //                            CRD1.AdresType='S' and OCRD.CardType='C' and U_IntegratedStatus='N'";

//                DataSet objDataSet = ConDb.DbDataFromBranch(StrSql);
//                sXmlString = objDataSet.GetXml();
//                oXmlDoc = new System.Xml.XmlDocument();
//                oXmlDoc.LoadXml(sXmlString);
//                oXmlDoc.Save((sPath + FileName));
//                SAPbobsCOM.BusinessPartners oBPMaster;

//                XmlDocument reader = new XmlDocument();
//                XmlDocument readerlines = new XmlDocument();
//                IFormatProvider ifp = new System.Globalization.CultureInfo("en-US", true);
//                reader.Load(sPath + FileName);

//                XmlNodeList list = reader.GetElementsByTagName(gen.row1);
//                int Error = 0;
//                foreach (XmlNode node in list)
//                {

//                    XmlElement Element = (XmlElement)node;
//                    string strCardCode1 = Element.GetElementsByTagName("CardCode")[0].InnerText.ToString();
//                    string strUnitCode = Element.GetElementsByTagName("U_unitcode")[0].InnerText.ToString();
//                    string strStatus = Element.GetElementsByTagName("U_status")[0].InnerText.ToString();
//                    //string UnitGetQrry = "SELECT * FROM [@NOR_UNITMASTER]";
//                    //SAPbobsCOM.Recordset oRsInv = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//                    //oRsInv.DoQuery(UnitGetQrry);
//                    //while (!oRsInv.EoF)
//                    //{
//                    if (strStatus == "Y")
//                    {
//                        //string UnitCode = oRsInv.Fields.Item("Code").Value.ToString();
//                        string QRY1 = "Select * from [@NOR_BRANCH_DTL] Where U_UnitId ='" + strUnitCode + "'";
//                        SAPbobsCOM.Recordset rsCompany = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//                        //SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
//                        rsCompany.DoQuery(QRY1);
//                        if (rsCompany.RecordCount > 0)
//                        {
//                            string server = rsCompany.Fields.Item("U_ServerName").Value.ToString();
//                            string DB = rsCompany.Fields.Item("U_CompanyDB").Value.ToString();
//                            string sUser = rsCompany.Fields.Item("U_SAPUserName").Value.ToString();
//                            string sPass = rsCompany.Fields.Item("U_SAPPassword").Value.ToString();
//                            string sqUser = rsCompany.Fields.Item("U_ServerUser").Value.ToString();
//                            string sqPass = rsCompany.Fields.Item("U_ServerPass").Value.ToString();
//                            gen.connectOtherCompany(server, DB, sUser, sPass, sqUser, sqPass);

//                            //oSales = (SAPbobsCOM.Documents)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
//                            oBPMaster = (SAPbobsCOM.BusinessPartners)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
//                            string strCardCode = Element.GetElementsByTagName("CardCode")[0].InnerText;
//                            string strCodeQrry = "Select * from OCRD Where CardCode ='" + strCardCode + "'";
//                            SAPbobsCOM.Recordset rsCustomer = ((SAPbobsCOM.Recordset)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
//                            rsCustomer.DoQuery(strCodeQrry);
//                            if (rsCustomer.RecordCount > 0)
//                            {
//                                oBPMaster.GetByKey(Element.GetElementsByTagName("CardCode")[0].InnerText);
//                                oBPMaster.CardName = Element.GetElementsByTagName("CardName")[0].InnerText;
//                                oBPMaster.Address = Element.GetElementsByTagName("Address1")[0].InnerText;

//                                //oBPMaster.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;
//                                string strAddress = Element.GetElementsByTagName("Address1")[0].InnerText;
//                                if (strAddress == "")
//                                    oBPMaster.Addresses.AddressName = Element.GetElementsByTagName("BPName")[0].InnerText;
//                                else
//                                    oBPMaster.Addresses.AddressName = Element.GetElementsByTagName("Address1")[0].InnerText;
//                                oBPMaster.Addresses.Block = Element.GetElementsByTagName("Address2")[0].InnerText;
//                                oBPMaster.Addresses.City = Element.GetElementsByTagName("Address3")[0].InnerText;
//                                oBPMaster.Addresses.ZipCode = Element.GetElementsByTagName("ZipCode")[0].InnerText;
//                                oBPMaster.FiscalTaxID.Address = Element.GetElementsByTagName("Address1")[0].InnerText;
//                                //oBPMaster.FiscalTaxID.TaxId0 = Element.GetElementsByTagName("CST")[0].InnerText;

//                                oBPMaster.ShipToBuildingFloorRoom = Element.GetElementsByTagName("Address1")[0].InnerText;
//                                oBPMaster.Block = Element.GetElementsByTagName("Address2")[0].InnerText;
//                                oBPMaster.City = Element.GetElementsByTagName("Address3")[0].InnerText;
//                                oBPMaster.ZipCode = Element.GetElementsByTagName("ZipCode")[0].InnerText;
//                                oBPMaster.Phone1 = Element.GetElementsByTagName("Phone1")[0].InnerText;
//                                oBPMaster.Cellular = Element.GetElementsByTagName("Cellular")[0].InnerText;
//                                // oBPMaster.EmailAddress = Element.GetElementsByTagName("E_mail")[0].InnerText;

//                                oBPMaster.CreditLimit = Convert.ToDouble(Element.GetElementsByTagName("U_CrediLmt")[0].InnerText);

//                                oBPMaster.SalesPersonCode = Convert.ToInt32(Element.GetElementsByTagName("SlpCode")[0].InnerText);
//                                oBPMaster.FiscalTaxID.TaxId0 =  Element.GetElementsByTagName("Pan")[0].InnerText;

//                                //oBPMaster.UserFields.Fields.Item("U_IntegratedStatus").Value = "I";
//                                if (Element.GetElementsByTagName("CardType")[0].InnerText.ToString() == "C")
//                                {
//                                    oBPMaster.CardType = SAPbobsCOM.BoCardTypes.cCustomer;
//                                }
//                                else if (Element.GetElementsByTagName("CardType")[0].InnerText.ToString() == "S")
//                                {
//                                    oBPMaster.CardType = SAPbobsCOM.BoCardTypes.cSupplier;
//                                }
//                                else if (Element.GetElementsByTagName("CardType")[0].InnerText.ToString() == "L")
//                                {
//                                    oBPMaster.CardType = SAPbobsCOM.BoCardTypes.cLid;
//                                }
//                                int iErrorCode = oBPMaster.Update();
//                                if (iErrorCode != 0)
//                                {
//                                    string sErrorMsg = Global.oCompny2.GetLastErrorDescription();
//                                    MessageBox.Show(sErrorMsg + "in '" + DB + "'", 1, "Ok", "", "");
//                                    Error = 1;

//                                }
//                                else
//                                {
//                                   // MessageBox.Show("Updated successfully to '" + DB + "'");//"","", 1,me"Ok", "", "");
//                                    // MessageBox.Show("Error in Export BP Master : " + sErrorMsg);
//                                    // StrSql = "Update [NOR_BP_MASTER] set IntegrationStatus='I' where BPcode='" + Element.GetElementsByTagName("BPCode")[0].InnerText + "'";
//                                    // ConDb.QueryNonExecuteBranch(StrSql);

//                                }

//                            }
//                            else
//                            {
//                                oBPMaster.CardCode = Element.GetElementsByTagName("CardCode")[0].InnerText;
//                                oBPMaster.CardName = Element.GetElementsByTagName("CardName")[0].InnerText;
//                                oBPMaster.Address = Element.GetElementsByTagName("Address1")[0].InnerText;

//                                oBPMaster.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;
//                                string strAddress = Element.GetElementsByTagName("Address1")[0].InnerText;
//                                if (strAddress == "")
//                                    oBPMaster.Addresses.AddressName = Element.GetElementsByTagName("CardName")[0].InnerText;
//                                else
//                                    oBPMaster.Addresses.AddressName = Element.GetElementsByTagName("Address1")[0].InnerText;
//                                oBPMaster.Addresses.Block = Element.GetElementsByTagName("Address2")[0].InnerText;
//                                oBPMaster.Addresses.City = Element.GetElementsByTagName("Address3")[0].InnerText;
//                                oBPMaster.Addresses.ZipCode = Element.GetElementsByTagName("ZipCode")[0].InnerText;
//                                oBPMaster.FiscalTaxID.Address = Element.GetElementsByTagName("Address1")[0].InnerText;
//                                //oBPMaster.FiscalTaxID.TaxId0 = Element.GetElementsByTagName("CST")[0].InnerText;

//                                oBPMaster.ShipToBuildingFloorRoom = Element.GetElementsByTagName("Address1")[0].InnerText;
//                                oBPMaster.Block = Element.GetElementsByTagName("Address2")[0].InnerText;
//                                oBPMaster.City = Element.GetElementsByTagName("Address3")[0].InnerText;
//                                oBPMaster.ZipCode = Element.GetElementsByTagName("ZipCode")[0].InnerText;
//                                oBPMaster.Phone1 = Element.GetElementsByTagName("Phone1")[0].InnerText;
//                                oBPMaster.Cellular = Element.GetElementsByTagName("Cellular")[0].InnerText;
//                                // oBPMaster.EmailAddress = Element.GetElementsByTagName("E_mail")[0].InnerText;

//                                //oBPMaster.CreditLimit = Convert.ToDouble(Element.GetElementsByTagName("CreditLine")[0].InnerText);

//                                //oBPMaster.SalesPersonCode = Convert.ToInt32(Element.GetElementsByTagName("SlpCode")[0].InnerText);
//                                //oBPMaster.WTCode =SAPbobsCOM.BoYesNoEnum.tNO.ToString();
//                                oBPMaster.FiscalTaxID.TaxId0 = Element.GetElementsByTagName("Pan")[0].InnerText;
//                                oBPMaster.FiscalTaxID.TaxId11 = Element.GetElementsByTagName("Tin")[0].InnerText;
//                                oBPMaster.FiscalTaxID.TaxId1 = Element.GetElementsByTagName("CST")[0].InnerText;
                                
//                                //oBPMaster.UserFields.Fields.Item("U_IntegratedStatus").Value = "I";
//                                if (Element.GetElementsByTagName("CardType")[0].InnerText.ToString() == "C")
//                                {
//                                    oBPMaster.CardType = SAPbobsCOM.BoCardTypes.cCustomer;
//                                }
//                                else if (Element.GetElementsByTagName("CardType")[0].InnerText.ToString() == "S")
//                                {
//                                    oBPMaster.CardType = SAPbobsCOM.BoCardTypes.cSupplier;
//                                }
//                                else if (Element.GetElementsByTagName("CardType")[0].InnerText.ToString() == "L")
//                                {
//                                    oBPMaster.CardType = SAPbobsCOM.BoCardTypes.cLid;
//                                }
//                              //  oBPMaster.CardType = SAPbobsCOM.BoCardTypes.cCustomer;
//                                oBPMaster.GroupCode = Convert.ToInt32(Element.GetElementsByTagName("GroupCode")[0].InnerText);
//                                int iErrorCode = oBPMaster.Add();
//                                if (iErrorCode != 0)
//                                {
//                                    string sErrorMsg = Global.oCompny2.GetLastErrorDescription();
//                                    MessageBox.Show(sErrorMsg + "in '" + strUnitCode + "'", 1, "Ok", "", "");
//                                    Error = 1;

//                                }
//                                else
//                                {
//                                    StrSql = "Update OCRD set U_IntegratedStatus='Y' where CardCode='" + Element.GetElementsByTagName("CardCode")[0].InnerText + "'";
//                                    ConDb.QueryNonExecuteBranch(StrSql);
//                                  //  MessageBox.Show("Updated successfully to '" + DB + "'");//, 1, "Ok", "", "");
//                                    // MessageBox.Show("Error in Export BP Master : " + sErrorMsg);
//                                    // StrSql = "Update [NOR_BP_MASTER] set IntegrationStatus='I' where BPcode='" + Element.GetElementsByTagName("BPCode")[0].InnerText + "'";
//                                    // ConDb.QueryNonExecuteBranch(StrSql);

//                                }

//                            }
//                        }
//                        //oRsInv.MoveNext();

//                    }
//                    if (Error == 0)
//                    {
//                        //StrSql = "Update OCRD set U_IntegratedStatus='I' where CardCode='" + Element.GetElementsByTagName("CardCode")[0].InnerText + "'";
//                        //ConDb.QueryNonExecuteBranch(StrSql);
//                    }
//                }



//            }

//            catch (Exception E) { MessageBox.Show(E.Message+"Customer"); }
//        }

        #endregion

        #region New Customer
       
        public void sCustomer()   //CUSTOMER
        {
            string sErrorMsg = "";
            try
            {
                string sPath = "";
                string FileName = "ItemPriceList.xml";
                string StrSql = "";
                string insertuser = "";

                General gen = new General();
                if (!File.Exists(sPath + FileName))
                { File.Create(sPath + FileName); }
                int recCount = 0;
                //SAPbobsCOM.BusinessPartners oBPMaster;
                System.Xml.XmlDocument oXmlDoc = null;
                string sXmlString = null;
                StrSql = @"SELECT OCRD.CardCode,ISNULL(OCRD.CardName,'-')CardName,ISNULL(OCRD.CardFName,'-')CardFName,ISNULL(CRD1.Building,'-')Building,ISNULL(CRD1.Street,'-')Street,isnull(unit.U_unitcode,'')[U_unit],
                           isnull(unit.U_CrediLmt,'')[U_CrediLmt],isnull(unit.U_status,'')[U_status],isnull(CRD1.ZipCode,0)ZipCode,
                           ISNULL(CRD1.Address,'-')Address,ISNULL(CRD1.Block,'-')Block, ISNULL(CRD1.State,'-')State,ISNULL(CRD1.Country,'-')Country,ISNULL(CRD1.County,'-')County,ISNULL(CRD1.StreetNo,'-')StreetNo,isnull(OCRD.Phone1,0)[Phone1],
                           isnull(OCRD.Phone2,0)[Phone2],isnull(OCRD.Fax,0)Fax,isnull(OCRD.E_Mail,'-')E_Mail,isnull(OCRD.IntrntSite,'-')IntrntSite,isnull(OCRD.City,0)City,
                           isnull(OCRD.Cellular,0)[Cellular],isnull(OCRD.SlpCode,0)SlpCode,OCRD.CardType,CAST(OCRD.CreditLine AS INT)CreditLine,isnull(OCRD.Balance,0)BpBalance,isnull(OCRD.ListNum,0)ListNum,isnull(OPLN.ListName,'')[ListName],
                           ISNULL((SELECT top 1 CRD7.TaxId1 FROM CRD7 WHERE CRD7.CardCode=OCRD.CardCode AND AddrType='S' and (TaxId1!='' or TaxId1 is not null)),0)[CST],
                           ISNULL((SELECT top 1 CRD7.TaxId11 FROM CRD7 WHERE CRD7.CardCode=OCRD.CardCode AND AddrType='S' and (TaxId11!='' or TaxId11 is not null)),0)[Tin],
                           ISNULL((SELECT top 1 CRD7.TaxId0 FROM CRD7 WHERE CRD7.CardCode=OCRD.CardCode  AND AddrType='S' and (TaxId0!='' or TaxId0 is not null)),0)[Pan],
                          unit.U_unitcode,OCRD.GroupCode
                                                                           
                          FROM OCRD  INNER JOIN  CRD1 ON OCRD.CardCode  = CRD1.CardCode 
                          INNER JOIN OPLN ON OCRD.ListNum = OPLN.ListNum  left join [@NOR_UNITALLOC] unit on OCRD.CardCode=unit.U_cuscode  WHERE OCRD.U_IntegratedStatus='N'  and unit.U_Status='Y'  and OCRD.CardType='C' ";
                //AND CRD1.AdresType ='B'
                DataSet objDataSet = ConDb.DbDataFromBranch(StrSql);
                sXmlString = objDataSet.GetXml();
                oXmlDoc = new System.Xml.XmlDocument();
                oXmlDoc.LoadXml(sXmlString);
                oXmlDoc.Save((sPath + FileName));
                SAPbobsCOM.BusinessPartners oBPMaster;

                XmlDocument reader = new XmlDocument();
                XmlDocument readerlines = new XmlDocument();
                IFormatProvider ifp = new System.Globalization.CultureInfo("en-US", true);
                reader.Load(sPath + FileName);

                XmlNodeList list = reader.GetElementsByTagName(gen.row1);
                int Error = 0;
                foreach (XmlNode node in list)
                {
                    XmlElement Element = (XmlElement)node;
                    string strCardCode1 = Element.GetElementsByTagName("CardCode")[0].InnerText.ToString();
                    string strUnitCode = Element.GetElementsByTagName("U_unitcode")[0].InnerText.ToString();
                    string strStatus = Element.GetElementsByTagName("U_status")[0].InnerText.ToString();
                   
                    if (strStatus == "Y")
                    {
                        string QRY1 = "Select * from [@NOR_BRANCH_DTL] Where U_UnitId ='" + strUnitCode + "'";
                        SAPbobsCOM.Recordset rsCompany = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        rsCompany.DoQuery(QRY1);
                        if (rsCompany.RecordCount > 0)
                        {
                            string server = rsCompany.Fields.Item("U_ServerName").Value.ToString();
                            string Licserver = rsCompany.Fields.Item("U_Licserver").Value.ToString();
                            string DB = rsCompany.Fields.Item("U_CompanyDB").Value.ToString();
                            string sUser = rsCompany.Fields.Item("U_SAPUserName").Value.ToString();
                            string sPass = rsCompany.Fields.Item("U_SAPPassword").Value.ToString();
                            string sqUser = rsCompany.Fields.Item("U_ServerUser").Value.ToString();
                            string sqPass = rsCompany.Fields.Item("U_ServerPass").Value.ToString();
                            gen.connectOtherCompany(server,Licserver, DB, sUser, sPass, sqUser, sqPass);

                            string strCardCode = Element.GetElementsByTagName("CardCode")[0].InnerText;
                            string strCodeQrry = "Select * from  [" + General.oCompany.CompanyDB + "].dbo.OCRD Where CardCode ='" + strCardCode + "'";
                            SAPbobsCOM.Recordset rsCustomer = ((SAPbobsCOM.Recordset)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                            rsCustomer.DoQuery(strCodeQrry);

                            Global.OBPA = (SAPbobsCOM.BusinessPartners)(General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners));

                            Global.OBPA.GetByKey(rsCustomer.Fields.Item("CardCode").Value.ToString());
                            Global.OBPB = (SAPbobsCOM.BusinessPartners)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners));
                           
                            int Errror = 0;
                            if (Global.OBPB.GetByKey(rsCustomer.Fields.Item("CardCode").Value.ToString()))
                            {
                                Global.OBPB.CardCode = Global.OBPA.CardCode;
                                Global.OBPB.CardName = Global.OBPA.CardName;
                                Global.OBPB.CardType = Global.OBPA.CardType;
                                Global.OBPB.CardForeignName = Global.OBPA.CardForeignName;
                                Global.OBPB.Phone1 = Global.OBPA.Phone1;
                                Global.OBPB.Phone2 = Global.OBPA.Phone2;
                                Global.OBPB.Fax = Global.OBPA.Fax;
                                Global.OBPB.EmailAddress = Global.OBPA.EmailAddress;
                                Global.OBPB.Website = Global.OBPA.Website;
                                Global.OBPB.Cellular = Global.OBPA.Cellular;
                                Global.OBPB.SalesPersonCode = Global.OBPA.SalesPersonCode;
                                Global.OBPB.CreditLimit = Global.OBPA.CreditLimit;

                                Global.OBPB.PriceListNum = Global.OBPA.PriceListNum;
                                Global.OBPB.FiscalTaxID.TaxId0 = Element.GetElementsByTagName("Pan")[0].InnerText.ToString();
                                Global.OBPB.FiscalTaxID.TaxId1 = Element.GetElementsByTagName("CST")[0].InnerText.ToString();
                                Global.OBPB.FiscalTaxID.TaxId11 = Element.GetElementsByTagName("Tin")[0].InnerText.ToString();
                                Global.OBPB.GroupCode = Global.OBPA.GroupCode;
                                Global.OBPB.PayTermsGrpCode = Global.OBPA.PayTermsGrpCode;
                                Global.OBPB.TaxExemptionLetterNum = Global.OBPA.TaxExemptionLetterNum;

                                Global.OBPB.Address = Global.OBPA.Address;
                                Global.OBPB.BillToBuildingFloorRoom = Global.OBPA.BillToBuildingFloorRoom;
                                Global.OBPB.BilltoDefault = Global.OBPA.BilltoDefault;
                                Global.OBPB.BillToState = Global.OBPA.BillToState;
                                Global.OBPB.ShipToBuildingFloorRoom = Global.OBPA.ShipToBuildingFloorRoom;
                                Global.OBPB.Block = Global.OBPA.Block;
                                Global.OBPB.City = Global.OBPA.City;

                                Global.OBPB.Country = Global.OBPA.Country;
                                Global.OBPB.County = Global.OBPA.County;
                                Global.OBPB.ZipCode = Global.OBPA.ZipCode;

                               // Global.OBPA.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                                Global.OBPA.Addresses.SetCurrentLine(0);
                                Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                                
                                Global.OBPB.Addresses.SetCurrentLine(0);

                                Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;

                                Global.OBPB.Addresses.Add();
                                Global.OBPA.Addresses.SetCurrentLine(1);
                                Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;

                                Global.OBPB.Addresses.SetCurrentLine(1);



                                Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;

                                Global.OBPB.Addresses.Add();

                                string VendorPpty64 = rsCustomer.Fields.Item("QryGroup64").Value.ToString();
                                if (VendorPpty64 == "Y")
                                {
                                    Global.OBPB.set_Properties(64, SAPbobsCOM.BoYesNoEnum.tYES);
                                }

                                Global.OBPB.PriceListNum = Global.OBPA.PriceListNum;
                                Error = Global.OBPB.Update();

                                if (Errror != 0)
                                {
                                    sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                    MessageBox.Show(sErrorMsg);
                                }
                            }
                            else
                            {
                                Global.OBPB.CardCode = Global.OBPA.CardCode;
                                Global.OBPB.CardName = Global.OBPA.CardName;
                                Global.OBPB.CardType = Global.OBPA.CardType;
                                Global.OBPB.CardForeignName = Global.OBPA.CardForeignName;
                                Global.OBPB.Phone1 = Global.OBPA.Phone1;
                                Global.OBPB.Phone2 = Global.OBPA.Phone2;
                                Global.OBPB.Fax = Global.OBPA.Fax;
                                Global.OBPB.EmailAddress = Global.OBPA.EmailAddress;
                                Global.OBPB.Website = Global.OBPA.Website;
                                Global.OBPB.Cellular = Global.OBPA.Cellular;
                                Global.OBPB.SalesPersonCode = Global.OBPA.SalesPersonCode;
                                Global.OBPB.CreditLimit = Global.OBPA.CreditLimit;
                                Global.OBPB.PriceListNum = Global.OBPA.PriceListNum;
                                Global.OBPB.FiscalTaxID.TaxId0 = Global.OBPA.FiscalTaxID.TaxId0;
                                Global.OBPB.FiscalTaxID.TaxId1 = Global.OBPA.FiscalTaxID.TaxId1;
                                Global.OBPB.FiscalTaxID.TaxId11 = Global.OBPA.FiscalTaxID.TaxId11;
                                Global.OBPB.GroupCode = Global.OBPA.GroupCode;
                                Global.OBPB.TaxExemptionLetterNum = Global.OBPA.TaxExemptionLetterNum;
                                Global.OBPB.PayTermsGrpCode = Global.OBPA.PayTermsGrpCode;
                                //  Global.OBPB.Add();


                                //Global.OBPB.Address = Global.OBPA.Address;
                                //Global.OBPB.BillToBuildingFloorRoom = Global.OBPA.BillToBuildingFloorRoom;
                                //Global.OBPB.BilltoDefault = Global.OBPA.BilltoDefault;
                                //Global.OBPB.BillToState = Global.OBPA.BillToState;
                                //Global.OBPB.ShipToBuildingFloorRoom = Global.OBPA.ShipToBuildingFloorRoom;
                                //Global.OBPB.Block = Global.OBPA.Block;
                                //Global.OBPB.City = Global.OBPA.City;
                                //Global.OBPB.ContactPerson = Global.OBPA.ContactPerson;
                                //Global.OBPB.Country = Global.OBPA.Country;
                                //Global.OBPB.County = Global.OBPA.County;
                                //Global.OBPB.ZipCode = Global.OBPA.ZipCode;



                                Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;

                                Global.OBPB.Addresses.SetCurrentLine(0);

                                //Global.OBPB.Addresses.SetCurrentLine(0);

                                Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;
                                Global.OBPB.Addresses.AddressName2 = Global.OBPA.Addresses.AddressName2;
                                Global.OBPB.Addresses.AddressName3 = Global.OBPA.Addresses.AddressName3;
                                Global.OBPB.Addresses.Add();
                                Global.OBPA.Addresses.SetCurrentLine(1);
                                Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;

                                Global.OBPB.Addresses.SetCurrentLine(1);



                                Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;

                                Global.OBPB.Addresses.Add();





                                string VendorPpty64 = rsCustomer.Fields.Item("QryGroup64").Value.ToString();
                                if (VendorPpty64 == "Y")
                                {
                                    Global.OBPB.set_Properties(64, SAPbobsCOM.BoYesNoEnum.tYES);
                                }
                                Errror = Global.OBPB.Add();


                                if (Errror != 0)
                                {
                                    sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                    MessageBox.Show(sErrorMsg);

                                }

                                if (Global.OBPB.GetByKey(rsCustomer.Fields.Item("CardCode").Value.ToString()))
                                {
                                    Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                                    //Global.OBPA.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;

                                    Global.OBPA.Addresses.SetCurrentLine(0);
                                    //Global.OBPB.Addresses.SetCurrentLine(0);
                                    Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                    Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                    Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                    Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                    Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                    Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                    Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                    Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                    Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                    Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;
                                    Global.OBPB.Addresses.AddressName2 = Global.OBPA.Addresses.AddressName2;
                                    Global.OBPB.Addresses.AddressName3 = Global.OBPA.Addresses.AddressName3;


                                    Errror = Global.OBPB.Update();
                                }



                            }
                            if (Errror != 0)
                            {
                                sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                MessageBox.Show(sErrorMsg);
                            }
                            else
                            {
                                StrSql = "Update OCRD set U_IntegratedStatus='Y' where CardCode='" + Element.GetElementsByTagName("CardCode")[0].InnerText + "'";
                                ConDb.QueryNonExecuteBranch(StrSql);
                            }
                        }
                        //oRsInv.MoveNext();

                    }
                    if (Error != 0)
                    {
                        sErrorMsg = Global.oCompny2.GetLastErrorDescription();    
                        //StrSql = "Update OCRD set U_IntegratedStatus='I' where CardCode='" + Element.GetElementsByTagName("CardCode")[0].InnerText + "'";
                        //ConDb.QueryNonExecuteBranch(StrSql);
                    }
                }
            }
            catch (Exception E)
            {
                MessageBox.Show(sErrorMsg);
             }
        }

        public void New_sCustomer()   //CUSTOMER
        {
            string sErrorMsg = "";
            try
            {
                string sPath = "";
                string FileName = "ItemPriceList.xml";
                string StrSql = "";
                string insertuser = "";
                bool _connection = false;
                string UnitGet = "";
                General gen = new General();
                if (!File.Exists(sPath + FileName))
                { File.Create(sPath + FileName); }
                int recCount = 0;
                //SAPbobsCOM.BusinessPartners oBPMaster;
                System.Xml.XmlDocument oXmlDoc = null;
                string sXmlString = null;
                int Errror = 0;
                int Error = 0;
                StrSql = @"SELECT distinct T2.*
                            FROM OCRD  T0 inner join [@NOR_UNITALLOC] T1 on T0.CardCode=T1.U_cuscode  
                            inner join [@NOR_BRANCH_DTL] T2 on T2.U_UnitId=T1.U_unitcode 
                              WHERE T0.U_IntegratedStatus='N'  and T1.U_Status='Y'  and T0.CardType='C'";
                SAPbobsCOM.Recordset rsCompany = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsCompany.DoQuery(StrSql);
                while (!rsCompany.EoF)
                {
                    string Licserver = rsCompany.Fields.Item("U_LicServer").Value.ToString();
                    string server = rsCompany.Fields.Item("U_ServerName").Value.ToString();
                    string DB = rsCompany.Fields.Item("U_CompanyDB").Value.ToString();
                    string sUser = rsCompany.Fields.Item("U_SAPUserName").Value.ToString();
                    string sPass = rsCompany.Fields.Item("U_SAPPassword").Value.ToString();
                    string sqUser = rsCompany.Fields.Item("U_ServerUser").Value.ToString();
                    string sqPass = rsCompany.Fields.Item("U_ServerPass").Value.ToString();
                    UnitGet = rsCompany.Fields.Item("U_UnitId").Value.ToString();
                    _connection = gen.connectOtherCompany(server,Licserver, DB, sUser, sPass, sqUser, sqPass);
                    if (_connection == true)
                    {
                        StrSql = @"SELECT OCRD.CardCode,ISNULL(OCRD.CardName,'-')CardName,ISNULL(OCRD.CardFName,'-')CardFName,ISNULL(CRD1.Building,'-')Building,ISNULL(CRD1.Street,'-')Street,isnull(unit.U_unitcode,'')[U_unit],
                           isnull(unit.U_CrediLmt,'')[U_CrediLmt],isnull(unit.U_status,'')[U_status],isnull(CRD1.ZipCode,0)ZipCode,
                           ISNULL(CRD1.Address,'-')Address,ISNULL(CRD1.Block,'-')Block, ISNULL(CRD1.State,'-')State,ISNULL(CRD1.Country,'-')Country,ISNULL(CRD1.County,'-')County,ISNULL(CRD1.StreetNo,'-')StreetNo,isnull(OCRD.Phone1,0)[Phone1],
                           isnull(OCRD.Phone2,0)[Phone2],isnull(OCRD.Fax,0)Fax,isnull(OCRD.E_Mail,'-')E_Mail,isnull(OCRD.IntrntSite,'-')IntrntSite,isnull(OCRD.City,0)City,
                           isnull(OCRD.Cellular,0)[Cellular],isnull(OCRD.SlpCode,0)SlpCode,OCRD.CardType,CAST(OCRD.CreditLine AS INT)CreditLine,isnull(OCRD.Balance,0)BpBalance,isnull(OCRD.ListNum,0)ListNum,isnull(OPLN.ListName,'')[ListName],
                           ISNULL((SELECT top 1 CRD7.TaxId1 FROM CRD7 WHERE CRD7.CardCode=OCRD.CardCode AND AddrType='S' and (TaxId1!='' or TaxId1 is not null)),0)[CST],
                           ISNULL((SELECT top 1 CRD7.TaxId11 FROM CRD7 WHERE CRD7.CardCode=OCRD.CardCode AND AddrType='S' and (TaxId11!='' or TaxId11 is not null)),0)[Tin],
                           ISNULL((SELECT top 1 CRD7.TaxId0 FROM CRD7 WHERE CRD7.CardCode=OCRD.CardCode  AND AddrType='S' and (TaxId0!='' or TaxId0 is not null)),0)[Pan],
                          unit.U_unitcode,OCRD.GroupCode  , OCRD.U_BPCatgry                                                                         
                          FROM OCRD  INNER JOIN  CRD1 ON OCRD.CardCode  = CRD1.CardCode 
                          INNER JOIN OPLN ON OCRD.ListNum = OPLN.ListNum 
                           left join [@NOR_UNITALLOC] unit on OCRD.CardCode=unit.U_cuscode  
                           WHERE OCRD.U_IntegratedStatus='N'  and unit.U_Status='Y'  and OCRD.CardType='C' and unit.U_UnitCode='" + UnitGet + "'";
                        //AND CRD1.AdresType ='B'
                        DataSet objDataSet = ConDb.DbDataFromBranch(StrSql);
                        sXmlString = objDataSet.GetXml();
                        oXmlDoc = new System.Xml.XmlDocument();
                        oXmlDoc.LoadXml(sXmlString);
                        oXmlDoc.Save((sPath + FileName));
                        //SAPbobsCOM.BusinessPartners oBPMaster;

                        XmlDocument reader = new XmlDocument();
                        XmlDocument readerlines = new XmlDocument();
                        IFormatProvider ifp = new System.Globalization.CultureInfo("en-US", true);
                        reader.Load(sPath + FileName);

                        XmlNodeList list = reader.GetElementsByTagName(gen.row1);
                        
                        foreach (XmlNode node in list)
                        {
                            XmlElement Element = (XmlElement)node;
                            string strCardCode1 = Element.GetElementsByTagName("CardCode")[0].InnerText.ToString();
                            string strUnitCode = Element.GetElementsByTagName("U_unitcode")[0].InnerText.ToString();
                            string strStatus = Element.GetElementsByTagName("U_status")[0].InnerText.ToString();

                            if (strStatus == "Y")
                            {
                                //string QRY1 = "Select * from [@NOR_BRANCH_DTL] Where U_UnitId ='" + strUnitCode + "'";
                                //SAPbobsCOM.Recordset rsCompany = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                //rsCompany.DoQuery(QRY1);
                                //if (rsCompany.RecordCount > 0)
                                //{
                                //    string server = rsCompany.Fields.Item("U_ServerName").Value.ToString();
                                //    string DB = rsCompany.Fields.Item("U_CompanyDB").Value.ToString();
                                //    string sUser = rsCompany.Fields.Item("U_SAPUserName").Value.ToString();
                                //    string sPass = rsCompany.Fields.Item("U_SAPPassword").Value.ToString();
                                //    string sqUser = rsCompany.Fields.Item("U_ServerUser").Value.ToString();
                                //    string sqPass = rsCompany.Fields.Item("U_ServerPass").Value.ToString();
                                //    gen.connectOtherCompany(server, DB, sUser, sPass, sqUser, sqPass);

                                    string strCardCode = Element.GetElementsByTagName("CardCode")[0].InnerText;
                                    string strCodeQrry = "Select * from  [" + General.oCompany.CompanyDB + "].dbo.OCRD Where CardCode ='" + strCardCode + "'";
                                    SAPbobsCOM.Recordset rsCustomer = ((SAPbobsCOM.Recordset)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                    rsCustomer.DoQuery(strCodeQrry);

                                    GC.Collect();
                                    Global.OBPA = (SAPbobsCOM.BusinessPartners)(General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners));

                                    Global.OBPA.GetByKey(rsCustomer.Fields.Item("CardCode").Value.ToString());
                                    Global.OBPB = (SAPbobsCOM.BusinessPartners)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners));

                                    if (Global.OBPB.GetByKey(rsCustomer.Fields.Item("CardCode").Value.ToString()))
                                    {
                                        //Global.OBPB.CardCode = Global.OBPA.CardCode;
                                        Global.OBPB.CardName = Global.OBPA.CardName;
                                        Global.OBPB.CardType = Global.OBPA.CardType;
                                        Global.OBPB.CardForeignName = Global.OBPA.CardForeignName;
                                        Global.OBPB.Phone1 = Global.OBPA.Phone1;
                                        Global.OBPB.Phone2 = Global.OBPA.Phone2;
                                        Global.OBPB.Fax = Global.OBPA.Fax;
                                        Global.OBPB.EmailAddress = Global.OBPA.EmailAddress;
                                        Global.OBPB.Website = Global.OBPA.Website;
                                        Global.OBPB.Cellular = Global.OBPA.Cellular;
                                        Global.OBPB.SalesPersonCode = Global.OBPA.SalesPersonCode;
                                        Global.OBPB.CreditLimit = Global.OBPA.CreditLimit;

                                        Global.OBPB.PriceListNum = Global.OBPA.PriceListNum;
                                        Global.OBPB.FiscalTaxID.TaxId0 = Global.OBPA.FiscalTaxID.TaxId0;
                                        Global.OBPB.FiscalTaxID.TaxId1 = Global.OBPA.FiscalTaxID.TaxId1;
                                        Global.OBPB.FiscalTaxID.TaxId11 = Global.OBPA.FiscalTaxID.TaxId11;
                                        //Global.OBPB.FiscalTaxID.TaxId0 = Element.GetElementsByTagName("Pan")[0].InnerText.ToString();
                                        //Global.OBPB.FiscalTaxID.TaxId1 = Element.GetElementsByTagName("CST")[0].InnerText.ToString();
                                        //Global.OBPB.FiscalTaxID.TaxId11 = Element.GetElementsByTagName("Tin")[0].InnerText.ToString();
                                        Global.OBPB.GroupCode = Global.OBPA.GroupCode;
                                        Global.OBPB.PayTermsGrpCode = Global.OBPA.PayTermsGrpCode;
                                        Global.OBPB.TaxExemptionLetterNum = Global.OBPA.TaxExemptionLetterNum;

                                        Global.OBPB.Address = Global.OBPA.Address;
                                        Global.OBPB.BillToBuildingFloorRoom = Global.OBPA.BillToBuildingFloorRoom;
                                        Global.OBPB.BilltoDefault = Global.OBPA.BilltoDefault;
                                        Global.OBPB.BillToState = Global.OBPA.BillToState;
                                        Global.OBPB.ShipToBuildingFloorRoom = Global.OBPA.ShipToBuildingFloorRoom;
                                        Global.OBPB.Block = Global.OBPA.Block;
                                        Global.OBPB.City = Global.OBPA.City;

                                        Global.OBPB.Country = Global.OBPA.Country;
                                        Global.OBPB.County = Global.OBPA.County;
                                        Global.OBPB.ZipCode = Global.OBPA.ZipCode;

                                        // Global.OBPA.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                                        Global.OBPB.Addresses.SetCurrentLine(0);
                                        Global.OBPA.Addresses.SetCurrentLine(0);
                                        Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                                        //Global.OBPB.Addresses.AddressType = Global.OBPA.Addresses.AddressType;
                                        Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                        Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                        Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                        Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                        Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                        Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                        Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                        Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                        Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                        Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;
                                        try
                                        {
                                            if (Global.OBPB.Addresses.AddressType == SAPbobsCOM.BoAddressType.bo_ShipTo)
                                            {
                                                if (Global.OBPA.Addresses.GSTIN != "") //added by Tamizh 10-Dec-2019
                                                {
                                                    Global.OBPB.Addresses.GstType = Global.OBPA.Addresses.GstType;
                                                    Global.OBPB.Addresses.GSTIN = Global.OBPA.Addresses.GSTIN;
                                                }
                                            }
                                            Global.OBPB.Addresses.GlobalLocationNumber = Global.OBPA.Addresses.GlobalLocationNumber;
                                        }
                                        catch { }
                                        Global.OBPB.Addresses.Add();
                                        
                                        Global.OBPA.Addresses.SetCurrentLine(1);
                                        Global.OBPB.Addresses.SetCurrentLine(1);
                                        Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;
                                       //Global.OBPB.Addresses.AddressType = Global.OBPA.Addresses.AddressType;
                                        
                                        Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                        Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                        Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                        Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                        Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                        Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                        Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                        Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                        Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                        Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;
                                        try
                                        {
                                            if (Global.OBPB.Addresses.AddressType == SAPbobsCOM.BoAddressType.bo_ShipTo)
                                            {
                                                if (Global.OBPA.Addresses.GSTIN != "") //added by Tamizh 10-Dec-2019
                                                {
                                                    Global.OBPB.Addresses.GstType = Global.OBPA.Addresses.GstType;
                                                    Global.OBPB.Addresses.GSTIN = Global.OBPA.Addresses.GSTIN;
                                                }
                                            }
                                            Global.OBPB.Addresses.GlobalLocationNumber = Global.OBPA.Addresses.GlobalLocationNumber;
                                        }
                                        catch { }

                                        Global.OBPB.Addresses.Add();

                                        string VendorPpty64 = rsCustomer.Fields.Item("QryGroup64").Value.ToString();
                                        if (VendorPpty64 == "Y")
                                        {
                                            Global.OBPB.set_Properties(64, SAPbobsCOM.BoYesNoEnum.tYES);
                                        }

                                        Global.OBPB.PriceListNum = Global.OBPA.PriceListNum;
                                        Errror = Global.OBPB.Update();
                                        //MessageBox.Show("Er"+Convert.ToString(Errror));
                                        //if (Errror != 0)
                                        //{
                                        //    sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                        //    MessageBox.Show(sErrorMsg);
                                        //}
                                    }
                                    else
                                    {
                                        Global.OBPB.CardCode = Global.OBPA.CardCode;
                                        Global.OBPB.CardName = Global.OBPA.CardName;
                                        Global.OBPB.CardType = Global.OBPA.CardType;
                                        Global.OBPB.CardForeignName = Global.OBPA.CardForeignName;
                                        Global.OBPB.Phone1 = Global.OBPA.Phone1;
                                        Global.OBPB.Phone2 = Global.OBPA.Phone2;
                                        Global.OBPB.Fax = Global.OBPA.Fax;
                                        Global.OBPB.EmailAddress = Global.OBPA.EmailAddress;
                                        Global.OBPB.Website = Global.OBPA.Website;
                                        Global.OBPB.Cellular = Global.OBPA.Cellular;
                                        Global.OBPB.SalesPersonCode = Global.OBPA.SalesPersonCode;
                                        Global.OBPB.CreditLimit = Global.OBPA.CreditLimit;
                                        Global.OBPB.PriceListNum = Global.OBPA.PriceListNum;
                                        Global.OBPB.FiscalTaxID.TaxId0 = Global.OBPA.FiscalTaxID.TaxId0;
                                        Global.OBPB.FiscalTaxID.TaxId1 = Global.OBPA.FiscalTaxID.TaxId1;
                                        Global.OBPB.FiscalTaxID.TaxId11 = Global.OBPA.FiscalTaxID.TaxId11;
                                        Global.OBPB.GroupCode = Global.OBPA.GroupCode;
                                        Global.OBPB.TaxExemptionLetterNum = Global.OBPA.TaxExemptionLetterNum;
                                        Global.OBPB.PayTermsGrpCode = Global.OBPA.PayTermsGrpCode;
                                        Global.OBPB.UserFields.Fields.Item("U_BPCatgry").Value = Global.OBPA.UserFields.Fields.Item("U_BPCatgry").Value;// added by Tamizh 19-Jul-2019
                                        //  Global.OBPB.Add();


                                        //Global.OBPB.Address = Global.OBPA.Address;
                                        //Global.OBPB.BillToBuildingFloorRoom = Global.OBPA.BillToBuildingFloorRoom;
                                        //Global.OBPB.BilltoDefault = Global.OBPA.BilltoDefault;
                                        //Global.OBPB.BillToState = Global.OBPA.BillToState;
                                        //Global.OBPB.ShipToBuildingFloorRoom = Global.OBPA.ShipToBuildingFloorRoom;
                                        //Global.OBPB.Block = Global.OBPA.Block;
                                        //Global.OBPB.City = Global.OBPA.City;
                                        //Global.OBPB.ContactPerson = Global.OBPA.ContactPerson;
                                        //Global.OBPB.Country = Global.OBPA.Country;
                                        //Global.OBPB.County = Global.OBPA.County;
                                        //Global.OBPB.ZipCode = Global.OBPA.ZipCode;



                                        //Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;

                                        Global.OBPB.Addresses.SetCurrentLine(0);
                                        Global.OBPA.Addresses.SetCurrentLine(0);
                                        //Global.OBPB.Addresses.SetCurrentLine(0);
                                        Global.OBPB.Addresses.AddressType = Global.OBPA.Addresses.AddressType;
                                        Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                        Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                        Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                        Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                        Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                        Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                        Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                        Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                        Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                        Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;
                                        Global.OBPB.Addresses.AddressName2 = Global.OBPA.Addresses.AddressName2;
                                        Global.OBPB.Addresses.AddressName3 = Global.OBPA.Addresses.AddressName3;
                                        try
                                        {
                                            if (Global.OBPB.Addresses.AddressType == SAPbobsCOM.BoAddressType.bo_ShipTo)
                                            {
                                                if (Global.OBPA.Addresses.GSTIN != "") //added by Tamizh 10-Dec-2019
                                                {
                                                    Global.OBPB.Addresses.GstType = Global.OBPA.Addresses.GstType;
                                                    Global.OBPB.Addresses.GSTIN = Global.OBPA.Addresses.GSTIN;
                                                }
                                            }
                                            Global.OBPB.Addresses.GlobalLocationNumber = Global.OBPA.Addresses.GlobalLocationNumber;
                                        }
                                        catch { }
                                        Global.OBPB.Addresses.Add();
                                        Global.OBPA.Addresses.SetCurrentLine(1);
                                        //Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;

                                        Global.OBPB.Addresses.SetCurrentLine(1);
                                        Global.OBPB.Addresses.AddressType = Global.OBPA.Addresses.AddressType;

                                      
                                        Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                        Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                        Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                        Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                        Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                        Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                        Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                        Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                        Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                        Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;
                                        try
                                        {
                                            if (Global.OBPB.Addresses.AddressType == SAPbobsCOM.BoAddressType.bo_ShipTo)
                                            {
                                                if (Global.OBPA.Addresses.GSTIN  != "")
                                                {
                                                Global.OBPB.Addresses.GstType = Global.OBPA.Addresses.GstType;
                                                Global.OBPB.Addresses.GSTIN = Global.OBPA.Addresses.GSTIN ;
                                                }
                                            }
                                            Global.OBPB.Addresses.GlobalLocationNumber = Global.OBPA.Addresses.GlobalLocationNumber;
                                        }
                                        catch { }
                                        Global.OBPB.Addresses.Add();
                                        

                                        string VendorPpty64 = rsCustomer.Fields.Item("QryGroup64").Value.ToString();
                                        if (VendorPpty64 == "Y")
                                        {
                                            Global.OBPB.set_Properties(64, SAPbobsCOM.BoYesNoEnum.tYES);
                                        }
                                        Errror = Global.OBPB.Add();


                                        if (Errror != 0)
                                        {
                                            sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                            MessageBox.Show(sErrorMsg);
                                            Error = 1;
                                         }

                                        Global.OBPB = (SAPbobsCOM.BusinessPartners)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners));
                                        if (Global.OBPB.GetByKey(rsCustomer.Fields.Item("CardCode").Value.ToString()))
                                        {
                                            Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                                            //Global.OBPA.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;

                                            Global.OBPA.Addresses.SetCurrentLine(0);
                                            //Global.OBPB.Addresses.SetCurrentLine(0);
                                            Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                            Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                            Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                            Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                            Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                            Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                            Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                            Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                            Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                            Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;
                                            Global.OBPB.Addresses.AddressName2 = Global.OBPA.Addresses.AddressName2;
                                            Global.OBPB.Addresses.AddressName3 = Global.OBPA.Addresses.AddressName3;
                                            try
                                            {
                                                if (Global.OBPB.Addresses.AddressType == SAPbobsCOM.BoAddressType.bo_BillTo)
                                                {
                                                    if (Global.OBPA.Addresses.GSTIN != "") //added by Tamizh 10-Dec-2019
                                                    {
                                                        Global.OBPB.Addresses.GstType = Global.OBPA.Addresses.GstType;
                                                        Global.OBPB.Addresses.GSTIN = Global.OBPA.Addresses.GSTIN;
                                                    }
                                                }
                                                Global.OBPB.Addresses.GlobalLocationNumber = Global.OBPA.Addresses.GlobalLocationNumber;
                                            }
                                            catch { }
                                            //Global.OBPB.Addresses.GstType = Global.OBPA.Addresses.GstType;
                                            //Global.OBPB.Addresses.GSTIN = Global.OBPA.Addresses.GSTIN;

                                            Errror = Global.OBPB.Update();
                                            //MessageBox.Show(Convert.ToString(Errror));
                                        }
                                    }
                                    if (Errror != 0)
                                    {
                                        sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                        MessageBox.Show(sErrorMsg);
                                        Error = 1;
                                    }
                                    else
                                    {
                                        StrSql = "Update  [@NOR_UNITALLOC] set U_Status='Y' where U_Cuscode='" + Element.GetElementsByTagName("CardCode")[0].InnerText + "' and U_unitcode='" + Element.GetElementsByTagName("U_unitcode")[0].InnerText.ToString() + "'";
                                        //StrSql = "Update OCRD set U_IntegratedStatus='Y' where CardCode='" + Element.GetElementsByTagName("CardCode")[0].InnerText + "'";
                                        ConDb.QueryNonExecuteBranch(StrSql);
                                    }
                               //oRsInv.MoveNext();

                            }
                            //if (Errror != 0)
                            //{
                            //    sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                            //    //StrSql = "Update OCRD set U_IntegratedStatus='I' where CardCode='" + Element.GetElementsByTagName("CardCode")[0].InnerText + "'";
                            //    //ConDb.QueryNonExecuteBranch(StrSql);
                            //}
                        }
                    }
                    rsCompany.MoveNext();
                }
                if (Error != 0)
                {
                   MessageBox.Show("All Customers are not completed");
                }
                else
                {
                    StrSql = "Update OCRD set U_IntegratedStatus='Y' where cardtype='C'";
                    ConDb.QueryNonExecuteBranch(StrSql);
                }
               
        }
            catch (Exception E)
            {
                MessageBox.Show(Convert.ToString(E.Message));
            }
        }


#region writelog

public void WriteSMSLog(string Str)
{
    FileStream fs;
    string chatlog = Application.StartupPath + @"\Log_" + DateTime.Today.ToString("yyyyMMdd") + ".txt";
    if (File.Exists(chatlog))
    {
    }
    else
    {
        fs = new FileStream(chatlog, FileMode.Create, FileAccess.Write);
        fs.Close();
    }
    // Dim objReader As New System.IO.StreamReader(chatlog)
    string sdate;
    sdate = DateTime.Now.ToShortTimeString().ToString();
    // objReader.Close()
    if (System.IO.File.Exists(chatlog) == true)
    {
        System.IO.StreamWriter objWriter = new System.IO.StreamWriter(chatlog, true);
        objWriter.WriteLine(sdate + " : " + Str);
        objWriter.Close();
    }
    else
    {
        System.IO.StreamWriter objWriter = new System.IO.StreamWriter(chatlog, false);
    }
}

#endregion
       


        #region New Vendoe
        public void sVendor()//CUSTOMER
        {
            string insertuser = "";
            string sErrorMsg = "";
            try
            {
                string sPath = "";
                string FileName = "ItemPriceList.xml";
                string StrSql = "";
         
                General gen = new General();
                if (!File.Exists(sPath + FileName))
                { File.Create(sPath + FileName); }
                int recCount = 0;
                // SAPbobsCOM.BusinessPartners oBPMaster;
                System.Xml.XmlDocument oXmlDoc = null;
                string sXmlString = null;
                StrSql = @"SELECT OCRD.CardCode,ISNULL(OCRD.CardName,'-')CardName,ISNULL(OCRD.CardFName,'-')CardFName,ISNULL(CRD1.Building,'-')Building,ISNULL(CRD1.Street,'-')Street,isnull(unit.U_unitcode,'')[U_unit],
                           isnull(unit.U_CrediLmt,'')[U_CrediLmt],isnull(unit.U_status,'')[U_status],isnull(CRD1.ZipCode,0)ZipCode,isnull(CRD1.AdresType,0)AdresType,
                           ISNULL(CRD1.Address,'-')Address,ISNULL(CRD1.Block,'-')Block, ISNULL(CRD1.State,'-')State,ISNULL(CRD1.Country,'-')Country,ISNULL(CRD1.County,'-')County,ISNULL(CRD1.StreetNo,'-')StreetNo,isnull(OCRD.Phone1,0)[Phone1],
                           isnull(OCRD.Phone2,0)[Phone2],isnull(OCRD.Fax,0)Fax,isnull(OCRD.E_Mail,'-')E_Mail,isnull(OCRD.IntrntSite,'-')IntrntSite,isnull(OCRD.City,0)City,
                           isnull(OCRD.Cellular,0)[Cellular],isnull(OCRD.SlpCode,0)SlpCode,OCRD.CardType,CAST(OCRD.CreditLine AS INT)CreditLine,isnull(OCRD.Balance,0)BpBalance,isnull(OCRD.ListNum,0)ListNum,isnull(OPLN.ListName,'')[ListName],
                           ISNULL((SELECT top 1 CRD7.TaxId1 FROM CRD7 WHERE CRD7.CardCode=OCRD.CardCode AND AddrType='S'),0)[CST],
                           ISNULL((SELECT top 1 CRD7.TaxId11 FROM CRD7 WHERE CRD7.CardCode=OCRD.CardCode AND AddrType='S'),0)[Tin],
                           ISNULL((SELECT top 1 CRD7.TaxId0 FROM CRD7 WHERE CRD7.CardCode=OCRD.CardCode  AND AddrType='S'),0)[Pan],
                          unit.U_unitcode,OCRD.GroupCode, OCRD.U_BPCatgry
                                                                           
                          FROM OCRD  INNER JOIN  CRD1 ON OCRD.CardCode  = CRD1.CardCode 
                          INNER JOIN OPLN ON OCRD.ListNum = OPLN.ListNum  left join [@NOR_UNITALLOC] unit on OCRD.CardCode=unit.U_cuscode  
                          WHERE OCRD.U_IntegratedStatus='N'  and unit.U_Status='Y' and OCRD.CardType='S' ";
                //                StrSql = @"select OCRD.CardCode,OCRD.CardName, isnull(convert(varchar(100),CRD1.Building),'')+''+ isnull(CRD1.Street,'') [Address1],
                //                            isnull(CRD1.Block,'') [Address2],isnull(CRD1.City,'')[Address3],isnull(CRD1.ZipCode,0)[ZipCode],isnull(OCRD.Phone1,0)[Phone1],
                //                            isnull(OCRD.Cellular,0)[Cellular],OCRD.SlpCode,OCRD.CardType,CAST(OCRD.CreditLine AS INT)CreditLine,OCRD.Balance,isnull(OPLN.ListName,'')[ListName],
                //                            isnull(CRD7.TaxId1,'')[CST],isnull(TaxId11,'')[Tin] from ocrd left join CRD1 on CRD1.CardCode=OCRD.CardCode
                //                            left join CRD7 on CRD7.CardCode=OCRD.CardCode and CRD7.Address='Ship to' left join OPLN on OPLN.ListNum=OCRD.ListNum where 
                //                            CRD1.AdresType='S' and OCRD.CardType='C' and U_IntegratedStatus='N'";
                //AND CRD1.AdresType ='B'
                DataSet objDataSet = ConDb.DbDataFromBranch(StrSql);
                WriteSMSLog(StrSql);
                sXmlString = objDataSet.GetXml();
                oXmlDoc = new System.Xml.XmlDocument();
                oXmlDoc.LoadXml(sXmlString);
                oXmlDoc.Save((sPath + FileName));
                //SAPbobsCOM.BusinessPartners oBPMaster;

                //MessageBox.Show(Convert.ToString(objDataSet.Tables[0].Rows.Count));

                XmlDocument reader = new XmlDocument();
                XmlDocument readerlines = new XmlDocument();
                IFormatProvider ifp = new System.Globalization.CultureInfo("en-US", true);
                reader.Load(sPath + FileName);

                XmlNodeList list = reader.GetElementsByTagName(gen.row1);
                foreach (XmlNode node in list)
                {
                    XmlElement Element = (XmlElement)node;
                    string strCardCode1 = Element.GetElementsByTagName("CardCode")[0].InnerText.ToString();
                    string strUnitCode = Element.GetElementsByTagName("U_unitcode")[0].InnerText.ToString();
                    string strStatus = Element.GetElementsByTagName("U_status")[0].InnerText.ToString();
                                           
                       //string UnitCode = oRsInv.Fields.Item("Code").Value.ToString();
                        string QRY1 = "Select * from [@NOR_BRANCH_DTL] Where U_UnitId ='" + strUnitCode + "'";
                        SAPbobsCOM.Recordset rsCompany = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                        rsCompany.DoQuery(QRY1);
                        WriteSMSLog(QRY1);
                        if (rsCompany.RecordCount > 0)
                        {
                            string server = rsCompany.Fields.Item("U_ServerName").Value.ToString();
                            string Licserver = rsCompany.Fields.Item("U_Licserver").Value.ToString();
                            string DB = rsCompany.Fields.Item("U_CompanyDB").Value.ToString();
                            string sUser = rsCompany.Fields.Item("U_SAPUserName").Value.ToString();
                            string sPass = rsCompany.Fields.Item("U_SAPPassword").Value.ToString();
                            string sqUser = rsCompany.Fields.Item("U_ServerUser").Value.ToString();
                            string sqPass = rsCompany.Fields.Item("U_ServerPass").Value.ToString();
                            gen.connectOtherCompany(server,Licserver, DB, sUser, sPass, sqUser, sqPass);
                            WriteSMSLog("Vendor:Server:" + server + "LicServer " + Licserver + "DB:" + DB + "SAPUser:" + sUser + "SAPPass" + sPass + "SQLUser" + sqUser + "SQLPass" + sqPass);
                            string strCardCode = Element.GetElementsByTagName("CardCode")[0].InnerText;
                            string strCodeQrry = "Select * from  [" + General.oCompany.CompanyDB + "].dbo.OCRD Where CardCode ='" + strCardCode + "'";
                            SAPbobsCOM.Recordset rsCustomer = ((SAPbobsCOM.Recordset)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                            rsCustomer.DoQuery(strCodeQrry);

                            GC.Collect();

                            Global.OBPA = (SAPbobsCOM.BusinessPartners)(General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners));



                            Global.OBPA.GetByKey(rsCustomer.Fields.Item("CardCode").Value.ToString());
                            Global.OBPB = (SAPbobsCOM.BusinessPartners)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners));
                            //Global.OBPB = (SAPbobsCOM.co)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.ocon));

                            int Errror = 0;
                            if (Global.OBPB.GetByKey(rsCustomer.Fields.Item("CardCode").Value.ToString()))
                            {
                               // Global.OBPB.GetByKey(rsCustomer.Fields.Item("CardCode").Value.ToString());
                                Global.OBPB.CardName = Global.OBPA.CardName;
                                Global.OBPB.CardType = Global.OBPA.CardType;
                                Global.OBPB.CardForeignName = Global.OBPA.CardForeignName;
                                Global.OBPB.Phone1 = Global.OBPA.Phone1;
                                Global.OBPB.Phone2 = Global.OBPA.Phone2;
                                Global.OBPB.Fax = Global.OBPA.Fax;
                                Global.OBPB.EmailAddress = Global.OBPA.EmailAddress;
                                Global.OBPB.Website = Global.OBPA.Website;
                                Global.OBPB.Cellular = Global.OBPA.Cellular;
                                Global.OBPB.SalesPersonCode = Global.OBPA.SalesPersonCode;
                                Global.OBPB.CreditLimit = Global.OBPA.CreditLimit;
                                Global.OBPB.UserFields.Fields.Item("U_BPCatgry").Value = Global.OBPA.UserFields.Fields.Item("U_BPCatgry").Value; // added by Tamizh 19-Jul-2019

                                Global.OBPB.PriceListNum = Global.OBPA.PriceListNum;
                                Global.OBPB.FiscalTaxID.TaxId0 = Global.OBPA.FiscalTaxID.TaxId0;
                                Global.OBPB.FiscalTaxID.TaxId1 = Global.OBPA.FiscalTaxID.TaxId1;
                                Global.OBPB.FiscalTaxID.TaxId11 = Global.OBPA.FiscalTaxID.TaxId11;
                                Global.OBPB.GroupCode = Global.OBPA.GroupCode;
                                Global.OBPB.PayTermsGrpCode = Global.OBPA.PayTermsGrpCode;
                                Global.OBPB.TaxExemptionLetterNum = Global.OBPA.TaxExemptionLetterNum;

                                Global.OBPB.Address = Global.OBPA.Address;
                                Global.OBPB.BillToBuildingFloorRoom = Global.OBPA.BillToBuildingFloorRoom;
                                Global.OBPB.BilltoDefault = Global.OBPA.BilltoDefault;
                                Global.OBPB.BillToState = Global.OBPA.BillToState;
                                Global.OBPB.ShipToBuildingFloorRoom = Global.OBPA.ShipToBuildingFloorRoom;
                                Global.OBPB.Block = Global.OBPA.Block;
                                Global.OBPB.City = Global.OBPA.City;

                                // Global.OBPB.ContactPerson = Global.OBPA.ContactPerson;
                                //  Global.OBPB.ContactPerson = Global.OBPA.ContactPerson;
                                Global.OBPB.Country = Global.OBPA.Country;
                                Global.OBPB.County = Global.OBPA.County;
                                Global.OBPB.ZipCode = Global.OBPA.ZipCode;
                                // Global.OBPA.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                                //Global.OBPB.GetByKey(rsCustomer.Fields.Item("CardCode").Value.ToString());
                                Global.OBPB.Addresses.SetCurrentLine(0);
                                Global.OBPA.Addresses.SetCurrentLine(0);
                                Global.OBPB.Addresses.AddressType =Global.OBPA.Addresses.AddressType;                                
                                //Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                                //Global.OBPB.Addresses.SetCurrentLine(0);
                                Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;
                                Global.OBPB.Addresses.AddressName2 = Global.OBPA.Addresses.AddressName2;
                                Global.OBPB.Addresses.AddressName3 = Global.OBPA.Addresses.AddressName3;

                                //MessageBox.Show("2");
                                try
                                {
                                    if (Global.OBPB.Addresses.AddressType == SAPbobsCOM.BoAddressType.bo_ShipTo)
                                    {
                                        if (Global.OBPA.Addresses.GSTIN != "") //added by Tamizh 10-Dec-2019
                                        {
                                            Global.OBPB.Addresses.GstType = Global.OBPA.Addresses.GstType;
                                            Global.OBPB.Addresses.GSTIN = Global.OBPA.Addresses.GSTIN;
                                        }
                                    }
                                    Global.OBPB.Addresses.GlobalLocationNumber = Global.OBPA.Addresses.GlobalLocationNumber;
                                    Global.OBPB.Addresses.Add();
                                }
                                catch { }
                                Global.OBPA.Addresses.SetCurrentLine(1);
                                Global.OBPB.Addresses.SetCurrentLine(1);
                                //Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;
                                Global.OBPB.Addresses.AddressType = Global.OBPA.Addresses.AddressType;

                               //// Global.OBPA.Addresses.SetCurrentLine(1);
                               // Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                               // Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                               // Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                               // Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                               // Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                               // Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                               // Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                               // Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                               // Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                               // Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;

                               // 
                                

                                Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;

                                //MessageBox.Show("3");
                                try
                                {
                                    if (Global.OBPB.Addresses.AddressType == SAPbobsCOM.BoAddressType.bo_ShipTo)
                                    {
                                        if (Global.OBPA.Addresses.GSTIN != "") //added by Tamizh 10-Dec-2019
                                        {
                                            Global.OBPB.Addresses.GstType = Global.OBPA.Addresses.GstType;
                                            Global.OBPB.Addresses.GSTIN = Global.OBPA.Addresses.GSTIN;
                                        }
                                    }
                                    Global.OBPB.Addresses.GlobalLocationNumber = Global.OBPA.Addresses.GlobalLocationNumber;
                                    Global.OBPB.Addresses.Add();
                                }
                                catch { }

                                string VendorPpty64 = rsCustomer.Fields.Item("QryGroup64").Value.ToString();
                                if (VendorPpty64 == "Y")
                                {
                                    Global.OBPB.set_Properties(64, SAPbobsCOM.BoYesNoEnum.tYES);
                                }
                                //Errror = Global.OBPB.Update();
                                Global.OBPB.PriceListNum = Global.OBPA.PriceListNum;
                                //Error = Global.OBPB.Update();
                                Errror = Global.OBPB.Update();
                                //if (Errror != 0)
                                //{
                                //    sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                //    //MessageBox.Show(sErrorMsg);
                                //    MessageBox.Show(sErrorMsg);
                                // }
                                //MessageBox.Show(Convert.ToString(Errror));
                             }
                            else
                            {
                                Global.OBPB.CardCode = Global.OBPA.CardCode;
                                Global.OBPB.CardName = Global.OBPA.CardName;
                                Global.OBPB.CardType = Global.OBPA.CardType;
                                Global.OBPB.CardForeignName = Global.OBPA.CardForeignName;
                                Global.OBPB.Phone1 = Global.OBPA.Phone1;
                                Global.OBPB.Phone2 = Global.OBPA.Phone2;
                                Global.OBPB.Fax = Global.OBPA.Fax;
                                Global.OBPB.EmailAddress = Global.OBPA.EmailAddress;
                                Global.OBPB.Website = Global.OBPA.Website;
                                Global.OBPB.Cellular = Global.OBPA.Cellular;
                                Global.OBPB.SalesPersonCode = Global.OBPA.SalesPersonCode;
                                Global.OBPB.CreditLimit = Global.OBPA.CreditLimit;
                                Global.OBPB.PriceListNum = Global.OBPA.PriceListNum;
                                Global.OBPB.FiscalTaxID.TaxId0 = Global.OBPA.FiscalTaxID.TaxId0;
                                Global.OBPB.FiscalTaxID.TaxId1 = Global.OBPA.FiscalTaxID.TaxId1;
                                Global.OBPB.FiscalTaxID.TaxId11 = Global.OBPA.FiscalTaxID.TaxId11;
                                Global.OBPB.GroupCode = Global.OBPA.GroupCode;
                                Global.OBPB.TaxExemptionLetterNum = Global.OBPA.TaxExemptionLetterNum;
                                Global.OBPB.PayTermsGrpCode = Global.OBPA.PayTermsGrpCode;
                                //  Global.OBPB.Add();


                                //Global.OBPB.Address = Global.OBPA.Address;
                                //Global.OBPB.BillToBuildingFloorRoom = Global.OBPA.BillToBuildingFloorRoom;
                                //Global.OBPB.BilltoDefault = Global.OBPA.BilltoDefault;
                                //Global.OBPB.BillToState = Global.OBPA.BillToState;
                                //Global.OBPB.ShipToBuildingFloorRoom = Global.OBPA.ShipToBuildingFloorRoom;
                                //Global.OBPB.Block = Global.OBPA.Block;
                                //Global.OBPB.City = Global.OBPA.City;
                                //Global.OBPB.ContactPerson = Global.OBPA.ContactPerson;
                                //Global.OBPB.Country = Global.OBPA.Country;
                                //Global.OBPB.County = Global.OBPA.County;
                                //Global.OBPB.ZipCode = Global.OBPA.ZipCode;

                                Global.OBPA.Addresses.SetCurrentLine(0);
                                Global.OBPB.Addresses.SetCurrentLine(0);
                                //Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                                Global.OBPB.Addresses.AddressType = Global.OBPA.Addresses.AddressType;
                            
                                //Global.OBPB.Addresses.SetCurrentLine(0);

                                Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;
                                Global.OBPB.Addresses.AddressName2 = Global.OBPA.Addresses.AddressName2;
                                Global.OBPB.Addresses.AddressName3 = Global.OBPA.Addresses.AddressName3;

                                //MessageBox.Show("4");
                                try
                                {
                                    if (Global.OBPB.Addresses.AddressType == SAPbobsCOM.BoAddressType.bo_ShipTo)
                                    {
                                        if (Global.OBPA.Addresses.GSTIN != "") //added by Tamizh 10-Dec-2019
                                        {
                                            Global.OBPB.Addresses.GstType = Global.OBPA.Addresses.GstType;
                                            Global.OBPB.Addresses.GSTIN = Global.OBPA.Addresses.GSTIN;
                                        }
                                    }
                                    Global.OBPB.Addresses.GlobalLocationNumber = Global.OBPA.Addresses.GlobalLocationNumber;
                                }
                                catch { }
                                Global.OBPB.Addresses.Add();
                            
                                Global.OBPA.Addresses.SetCurrentLine(1);
                                Global.OBPB.Addresses.SetCurrentLine(1);
                                //Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;
                                Global.OBPB.Addresses.AddressType = Global.OBPA.Addresses.AddressType;
                                

                                Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;

                                //MessageBox.Show("5");
                                try
                                {
                                    if (Global.OBPB.Addresses.AddressType == SAPbobsCOM.BoAddressType.bo_ShipTo)
                                    {
                                        if (Global.OBPA.Addresses.GSTIN != "") //added by Tamizh 10-Dec-2019
                                        {
                                            Global.OBPB.Addresses.GstType = Global.OBPA.Addresses.GstType;
                                            Global.OBPB.Addresses.GSTIN = Global.OBPA.Addresses.GSTIN;
                                        }
                                    }
                                    Global.OBPB.Addresses.GlobalLocationNumber = Global.OBPA.Addresses.GlobalLocationNumber;
                                }
                                catch (Exception e)
                                {
                                    MessageBox.Show(Convert.ToString(e.Message));
                                }
                                Global.OBPB.Addresses.Add();
                                
                                string VendorPpty64 = rsCustomer.Fields.Item("QryGroup64").Value.ToString();
                                if (VendorPpty64 == "Y")
                                {
                                    Global.OBPB.set_Properties(64, SAPbobsCOM.BoYesNoEnum.tYES);
                                }
                                Errror = Global.OBPB.Add();

                                //if (Errror != 0)
                                //{
                                //    sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                //    MessageBox.Show(sErrorMsg);
                                // }

                                ////Global.OBPB = (SAPbobsCOM.BusinessPartners)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners));
                                //if (Global.OBPB.GetByKey(rsCustomer.Fields.Item("CardCode").Value.ToString()))
                                //{
                                //    Global.OBPB.Addresses.AddressType = Global.OBPA.Addresses.AddressType;
                                //    //Global.OBPB.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                                //    //Global.OBPA.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;

                                //    Global.OBPA.Addresses.SetCurrentLine(0);
                                //    //Global.OBPB.Addresses.SetCurrentLine(0);
                                //    Global.OBPB.Addresses.AddressName = Global.OBPA.Addresses.AddressName;
                                //    Global.OBPB.Addresses.Block = Global.OBPA.Addresses.Block;
                                //    Global.OBPB.Addresses.BuildingFloorRoom = Global.OBPA.Addresses.BuildingFloorRoom;
                                //    Global.OBPB.Addresses.City = Global.OBPA.Addresses.City;
                                //    Global.OBPB.Addresses.Country = Global.OBPA.Addresses.Country;
                                //    Global.OBPB.Addresses.County = Global.OBPA.Addresses.County;
                                //    Global.OBPB.Addresses.State = Global.OBPA.Addresses.State;
                                //    Global.OBPB.Addresses.Street = Global.OBPA.Addresses.Street;
                                //    Global.OBPB.Addresses.StreetNo = Global.OBPA.Addresses.StreetNo;
                                //    Global.OBPB.Addresses.ZipCode = Global.OBPA.Addresses.ZipCode;
                                //    Global.OBPB.Addresses.AddressName2 = Global.OBPA.Addresses.AddressName2;
                                //    Global.OBPB.Addresses.AddressName3 = Global.OBPA.Addresses.AddressName3;
                                //    //MessageBox.Show("6");
                                //    try
                                //    {
                                //        if (Global.OBPB.Addresses.AddressType == SAPbobsCOM.BoAddressType.bo_ShipTo)
                                //        {
                                //            Global.OBPB.Addresses.GstType = Global.OBPA.Addresses.GstType;
                                //            Global.OBPB.Addresses.GSTIN = Global.OBPA.Addresses.GSTIN;
                                //        }
                                //        Global.OBPB.Addresses.GlobalLocationNumber = Global.OBPA.Addresses.GlobalLocationNumber;
                                //    }
                                //    catch { }
                                //    Errror = Global.OBPB.Update();
                                //    MessageBox.Show(Convert.ToString(Errror));
                                    
                                //} //If loop to update customer addresss


                            }//If loop end (Add/Update) 
                            if (Errror != 0)
                            {
                                //MessageBox.Show("Error");
                                sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                MessageBox.Show(sErrorMsg);
                            }
                            else
                            {
                                StrSql = "Update OCRD set U_IntegratedStatus='Y' where CardCode='" + Element.GetElementsByTagName("CardCode")[0].InnerText + "'";
                                ConDb.QueryNonExecuteBranch(StrSql);
                            }
                            }//company checking query
                    
                  }//For Loop End
                MessageBox.Show("Vendor Data's Imported Sucessfully");
         
            }//Try End
            catch (Exception E) { MessageBox.Show("Error:"+Convert.ToString(E.Message)); }
        
        }//Function End

        #endregion

        public void SALESEMP()
        {
            try
            {
                string sPath = "";
                string FileName = "SALESEMP.xml";
                string StrSql = "";
                SAPbobsCOM.SalesPersons oSalesPerson;
                int Error = 0;
                General gen = new General();
                if (!File.Exists(sPath + FileName))
                { File.Create(sPath + FileName); }
                int recCount = 0;
                System.Xml.XmlDocument oXmlDoc = null;
                string sXmlString = null;
                StrSql = @"SELECT SlpCode,SlpName,isnull(Memo,'')Memo,Commission,GroupCode,Locked,DataSource,UserSign,isnull(EmpID,'')EmpID FROM OSLP";
                DataSet objDataSet = ConDb.DbDataFromSAP(StrSql);
                sXmlString = objDataSet.GetXml();
                // sXmlString = oRsInv.GetAsXML();
                oXmlDoc = new System.Xml.XmlDocument();
                oXmlDoc.LoadXml(sXmlString);
                oXmlDoc.Save((sPath + FileName));



                XmlDocument reader = new XmlDocument();
                XmlDocument readerlines = new XmlDocument();
                IFormatProvider ifp = new System.Globalization.CultureInfo("en-US", true);
                reader.Load(sPath + FileName);

                XmlNodeList list = reader.GetElementsByTagName(gen.row1);


                foreach (XmlNode node in list)
                {
                    XmlElement Element = (XmlElement)node;
                    string UnitGetQrry = "SELECT * FROM [@NOR_UNITMASTER]";
                    SAPbobsCOM.Recordset oRsInv = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRsInv.DoQuery(UnitGetQrry);
                    while (!oRsInv.EoF)
                    {
                        string UnitCode = oRsInv.Fields.Item("Code").Value.ToString();
                        string QRY1 = "Select * from [@NOR_BRANCH_DTL] Where U_UnitId ='" + UnitCode + "'";
                        SAPbobsCOM.Recordset rsCompany = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                        rsCompany.DoQuery(QRY1);
                        if (rsCompany.RecordCount > 0)
                        {
                            string server = rsCompany.Fields.Item("U_ServerName").Value.ToString();
                            string Licserver = rsCompany.Fields.Item("U_Licserver").Value.ToString();
                            string DB = rsCompany.Fields.Item("U_CompanyDB").Value.ToString();
                            string sUser = rsCompany.Fields.Item("U_SAPUserName").Value.ToString();
                            string sPass = rsCompany.Fields.Item("U_SAPPassword").Value.ToString(); ;
                            string sqUser = rsCompany.Fields.Item("U_ServerUser").Value.ToString();
                            string sqPass = rsCompany.Fields.Item("U_ServerPass").Value.ToString();
                            gen.connectOtherCompany(server,Licserver, DB, sUser, sPass, sqUser, sqPass);

                            //oSales = (SAPbobsCOM.Documents)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                            oSalesPerson = (SAPbobsCOM.SalesPersons)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesPersons);
                            string strEmployeeName = Element.GetElementsByTagName("SlpName")[0].InnerText;
                            if (strEmployeeName != "-No Sales Employee-")
                            {
                                string strCodeQrry = "Select SlpCode from OSLP Where SlpName ='" + strEmployeeName + "'";
                                SAPbobsCOM.Recordset rsEmployee = ((SAPbobsCOM.Recordset)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                                rsEmployee.DoQuery(strCodeQrry);
                                if (rsEmployee.RecordCount > 0)
                                {

                                    oSalesPerson.GetByKey(Convert.ToInt32(rsEmployee.Fields.Item("SlpCode").Value.ToString()));
                                    // oSalesPerson.SalesEmployeeCode =Convert.ToInt32(Element.GetElementsByTagName("SlpCode")[0].InnerText);
                                    oSalesPerson.SalesEmployeeName = Element.GetElementsByTagName("SlpName")[0].InnerText;
                                    oSalesPerson.CommissionForSalesEmployee = Convert.ToDouble(Element.GetElementsByTagName("Commission")[0].InnerText);
                                    oSalesPerson.CommissionGroup = Convert.ToInt32(Element.GetElementsByTagName("GroupCode")[0].InnerText);
                                    //  oSalesPerson.Locked = Element.GetElementsByTagName("Locked")[0].InnerText;
                                    int iErrorCode = oSalesPerson.Update();
                                    if (iErrorCode != 0)
                                    {
                                        string sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                        MessageBox.Show(sErrorMsg + "in '" + DB + "'");
                                        Error = 1;

                                    }
                                    else
                                    {
                                        MessageBox.Show("Updated successfully to '" + DB + "'");
                                        //General.SapApplication.MessageBox("Updated successfully to '" + DB + "'", 2, "Ok", "", "");
                                        // MessageBox.Show("Error in Export BP Master : " + sErrorMsg);
                                        // StrSql = "Update [NOR_BP_MASTER] set IntegrationStatus='I' where BPcode='" + Element.GetElementsByTagName("BPCode")[0].InnerText + "'";
                                        // ConDb.QueryNonExecuteBranch(StrSql);

                                    
                                    }
                                }
                                else
                                {
                                    oSalesPerson.SalesEmployeeName = Element.GetElementsByTagName("SlpName")[0].InnerText;
                                    oSalesPerson.CommissionForSalesEmployee = Convert.ToDouble(Element.GetElementsByTagName("Commission")[0].InnerText);
                                    oSalesPerson.CommissionGroup = Convert.ToInt32(Element.GetElementsByTagName("GroupCode")[0].InnerText);
                                    // oSalesPerson.Locked = Element.GetElementsByTagName("Locked")[0].InnerText;
                                    //oSalesPerson.EmployeeID = Convert.ToInt32(Element.GetElementsByTagName("EmpID")[0].InnerText);
                                    int iErrorCode = oSalesPerson.Add();
                                    if (iErrorCode != 0)
                                    {
                                        string sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                        MessageBox.Show(sErrorMsg + "in '" + DB + "'");
                                        Error = 1;

                                    }
                                    else
                                    {
                                        //General.SapApplication.MessageBox("Updated successfully to '" + DB + "'", 1, "Ok", "", "");
                                        MessageBox.Show("Updated successfully to '" + DB + "'");
                                        // StrSql = "Update [NOR_BP_MASTER] set IntegrationStatus='I' where BPcode='" + Element.GetElementsByTagName("BPCode")[0].InnerText + "'";
                                        // ConDb.QueryNonExecuteBranch(StrSql);

                                    
                                    }
                                }
                            }
                        }
                        oRsInv.MoveNext();
                    }
                }
            }
            catch (Exception E) { MessageBox.Show(E.Message+"SalesEmployee"); }

        }

        private void btnExport_Click(object sender, EventArgs e)
        {

        }//SALES EMPLOYEE
//------------------------------------WareHouse
        public void WareHouse()
        {
            try
            {
                string sPath = "";
                string FileName = "Warehouse.xml";
                string StrSql = "";
                SAPbobsCOM.Warehouses oWareHouse;
                int Error = 0;
                General gen = new General();
                if (!File.Exists(sPath + FileName))
                { File.Create(sPath + FileName); }
                int recCount = 0;
                System.Xml.XmlDocument oXmlDoc = null;
                string sXmlString = null;
                StrSql = @"SELECT WhsCode,WhsName,Location,isnull(WhShipTo,'') WhShipTo,isnull(Street,'') Street,isNull(StreetNo,'') StreetNo,isnull(Block,'') Block,isnull(Building,'') Building,isnull(ZipCode,'') ZipCode,isnull(City,'') City,isnull(County,'') County,isnull(Country,'') Country,
isnull(State,'') State,Nettable,DropShip,Excisable,isnull(U_Unit,'') U_Unit,isnull(U_WhsType,'') U_WhsType FROM OWHS where U_IntegratedStatus ='N'";
                DataSet objDataSet = ConDb.DbDataFromSAP(StrSql);
                sXmlString = objDataSet.GetXml();
                // sXmlString = oRsInv.GetAsXML();
                oXmlDoc = new System.Xml.XmlDocument();
                oXmlDoc.LoadXml(sXmlString);
                oXmlDoc.Save((sPath + FileName));



                XmlDocument reader = new XmlDocument();
                XmlDocument readerlines = new XmlDocument();
                IFormatProvider ifp = new System.Globalization.CultureInfo("en-US", true);
                reader.Load(sPath + FileName);

                XmlNodeList list = reader.GetElementsByTagName(gen.row1);


                foreach (XmlNode node in list)
                {
                    XmlElement Element = (XmlElement)node;
                    //string UnitGetQrry = "SELECT * FROM [@NOR_UNITMASTER]";
                    //SAPbobsCOM.Recordset oRsInv = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //oRsInv.DoQuery(UnitGetQrry);
                    string strUnit = Element.GetElementsByTagName("U_Unit")[0].InnerText;
                    //while (!oRsInv.EoF)
                    //{
                    //string UnitCode = oRsInv.Fields.Item("Code").Value.ToString();
                    string QRY1 = "Select * from [@NOR_BRANCH_DTL] Where U_UnitId ='" + strUnit + "'";
                    SAPbobsCOM.Recordset rsCompany = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                    rsCompany.DoQuery(QRY1);
                    if (rsCompany.RecordCount > 0)
                    {
                        string server = rsCompany.Fields.Item("U_ServerName").Value.ToString();
                        string Licserver = rsCompany.Fields.Item("U_Licserver").Value.ToString();
                        string DB = rsCompany.Fields.Item("U_CompanyDB").Value.ToString();
                        string sUser = rsCompany.Fields.Item("U_SAPUserName").Value.ToString();
                        string sPass = rsCompany.Fields.Item("U_SAPPassword").Value.ToString(); ;
                        string sqUser = rsCompany.Fields.Item("U_ServerUser").Value.ToString();
                        string sqPass = rsCompany.Fields.Item("U_ServerPass").Value.ToString();
                        gen.connectOtherCompany(server,Licserver, DB, sUser, sPass, sqUser, sqPass);

                        //oSales = (SAPbobsCOM.Documents)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                        oWareHouse = (SAPbobsCOM.Warehouses)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWarehouses);
                        string strEmployeeName = Element.GetElementsByTagName("WhsCode")[0].InnerText;
                        string strCodeQrry = "SELECT * FROM OWHS WHERE WhsCode='" + strEmployeeName + "'";
                        SAPbobsCOM.Recordset rsEmployee = ((SAPbobsCOM.Recordset)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                        rsEmployee.DoQuery(strCodeQrry);
                        if (rsEmployee.RecordCount > 0)
                        {

                            oWareHouse.GetByKey((rsEmployee.Fields.Item("WhsCode").Value.ToString()));
                            //oWareHouse.WarehouseCode = Element.GetElementsByTagName("WhsCode")[0].InnerText;
                            oWareHouse.WarehouseName = Element.GetElementsByTagName("WhsName")[0].InnerText;
                            oWareHouse.Location = Convert.ToInt32(Element.GetElementsByTagName("Location")[0].InnerText);
                            oWareHouse.WHShipToName = Element.GetElementsByTagName("WhShipTo")[0].InnerText;
                            oWareHouse.Street = Element.GetElementsByTagName("Street")[0].InnerText;
                            oWareHouse.Block = Element.GetElementsByTagName("Block")[0].InnerText;
                            oWareHouse.BuildingFloorRoom = Element.GetElementsByTagName("Building")[0].InnerText;
                            oWareHouse.ZipCode = Element.GetElementsByTagName("ZipCode")[0].InnerText;
                            oWareHouse.City = Element.GetElementsByTagName("City")[0].InnerText;
                            oWareHouse.County = Element.GetElementsByTagName("County")[0].InnerText;
                            oWareHouse.Country = Element.GetElementsByTagName("Country")[0].InnerText;
                            oWareHouse.State = Element.GetElementsByTagName("State")[0].InnerText;
                            if (Element.GetElementsByTagName("Nettable")[0].InnerText.ToString() == "Y")
                            {
                                oWareHouse.Nettable = SAPbobsCOM.BoYesNoEnum.tYES;
                            }
                            else
                            {
                                oWareHouse.Nettable = SAPbobsCOM.BoYesNoEnum.tNO;
                            }
                            //oWareHouse.Nettable = Element.GetElementsByTagName("Nettable")[0].InnerText;
                            // oWareHouse.DropShip = Element.GetElementsByTagName("DropShip")[0].InnerText;
                            if (Element.GetElementsByTagName("DropShip")[0].InnerText.ToString() == "Y")
                            {
                                oWareHouse.DropShip = SAPbobsCOM.BoYesNoEnum.tYES;
                            }
                            else
                            {
                                oWareHouse.DropShip = SAPbobsCOM.BoYesNoEnum.tNO;
                            }
                            //oWareHouse.Excisable = Element.GetElementsByTagName("Excisable")[0].InnerText;
                            if (Element.GetElementsByTagName("Excisable")[0].InnerText.ToString() == "Y")
                            {
                                oWareHouse.Excisable = SAPbobsCOM.BoYesNoEnum.tYES;
                            }
                            else
                            {
                                oWareHouse.Excisable = SAPbobsCOM.BoYesNoEnum.tNO;
                            }
                            //oWareHouse.UserFields.Fields.Item("U_WhsType").Value = Element.GetElementsByTagName("U_WhsType")[0].InnerText;
                            int iErrorCode = oWareHouse.Update();
                            if (iErrorCode != 0)
                            {
                                string sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                MessageBox.Show(sErrorMsg + "in '" + DB + "'");
                                Error = 1;

                            }
                            else
                            {
                                MessageBox.Show("Updated successfully to '" + DB + "'");
                                // MessageBox.Show("Error in Export BP Master : " + sErrorMsg);
                                // StrSql = "Update [NOR_BP_MASTER] set IntegrationStatus='I' where BPcode='" + Element.GetElementsByTagName("BPCode")[0].InnerText + "'";
                                // ConDb.QueryNonExecuteBranch(StrSql);

                            }
                        }
                        else
                        {
                            // SELECT WhsCode,WhsName,Location,WhShipTo,Street,StreetNo,Block,Building,ZipCode,City,County,Country,
                            //State,Nettable,DropShip,Excisable,U_Unit,U_WhsType FROM OWHS
                            oWareHouse.WarehouseCode = Element.GetElementsByTagName("WhsCode")[0].InnerText;
                            oWareHouse.WarehouseName = Element.GetElementsByTagName("WhsName")[0].InnerText;
                            oWareHouse.Location = Convert.ToInt32(Element.GetElementsByTagName("Location")[0].InnerText);
                            oWareHouse.WHShipToName = Element.GetElementsByTagName("WhShipTo")[0].InnerText;
                            oWareHouse.Street = Element.GetElementsByTagName("Street")[0].InnerText;
                            oWareHouse.Block = Element.GetElementsByTagName("Block")[0].InnerText;
                            oWareHouse.BuildingFloorRoom = Element.GetElementsByTagName("Building")[0].InnerText;
                            oWareHouse.ZipCode = Element.GetElementsByTagName("ZipCode")[0].InnerText;
                            oWareHouse.City = Element.GetElementsByTagName("City")[0].InnerText;
                            oWareHouse.County = Element.GetElementsByTagName("County")[0].InnerText;
                            oWareHouse.Country = Element.GetElementsByTagName("Country")[0].InnerText;
                            oWareHouse.State = Element.GetElementsByTagName("State")[0].InnerText;
                            if (Element.GetElementsByTagName("Nettable")[0].InnerText.ToString() == "Y")
                            {
                                oWareHouse.Nettable = SAPbobsCOM.BoYesNoEnum.tYES;
                            }
                            else
                            {
                                oWareHouse.Nettable = SAPbobsCOM.BoYesNoEnum.tNO;
                            }
                            //oWareHouse.Nettable = Element.GetElementsByTagName("Nettable")[0].InnerText;
                            // oWareHouse.DropShip = Element.GetElementsByTagName("DropShip")[0].InnerText;
                            if (Element.GetElementsByTagName("DropShip")[0].InnerText.ToString() == "Y")
                            {
                                oWareHouse.DropShip = SAPbobsCOM.BoYesNoEnum.tYES;
                            }
                            else
                            {
                                oWareHouse.DropShip = SAPbobsCOM.BoYesNoEnum.tNO;
                            }
                            //oWareHouse.Excisable = Element.GetElementsByTagName("Excisable")[0].InnerText;
                            if (Element.GetElementsByTagName("Excisable")[0].InnerText.ToString() == "Y")
                            {
                                oWareHouse.Excisable = SAPbobsCOM.BoYesNoEnum.tYES;
                            }
                            else
                            {
                                oWareHouse.Excisable = SAPbobsCOM.BoYesNoEnum.tNO;
                            }
                            //oWareHouse.UserFields.Fields.Item("U_WhsType").Value = Element.GetElementsByTagName("U_WhsType")[0].InnerText;
                            //oWareHouse.WarehouseCode = Element.GetElementsByTagName("SlpName")[0].InnerText;

                            //oWareHouse.CommissionForSalesEmployee = Convert.ToDouble(Element.GetElementsByTagName("Commission")[0].InnerText);
                            //oWareHouse.CommissionGroup = Convert.ToInt32(Element.GetElementsByTagName("GroupCode")[0].InnerText);
                            // oSalesPerson.Locked = Element.GetElementsByTagName("Locked")[0].InnerText;
                            //oSalesPerson.EmployeeID = Convert.ToInt32(Element.GetElementsByTagName("EmpID")[0].InnerText);
                            int iErrorCode = oWareHouse.Add();
                            if (iErrorCode != 0)
                            {
                                string sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                //MessageBox.Show(sErrorMsg + "in '" + DB + "'", 1, "Ok", "", "");
                                Error = 1;

                            }
                            else
                            {
                                MessageBox.Show("Updated successfully to '" + DB + "'");
                                // MessageBox.Show("Error in Export BP Master : " + sErrorMsg);
                                // StrSql = "Update [NOR_BP_MASTER] set IntegrationStatus='I' where BPcode='" + Element.GetElementsByTagName("BPCode")[0].InnerText + "'";
                                // ConDb.QueryNonExecuteBranch(StrSql);

                            }
                        }
                    }
                    //    oRsInv.MoveNext();
                    //}
                }
            }
            catch (Exception E) { MessageBox.Show(E.Message +"WareHouse"); }


        }



        //---------------------------------------Location

        public void Location()
        {
            try
            {
                string sPath = "";
                string FileName = "Location.xml";
                string StrSql = "";
                SAPbobsCOM.WarehouseLocations oLocation;
                int Error = 0;
                General gen = new General();
                if (!File.Exists(sPath + FileName))
                { File.Create(sPath + FileName); }
                int recCount = 0;
                System.Xml.XmlDocument oXmlDoc = null;
                string sXmlString = null;
                StrSql = @"SELECT Location,ISNULL(Street,'')Street,ISNULL(Block,'')Block,ISNULL(Building,'')Building,ISNULL(ZipCode,'')ZipCode,ISNULL(City,'')City,ISNULL(County,'')County,ISNULL(Country,'')Country,ISNULL(State,'')State,ISNULL(TanNo,'')TanNo,ISNULL(PanNo,'')PanNo FROM OLCT  ";
                DataSet objDataSet = ConDb.DbDataFromSAP(StrSql);
                sXmlString = objDataSet.GetXml();
                // sXmlString = oRsInv.GetAsXML();
                oXmlDoc = new System.Xml.XmlDocument();
                oXmlDoc.LoadXml(sXmlString);
                oXmlDoc.Save((sPath + FileName));



                XmlDocument reader = new XmlDocument();
                XmlDocument readerlines = new XmlDocument();
                IFormatProvider ifp = new System.Globalization.CultureInfo("en-US", true);
                reader.Load(sPath + FileName);

                XmlNodeList list = reader.GetElementsByTagName(gen.row1);


                foreach (XmlNode node in list)
                {
                    XmlElement Element = (XmlElement)node;
                    string UnitGetQrry = "SELECT * FROM [@NOR_UNITMASTER]";
                    SAPbobsCOM.Recordset oRsInv = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRsInv.DoQuery(UnitGetQrry);
                    string strUnit = oRsInv.Fields.Item("Code").Value.ToString();
                    //while (!oRsInv.EoF)
                    //{
                    //string UnitCode = oRsInv.Fields.Item("Code").Value.ToString();
                    string QRY1 = "Select * from [@NOR_BRANCH_DTL] Where U_UnitId ='" + strUnit + "'";
                    SAPbobsCOM.Recordset rsCompany = (SAPbobsCOM.Recordset)General.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                    rsCompany.DoQuery(QRY1);
                    if (rsCompany.RecordCount > 0)
                    {
                        string server = rsCompany.Fields.Item("U_ServerName").Value.ToString();
                        string Licserver = rsCompany.Fields.Item("U_Licserver").Value.ToString();
                        string DB = rsCompany.Fields.Item("U_CompanyDB").Value.ToString();
                        string sUser = rsCompany.Fields.Item("U_SAPUserName").Value.ToString();
                        string sPass = rsCompany.Fields.Item("U_SAPPassword").Value.ToString(); ;
                        string sqUser = rsCompany.Fields.Item("U_ServerUser").Value.ToString();
                        string sqPass = rsCompany.Fields.Item("U_ServerPass").Value.ToString();
                        gen.connectOtherCompany(server,Licserver, DB, sUser, sPass, sqUser, sqPass);

                        //oSales = (SAPbobsCOM.Documents)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                        oLocation = (SAPbobsCOM.WarehouseLocations)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWarehouseLocations);
                        string strEmployeeName = Element.GetElementsByTagName("Location")[0].InnerText;
                        string strCodeQrry = "SELECT * FROM OLCT WHERE Location='" + strEmployeeName + "'";
                        SAPbobsCOM.Recordset rsEmployee = ((SAPbobsCOM.Recordset)(Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                        rsEmployee.DoQuery(strCodeQrry);
                        if (rsEmployee.RecordCount > 0)
                        {
                            string strUpdate = "UPDATE OLCT SET Location='" + Element.GetElementsByTagName("Location")[0].InnerText + "',Country='" + Element.GetElementsByTagName("Country")[0].InnerText + "',State='" + Element.GetElementsByTagName("State")[0].InnerText + "',PanNo='" + Element.GetElementsByTagName("PanNo")[0].InnerText + "',City='" + Element.GetElementsByTagName("City")[0].InnerText + "' WHERE Location='" + Element.GetElementsByTagName("Location")[0].InnerText + "'";
                            SAPbobsCOM.Recordset rsCompany1 = (SAPbobsCOM.Recordset)Global.oCompny2.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            //SAPbobsCOM.Recordset rsCompany = ((SAPbobsCOM.Recordset)(Global.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                            rsCompany1.DoQuery(strUpdate);

                            MessageBox.Show("Updated successfully to '" + DB + "'");
                              
                        }
                        else
                        {
                            // SELECT WhsCode,WhsName,Location,WhShipTo,Street,StreetNo,Block,Building,ZipCode,City,County,Country,
                            //State,Nettable,DropShip,Excisable,U_Unit,U_WhsType FROM OWHS
                           /// oLocation.WarehouseCode = Element.GetElementsByTagName("WhsCode")[0].InnerText;
                            //Location,ISNULL(Street,'')Street,ISNULL(Block,'')Block,ISNULL(Building,'')Building,ISNULL(ZipCode,'')ZipCode,ISNULL(City,'')City,ISNULL(County,'')County,ISNULL(Country,'')Country,ISNULL(State,'')State,ISNULL(TanNo,'')TanNo,ISNULL(PanNo,'')PanNo  
                            oLocation.Name = Element.GetElementsByTagName("Location")[0].InnerText;
                            oLocation.Street = Element.GetElementsByTagName("Street")[0].InnerText;
                            //oLocation.BuildingFloorRoom = Element.GetElementsByTagName("WhShipTo")[0].InnerText;
                            // oLocation.Street = Element.GetElementsByTagName("Street")[0].InnerText;
                            oLocation.Block = Element.GetElementsByTagName("Block")[0].InnerText;
//oLocation.BuildingFloorRoom = Element.GetElementsByTagName("Building")[0].InnerText;
                            oLocation.ZipCode = Element.GetElementsByTagName("ZipCode")[0].InnerText;
                            oLocation.City = Element.GetElementsByTagName("City")[0].InnerText;
                            oLocation.County = Element.GetElementsByTagName("County")[0].InnerText;
                            oLocation.Country = Element.GetElementsByTagName("Country")[0].InnerText;
                            oLocation.State = Element.GetElementsByTagName("State")[0].InnerText;
                            oLocation.PANNumber = Element.GetElementsByTagName("PanNo")[0].InnerText;
                            oLocation.TANNumber = Element.GetElementsByTagName("TanNo")[0].InnerText;

                           
                            //oWareHouse.UserFields.Fields.Item("U_WhsType").Value = Element.GetElementsByTagName("U_WhsType")[0].InnerText;
                            //oWareHouse.WarehouseCode = Element.GetElementsByTagName("SlpName")[0].InnerText;

                            //oWareHouse.CommissionForSalesEmployee = Convert.ToDouble(Element.GetElementsByTagName("Commission")[0].InnerText);
                            //oWareHouse.CommissionGroup = Convert.ToInt32(Element.GetElementsByTagName("GroupCode")[0].InnerText);
                            // oSalesPerson.Locked = Element.GetElementsByTagName("Locked")[0].InnerText;
                            //oSalesPerson.EmployeeID = Convert.ToInt32(Element.GetElementsByTagName("EmpID")[0].InnerText);
                            int iErrorCode = oLocation.Add();
                            if (iErrorCode != 0)
                            {
                                string sErrorMsg = Global.oCompny2.GetLastErrorDescription();
                                MessageBox.Show(sErrorMsg + "in '" + DB + "'");
                                Error = 1;

                            }
                            else
                            {
                                MessageBox.Show("Updated successfully to '" + DB + "'");
                                // MessageBox.Show("Error in Export BP Master : " + sErrorMsg);
                                // StrSql = "Update [NOR_BP_MASTER] set IntegrationStatus='I' where BPcode='" + Element.GetElementsByTagName("BPCode")[0].InnerText + "'";
                                // ConDb.QueryNonExecuteBranch(StrSql);

                            }
                        }
                    }
                    //    oRsInv.MoveNext();
                    //}
                }
            }
            catch(Exception E) { MessageBox.Show(E.Message+ "Location"); }

        }

        //-----------------------------------------------

   
    }
#endregion
}