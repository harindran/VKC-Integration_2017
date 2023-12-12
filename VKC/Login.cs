using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace BranchIntegrator
{
    public partial class Login : Form
    {
        SAPbobsCOM.Company oCompany;
        public Login()
        {
            InitializeComponent();
        }

        private void Login_Load(object sender, EventArgs e)
        {

            SAPbobsCOM.Recordset oRecordSet;


            ConectApp("", "manager", "manager");
            oRecordSet = oCompany.GetCompanyList();
            while (!(oRecordSet.EoF == true))
            {
                // add the value of the first field of the Recordset
                ddlCompany.Items.Add(oRecordSet.Fields.Item(0).Value);
                // move the record pointer to the next row
                oRecordSet.MoveNext();
            }

            oCompany.Disconnect();

        }


        private  void start()
        {

           
            
            Application.Run(new Form1(oCompany ));

        }

        private void btConnect_Click(object sender, EventArgs e)
        {
            string sErrorMsg = "";
            ConectApp(ddlCompany.Text , txtUser.Text , txtPass.Text );
            if (oCompany.Connected == true)
            {
               // Thread t = new Thread(new ThreadStart(start));
               // t.Start();
                Form1 f = new Form1(oCompany);
                f.Show();
                this.Hide();
              
                
            }

            else
            {
                sErrorMsg = oCompany.GetLastErrorDescription();
                // wMsgBox(sErrorMsg);
            }


           
        }
        private void ConectApp(string company, string user, string pass)
        {

            int iErrorCode = 0;
            string sErrorMsg = "";
            oCompany = new SAPbobsCOM.Company();
     
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English; //  change to your language
            oCompany.UseTrusted = false;
            oCompany.CompanyDB = company;
            oCompany.UserName = user;
            oCompany.Password = pass;
            oCompany.DbUserName = "sa";




            oCompany.Server = "WIN-WSEGEPIXKQP"; //  change to your company server
            oCompany.DbPassword = "sapb1@1234";
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
       

            //oCompany.Server = "SUNNOR\\SUN"; //  change to your company server
            //oCompany.DbPassword = "sapb1";
            //oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005;
         
            iErrorCode = oCompany.Connect();
            if (iErrorCode != 0)
            {
                if (company != string.Empty )
                MessageBox.Show( oCompany.GetLastErrorDescription());
                //SBO_Application.MessageBox(sErrorMsg, 1, "Ok", "", "");
            }
            if (oCompany.Connected == true)
            {

            }
        }
    }
}