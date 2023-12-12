using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.Windows.Forms;
using System.IO;
using System.Threading;



namespace BranchIntegrator
{
    class DbConnection
    {
        public static SqlConnection objConBranch = new SqlConnection();
        public static string strConBranch = String.Empty;
        public static string strFileBranch = AppDomain.CurrentDomain.BaseDirectory + "\\BIMS_SANKARLA.dbo";
        public static SqlConnection objConSAP = new SqlConnection();

        public static string strConSAP = String.Empty;
        public static string strFileSAP = AppDomain.CurrentDomain.BaseDirectory + "\\NEAHDB.dbo";
        public static SqlConnection objConSAP1 = new SqlConnection();

        public static string strConSAP1 = String.Empty;
        public static string strFileSAP1 = AppDomain.CurrentDomain.BaseDirectory + "\\BRANCH_SAP.dbo";
        SqlDataAdapter sSqlDataAdapter = new SqlDataAdapter();

       


        public DbConnection()
        {
           try
            {
            if (objConBranch.State == ConnectionState.Open) objConBranch.Close() ;
            {
                StreamReader Read = new StreamReader(Application.StartupPath + "\\Profile.ini");
                string[] s = Read.ReadToEnd().Replace("\r\n\t", "\r").Split('\r');
                strConBranch = "Data Source = " + s[0].Split(':')[1].ToString() + ";Initial Catalog=" + s[2].Split(':')[1].ToString() + ";User ID=" + s[5].Split(':')[1] + ";password=" + s[6].Split(':')[1].ToString();
               // strConBranch = "Data Source = " + s[0].Split(':')[1].ToString() + ";Initial Catalog=" + s[2].Split(':')[1].ToString() + ";User ID=" + s[5].Split(':')[1] + ";password=" + s[6].Split(':')[1].ToString() + ";Network Library=DBMSSOCN";
                objConBranch.ConnectionString = strConBranch;
                objConBranch.Open();
            }

                if (objConSAP.State == ConnectionState.Open) objConSAP.Close();
                {
                    StreamReader Read = new StreamReader(Application.StartupPath + "\\Profile.ini");
                    string[] s = Read.ReadToEnd().Replace("\r\n\t", "\r").Split('\r');
                    strConSAP = "Data Source = " + s[0].Split(':')[1].ToString() + ";Initial Catalog=" + s[2].Split(':')[1].ToString() + ";User ID=" + s[5].Split(':')[1] + ";password=" + s[6].Split(':')[1].ToString();
                    objConSAP.ConnectionString = strConSAP;
                    objConSAP.Open();
                }
                if (objConSAP1.State == ConnectionState.Open) objConSAP1.Close();
                {
                    StreamReader Read = new StreamReader(Application.StartupPath + "\\Profile.ini");
                    string[] s = Read.ReadToEnd().Replace("\r\n\t", "\r").Split('\r');
                    strConSAP1 = "Data Source = " + s[0].Split(':')[1].ToString() + ";Initial Catalog=" + s[2].Split(':')[1].ToString() + ";User ID=" + s[5].Split(':')[1] + ";password=" + s[6].Split(':')[1].ToString();
                    objConSAP1.ConnectionString = strConSAP1;
                    objConSAP1.Open();
                }
            }
            catch
            {
                MessageBox.Show("Error in connection.");
            }
            }

        public DataSet DbDataFromBranch(string StrSql)
        {
            DataSet objDs = new DataSet();
            objDs.Clear();
            SqlDataAdapter da = new SqlDataAdapter(StrSql, objConBranch);
            da.Fill(objDs, "Data");
            return objDs;
        }
        public DataSet DbDataFromSAP   (string StrSql)
        {
            DataSet objDs = new DataSet();
            objDs.Clear();
            SqlDataAdapter da = new SqlDataAdapter(StrSql, objConSAP);
            da.Fill(objDs, "Data");
            return objDs;
        }
        public DataSet DbDataFromSAP1  (string StrSql)
        {
            DataSet objDs = new DataSet();
            objDs.Clear();
            SqlDataAdapter da = new SqlDataAdapter(StrSql, objConSAP1);
            da.Fill(objDs, "Data");
            return objDs;
        }

        public SqlCommand QueryNonExecuteBranch(string StrSql)
        {
            if (objConBranch.State == ConnectionState.Closed) objConBranch.Open();
            SqlCommand objCmd = new SqlCommand(StrSql, objConBranch);
            objCmd.ExecuteNonQuery();
            objConBranch.Close();
            return objCmd;
        }
        public SqlCommand  QueryNonExecuteSAP   (string StrSql)
        {
            if (objConSAP.State == ConnectionState.Closed) objConSAP.Open();
            SqlCommand objCmd = new SqlCommand(StrSql, objConSAP);
            objCmd.ExecuteNonQuery();
            objConSAP.Close();
            return objCmd;
       
        }
        public void   QueryNonExecuteSAP1  (string StrSql)
        {
            if (objConSAP1.State == ConnectionState.Closed) objConSAP1.Open();
            SqlCommand objCmd = new SqlCommand(StrSql, objConSAP);
            objCmd.ExecuteNonQuery();
            objConSAP1.Close();
        }       
        public string ScalarExecuteBranch  (string StrSql)
        {
            if (objConBranch.State == ConnectionState.Closed) objConBranch.Open();
            SqlCommand objCmdBranch = new SqlCommand(StrSql, objConBranch);
            if (objCmdBranch.ExecuteScalar() is DBNull)
                return "";
            return "Success";
        }
        public string ScalarExecuteSAP     (string StrSql)
        {
            if (objConSAP.State == ConnectionState.Closed) objConSAP.Open();
            SqlCommand objCmdSAP = new SqlCommand(StrSql, objConSAP);
            if (objCmdSAP.ExecuteScalar() is DBNull)
                return "";
            return "Success";
        }
        public string ScalarExecuteSAP1    (string StrSql)
        {
            if (objConSAP1.State == ConnectionState.Closed) objConSAP1.Open();
            SqlCommand objCmdSAP1 = new SqlCommand(StrSql, objConSAP1);
            if (objCmdSAP1.ExecuteScalar() is DBNull)
                return "";
            return "Success";
        }

        public SqlDataReader DbReaderBranch(string StrSql)
        {
            if (objConBranch.State == ConnectionState.Closed) objConBranch.Open();
            SqlCommand objCmdBranch = new SqlCommand(StrSql, objConBranch);
            SqlDataReader objReaderBranch = objCmdBranch.ExecuteReader();
            return objReaderBranch;
        }
        public SqlDataReader DbReaderSAP   (string StrSql)
        {
            if (objConSAP.State == ConnectionState.Closed) objConSAP.Open();
            SqlCommand objCmdSAP = new SqlCommand(StrSql, objConSAP);
            SqlDataReader objReaderSAP = objCmdSAP.ExecuteReader();
            return objReaderSAP;
        }
        public SqlDataReader DbReaderSAP1  (string StrSql)
        {
            if (objConSAP1.State == ConnectionState.Closed) objConSAP1.Open();
            SqlCommand objCmdSAP1 = new SqlCommand(StrSql, objConSAP1);
            SqlDataReader objReaderSAP1 = objCmdSAP1.ExecuteReader();
            return objReaderSAP1;
        }
    }
}
