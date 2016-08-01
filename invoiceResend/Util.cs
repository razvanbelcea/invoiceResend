using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace invoiceResend
{
    class Util
    { 
        public string finttheinvoice(string invoiceno, string tillno)
        {      
            var ltanoquery = "select top 1 lTaNmbr from maitxinvoice where " +
                             "lWorkstationNmbr =" + "'" + tillno + "'" + " and " + "szInvoiceSequenceNumber=" +
                             "'" + invoiceno.PadLeft(6, '0') + "'";
            
           var lTaNmbr = getstoreinfo(ltanoquery, true);

            return lTaNmbr;
        }

        public string getdate(string invoiceno)
        {
            var Datequery = "select top 1 szdate from maitxinvoice where szInvoiceSequenceNumber =" + "'" + invoiceno.PadLeft(6, '0') + "'";
            var szDate = getstoreinfo(Datequery, false);
            return szDate;
        }

        public string ip;
        private string getstoreinfo(string query, bool str)
        {
            try
            {
                string queryString = query;
                var fullname = ip + ".mpos.madm.net";
                string connectionString = "Data Source=" + fullname + ";Database=TPCentralDB;Integrated Security=SSPI;";
                var result = "";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlCommand command = new SqlCommand(queryString, connection);
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    try
                    {
                        while (reader.Read())
                        {
                            if (str == false)
                            { result = Convert.ToString(reader.GetString(0)); }
                            else
                            {
                                result = Convert.ToString(reader.GetInt32(0));
                            }
                            return result;
                        }
                        return result;
                    }
                    finally
                    {
                        reader.Close();
                        connection.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }

        public DataTable Filldatatable(string filepath)
        {
            try
            {
                String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                filepath +
                ";Extended Properties='Excel 12.0 XML;HDR=Yes;';";

                OleDbConnection con = new OleDbConnection(constr);
                con.Open();
               var dtSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
              var  Sheet1 = dtSchema.Rows[0].Field<string>("TABLE_NAME");
                OleDbCommand oconn = new OleDbCommand("Select * From [" + Sheet1 + "] where store is not null", con);

                OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
                DataTable data = new DataTable();
                sda.Fill(data);
                return data;
            }
            catch (Exception ex)
            {
               MessageBox.Show("FillDataTable: " + ex.Message);
                return null;
            }
        }

        public string copyinvoice(string file, string filename, string ipul)
        {
            try
            {
                var fullname = ipul + ".mpos.madm.net";
                var destinationpath = @"\\" + fullname + @"\e$\TpDotnet\Excise\";
                if (Directory.Exists(destinationpath))
                {
                    if (!File.Exists(destinationpath + filename))
                    {
                     //   File.Copy(file+ @"\" + filename, destinationpath + @"\" + filename);
                        Console.WriteLine(filename + " copied to " + destinationpath);
                        return "Success"; 
                    }
                    else
                    {
                        return "file already present";
                    }
                }
                else
                {
                    return "directory not found";
                }
            }
            catch (Exception ex)
            {
                return "ups:" + ex.Message;
            }

        }

    }
}
