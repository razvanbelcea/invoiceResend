using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace invoiceResend
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private bool _isfilled;
        private bool _filefound;
        private string _result;
        private int _realpath;
        private int maxvalue;

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.Refresh();
            filltable();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (_isfilled)
            {
                if (radioButton1.Checked)
                {
                    Task.Factory.StartNew(Checkinvoice);
                } 
                else if (radioButton2.Checked)
                {
                    Checkforro();
                }       
            }
            else
            {
                MessageBox.Show("Data is not loaded!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool filltable()
        {
            try
            {
                OpenFileDialog ofImport = new OpenFileDialog();
                ofImport.Title = "Select file";
                ofImport.InitialDirectory = @"c:\";
                ofImport.Filter = "Excel Sheet(*.xlsx)|*.xlsx";
                ofImport.FilterIndex = 1;
                ofImport.RestoreDirectory = true;

                if (ofImport.ShowDialog() == DialogResult.OK)
                {

                    string path = System.IO.Path.GetFullPath(ofImport.FileName);
                    var dbt = new Util();
                    var data = dbt.Filldatatable(path);
                    data.Columns.Add("Status", typeof(System.String)).SetOrdinal(0);
                    dataGridView1.DataSource = data;
                    maxvalue = dataGridView1.Rows.Count;
                    progressBar1.Maximum = maxvalue;
                    return _isfilled = true;                    
                }
                return _isfilled = false;
            }
            catch (Exception ex)
            {
               Console.WriteLine(ex.Message);
                return false;
            }
        }

        private void Checkinvoice()
        {
            this.Invoke(new Action(() =>
            {
                radioButton1.Enabled = false;
                radioButton2.Enabled = false;
                button1.Enabled = false;
                button2.Enabled = false;
            }));           
            var i = 0;
            var util = new Util();
            var data = util.Filldatatable(AppDomain.CurrentDomain.BaseDirectory + @"RUStores.xlsx");
            foreach (DataRow[] results in from object rows in dataGridView1.Rows select int.Parse(dataGridView1.Rows[i].Cells[1].Value.ToString()) into v select data.Select("store =" + v))
            {
                util.ip = results[0].ItemArray[1].ToString();
                var fullname = util.ip + ".mpos.madm.net";
                string connectionString = "Data Source=" + fullname + ";Database=TPCentralDB;Integrated Security=SSPI;";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    try
                    {
                        connection.Open();
                        if (connection.State == ConnectionState.Open)
                        {
                            var lta = util.finttheinvoice(dataGridView1.Rows[i].Cells[3].Value.ToString(),
                                dataGridView1.Rows[i].Cells[2].Value.ToString());
                            var invname = dataGridView1.Rows[i].Cells[1].Value + "_" +
                                          dataGridView1.Rows[i].Cells[2].Value + "_" + lta;
                            var szdate = util.getdate(dataGridView1.Rows[i].Cells[3].Value.ToString());
                            if (szdate != "")
                            {
                                Findxml(results[0].ItemArray[1].ToString(),szdate,invname,i, results[0].ItemArray[1].ToString());
                            }
                            else
                            {
                                dataGridView1.Rows[i].Cells[0].Value = "invoice is missing from DB";
                                Console.WriteLine("invoice number not found..");
                            }                           
                        }
                    }
                    catch (SqlException ex)
                    {
                        dataGridView1.Rows[i].Cells[0].Value = ex.Message;
                        dataGridView1.Rows[i].Cells[0].Style.ForeColor = Color.Red;
                        Console.WriteLine(ex.Message);
                    }
                }
                i = i + 1;
                this.Invoke(new Action(() =>
                {
                  //  TaskbarProgress.SetState(this.Handle, TaskbarProgress.TaskbarStates.Indeterminate);
                    TaskbarProgress.SetValue(this.Handle, i, maxvalue);
                    TaskbarProgress.SetState(this.Handle, TaskbarProgress.TaskbarStates.Error);
                    progressBar1.Value = i;
                }));
            }
            this.Invoke(new Action(() =>
            {
                radioButton1.Enabled = true;
                radioButton2.Enabled = true;
                button1.Enabled = true;
                button2.Enabled = true;
                progressBar1.Value = 0;
            }));
        }

        private string Checkforro()
        {
            this.Invoke(new Action(() =>
            {
                radioButton1.Enabled = false;
                radioButton2.Enabled = false;
                button1.Enabled = false;
                button2.Enabled = false;
            }));
            foreach (var row in dataGridView1.Rows)
            {
                
            }
            return "succes";
        }

        private void Findxml(string hostname, string szdate, string invname, int i, string ipul)
        {
            var processedpath = @"\\" + hostname + ".mpos.madm.net" +
                                @"\e$\Journal\Transactions\" + szdate + @"\Processed";
            var failedpath = @"\\" + hostname + ".mpos.madm.net" +
                             @"\e$\Journal\Transactions\" + szdate + @"\Failed";
            if (Directory.Exists(processedpath))
            {
                DirectoryInfo di = new DirectoryInfo(processedpath);
                foreach (var file in di.EnumerateFiles(invname + "*"))
                {
                    _realpath = 1;
                    var util = new Util();
                    _result = util.copyinvoice(file.DirectoryName, file.Name, ipul);
                    _filefound = true;
                    break;
                }
                if (_realpath != 1)
                {
                    if (Directory.Exists(failedpath))
                    {
                        DirectoryInfo did = new DirectoryInfo(failedpath);
                        foreach (var file in did.EnumerateFiles(invname + "*"))
                        {
                            _realpath = 2;
                            var util = new Util();
                            _result = util.copyinvoice(file.DirectoryName, file.Name, ipul);
                            _filefound = true;
                            break;
                        }
                    }
                }               
            }
            dataGridView1.Rows[i].Cells[0].Value = _filefound ? _result : "XML not found";
            }

        private void Findpdf()
        {
            
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                button1.Enabled = true;
                button2.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
                button2.Enabled = false;
            }
        }
    }
}
