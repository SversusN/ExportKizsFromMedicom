using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportKizsFromMedicom
{
    public partial class ExportKizs : Form
    {
        public ExportKizs()
        {
            InitializeComponent();
            connectionString = textBox1.Text;

        }
        private string connectionString;

        private void button1_Click(object sender, EventArgs e)
        {
            //string constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Bases;Extended Properties=dBASE IV";

            //string connString = @"Provider=vfpoledb.1;Data Source=C:\Bases;Exclusive=false;Nulls=false";
            string connString = connectionString;
            using (OleDbConnection con = new OleDbConnection(connString))
            {
                var sql = "select pm.cim,p.nomk_ls, p.series, p.price from " + @"productmark.dbf pm"+" inner join " + @"product.dbf p" + " on p.idproduct = pm.idproduct where pm.ko_src > 0";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                con.Open();
                DataSet ds = new DataSet(); ;
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(ds);
                string destinationFile = Environment.CurrentDirectory;// + "\\" + "sample.csv";
                int progress = 0;
                int allRows = 0 ;
                string fileName = string.Empty;
                DBSetToCsv.ExportDataSetToCsvFile(ds, destinationFile,out progress,out allRows, out fileName);
                progressBar1.Maximum = allRows;
                progressBar1.Value = progress;
                MessageBox.Show($"Выгружен файл {fileName}");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            connectionString = textBox1.Text;
        }
    }
}

