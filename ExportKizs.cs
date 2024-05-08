using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ExportKizsFromMedicom
{
    public partial class ExportKizs : Form
    {
        public ExportKizs()
        {
            InitializeComponent();
            connectionString = textBox1.Text;
            queryRemains = textBox2.Text;
            queryKizs = textBox3.Text;
            
        }
        private string connectionString;

        public string queryRemains;// = "SELECT \r\nn.\"NomenclatureId\" as CODE,\r\n   Concat(n.\"TorgName\", ' ' ,n.\"DosageNameRs\",' ', n.\"PackNamePrimary\") as GOODS, \r\n  n.\"ProducerName\", \r\n  n.\"IsGnvlp\",\r\n  COALESCE(cast(p.\"RestAmount\" as NUMERIC), cast(p.\"Amount\" as NUMERIC)) as QTY,\r\n  pp.\"Price\", \r\n  p.\"Ean13\", \r\n  p.\"SupplierPrice\", \r\n  p.\"SupplierPriceWithOutNds\", \r\n  n.\"CountryName\", \r\ncase when (pa.\"Series\" = ''::text) then pa.\"Article\"\r\nwhen (pa.\"Series\" = ''::text and  pa.\"Article\" = ''::text )then replace(replace((p.\"ExpireDate\"::Date - interval '4 year')::text, ' 00:00:00', ''),'-','')\r\nelse pa.\"Series\" END as Series,\r\n\r\n  p.\"ExpireDate\",\r\n  i.\"InvoiceNumber\",\r\n  i.\"InvoiceDate\",\r\n  p.\"LocProductId\",  pp.\"Price\"*p.\"RestAmount\"\r\nFROM \r\n  public.\"Products\" p\r\n inner join  public.\"Nomenclature\" n on n.\"NomenclatureId\" =p.\"NomenclatureId\"\r\n inner join public.\"PriceProducts\" pp on pp.\"LocProductId\" = p.\"LocProductId\"\r\n left join public.\"ProductAttributes\" pa on pa.\"LocProductId\"  = p.\"LocProductId\" \r\nleft join public.\"Invoices\" i on p.\"LocInvoiceId\" = i.\"LocInvoiceId\"\r\nWHERE \r\n \r\n(p.\"RestAmount\" > 0 )\r\norder by n.\"TorgName\"";
        public string queryKizs; //= "select pm.cim,p.nomk_ls, p.series, p.price from " + @"productmark.dbf pm" + " inner join " + @"product.dbf p" + " on p.idproduct = pm.idproduct where pm.ko_src > 0";


        private void button1_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
          
            string connString = connectionString;

            if (queryKizs == null && queryRemains == null) {
                MessageBox.Show("Так нельзя! Скрипты заполнить надо.");
            }
            ExportCsv(connString, "remains" ,textBox2.Text);
            ExportCsv(connString, "kizs", textBox3.Text);

        }

        private void ExportCsv(string connection,string name, string query)
        {
            Console.WriteLine(query);
            using OleDbConnection con = new(connection);
            OleDbCommand cmd = new(query, con);
            con.Open();
            DataSet ds = new(); ;
            OleDbDataAdapter da = new(cmd);
            try
            {
                da.Fill(ds);
            } catch (Exception ex) { MessageBox.Show(ex.ToString(),"Проверь запрос!");return; }
            string destinationFile = Environment.CurrentDirectory;
            string fileName = string.Empty;
            DBSetToCsv.ExportDataSetToCsvFile(ds, destinationFile,name , out int progress, out int allRows, out fileName);
            progressBar1.Maximum = allRows;
            progressBar1.Value = progress;
            MessageBox.Show($"Выгружен файл {fileName}");
            con.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            connectionString = textBox1.Text;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            queryRemains = textBox2.Text;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            queryKizs = textBox3.Text;
        }

    }
}

