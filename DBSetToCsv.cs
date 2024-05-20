using System;
using System.Data;
using System.IO;
using System.Linq;

namespace ExportKizsFromMedicom
{
    public static class DBSetToCsv
    {
        internal static void ExportDataSetToCsvFile(DataSet _DataSet, string DestinationCsvDirectory, string name, out int progress, out int allRows, out string fileName)
        { 
            progress = 0;
           
            fileName = string.Empty;
            try
            {

                var dData = _DataSet.Tables[0].DefaultView.ToTable();
                var DDT = dData.AsEnumerable()
                    .GroupBy(r => new { Col1 = r["idproduct"] })
                    .Select(g => g.OrderBy(r => r["Товар"]).First())
                    .CopyToDataTable();


                foreach (var d in DDT.AsEnumerable())
                {

                    var kizs = dData.AsEnumerable()

                                           .Where(dt => dt["idproduct"].Equals(d["idproduct"]))

                                           .Select(e => e["sgtin"].ToString().Trim())

                                           .ToArray();

                    var kiz = string.Join("^", kizs);
                   
                    d.SetField("sgtin", kiz);
                }
                allRows = DDT.Rows.Count;

                string MyFile = @DestinationCsvDirectory + $"\\{name}"  + DateTime.Now.ToString("yyyyMMddhhMMssffff") + ".csv";//+ DateTime.Now.ToString("ddMMyyyyhhMMssffff")
                    fileName = MyFile;
            

                    using (var outputFile = File.CreateText(MyFile))
                    {
                        string CsvText = string.Empty;

                        foreach (DataColumn DC in DDT.Columns)
                        {
                            if (CsvText != "")
                                CsvText = CsvText + ";" + DC.ColumnName.ToString().Trim();
                            else
                                CsvText = DC.ColumnName.ToString();
                        }
                        outputFile.WriteLine(CsvText.ToString().TrimEnd(';'));
                        CsvText = string.Empty;
                        
                        foreach (DataRow DDR in DDT.Rows)
                        {
                            foreach (DataColumn DCC in DDT.Columns)
                            {
                                if (CsvText != "")
                                    CsvText = CsvText + ";" + DDR[DCC.ColumnName.ToString().Trim()].ToString().Trim();
                                else
                                    CsvText = DDR[DCC.ColumnName.ToString().Trim()].ToString().Trim();
                            }
                            outputFile.WriteLine(CsvText.ToString().Trim().TrimEnd(';'));
                            CsvText = string.Empty;
                            progress += 1;
                        }
                        System.Threading.Thread.Sleep(1000);
                    }
                
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
        }
    }
}
