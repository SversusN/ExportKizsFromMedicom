using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportKizsFromMedicom
{
    public static class DBSetToCsv
    {
        internal static void ExportDataSetToCsvFile(DataSet _DataSet, string DestinationCsvDirectory, out int progress, out int allRows, out string fileName)
        { 
            progress = 0;
            allRows = _DataSet.Tables[0].Rows.Count;
            fileName = string.Empty;
            try
            {
                foreach (DataTable DDT in _DataSet.Tables)
                {
                    String MyFile = @DestinationCsvDirectory + "\\KizRemains"  + DateTime.Now.ToString("yyyyMMddhhMMssffff") + ".csv";//+ DateTime.Now.ToString("ddMMyyyyhhMMssffff")
                    fileName = MyFile;
                    using (var outputFile = File.CreateText(MyFile))
                    {
                        String CsvText = string.Empty;

                        foreach (DataColumn DC in DDT.Columns)
                        {
                            if (CsvText != "")
                                CsvText = CsvText + ";" + DC.ColumnName.ToString();
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
                                    CsvText = CsvText + ";" + DDR[DCC.ColumnName.ToString()].ToString();
                                else
                                    CsvText = DDR[DCC.ColumnName.ToString()].ToString();
                            }
                            outputFile.WriteLine(CsvText.ToString().TrimEnd(';'));
                            CsvText = string.Empty;
                            progress += 1;
                        }
                        System.Threading.Thread.Sleep(1000);
                    }
                }
            }
            catch (Exception Ex)
            {
                throw Ex;
            }
        }
    }
}
