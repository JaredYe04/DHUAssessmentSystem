using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 考核系统.Utils
{
    internal class FileIO
    {
        //DataGridView转为CSV
        public static void DataGridViewToCSV(DataGridView dataGridView, string fileName)
        {
            string stOutput = "";
            // Export titles:
            string sHeaders = "";
            for (int j = 0; j < dataGridView.Columns.Count; j++)
                sHeaders = sHeaders.ToString() + Convert.ToString(dataGridView.Columns[j].HeaderText) + ",";
            stOutput += sHeaders + "\n";
            // Export data.
            for (int i = 0; i < dataGridView.RowCount - 1; i++)
            {
                string stLine = "";
                for (int j = 0; j < dataGridView.Rows[i].Cells.Count; j++)
                    stLine = stLine.ToString() + Convert.ToString(dataGridView.Rows[i].Cells[j].Value) + ",";
                stOutput += stLine + "\n";
            }
            Encoding encoding = Encoding.GetEncoding("gb2312");
            byte[] output = encoding.GetBytes(stOutput);
            System.IO.FileStream fs = new System.IO.FileStream(fileName, System.IO.FileMode.Create);
            System.IO.BinaryWriter bw = new System.IO.BinaryWriter(fs);
            bw.Write(output, 0, output.Length); //write the encoded file
            bw.Flush();
            bw.Close();
            fs.Close();
        }


        //CSV转为DataGridView
        public static void CSVToDataGridView(DataGridView dataGridView, string fileName)
        {
            string[] lines = System.IO.File.ReadAllLines(fileName);
            if (lines.Length == 0)
            {
                return;
            }
            string[] headers = lines[0].Split(',');
            for (int i = 0; i < headers.Length; i++)
            {
                dataGridView.Columns.Add(headers[i], headers[i]);
            }
            for (int i = 1; i < lines.Length; i++)
            {
                string[] values = lines[i].Split(',');
                dataGridView.Rows.Add(values);
            }
        }

        //DataGridView转为Excel
        public static void DataGridViewToExcel(DataGridView dataGridView, string fileName)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Application.Workbooks.Add(true);
            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                excel.Cells[1, i + 1] = dataGridView.Columns[i].HeaderText;
            }
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView.Columns.Count; j++)
                {
                    if (dataGridView[j, i].Value != null)
                    {
                        excel.Cells[i + 2, j + 1] = dataGridView[j, i].Value.ToString();

                    }
                    else
                    {
                        excel.Cells[i + 2, j + 1] = "";
                    }
                }
            }
            excel.Visible = false;
            excel.ActiveWorkbook.SaveAs(fileName);
            excel.Quit();
        }
    }
}
