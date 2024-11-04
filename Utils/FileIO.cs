using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
namespace 考核系统.Utils
{
    internal class FileIO
    {
        public static Dictionary<string,DataGridView> ImportMultiSheets(string fileName)
        {
            var dict=new Dictionary<string, DataGridView>();

            
            var excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            var workbook = excel.Workbooks.Open(fileName);
            var sheet_count = workbook.Sheets.Count;

            for(int sheet_idx = 1; sheet_idx <= sheet_count; ++sheet_idx)
            {
                var dataGrid = new DataGridView();
                var worksheet = workbook.Sheets.Item[sheet_idx];
                var range = worksheet.UsedRange;
                var rows = range.Rows.Count;
                var cols = range.Columns.Count;
                //前两行第一行是合并单元格后的标题，不用管，第二行是表头，第三行开始是数据

                for (int i = 1; i <= cols; i++)
                {
                    dataGrid.Columns.Add(range.Cells[2, i].Value.ToString(), range.Cells[2, i].Value.ToString());
                }
                dataGrid.Rows.Clear();
                for (int i = 3; i <= rows; i++)
                {
                    if (range.Cells[i, 1].Value == null) continue;//如果第一列为空，则不写入
                    var row = new List<object>();
                    for (int j = 1; j <= cols; j++)
                    {
                        row.Add(range.Cells[i, j].Value);
                    }
                    dataGrid.Rows.Add(row.ToArray());
                }

                dataGrid.Columns.Insert(0, new DataGridViewTextBoxColumn() { HeaderText = "编号", Name = "编号" });//插入单位编号列,空列

                dict[worksheet.Name] = dataGrid;
            }


            workbook.Close();
            excel.Quit();

            return dict;
        }
        public static DataGridView ImportSingleSheet(string fileName)
        {
            var dataGrid = new DataGridView();
            var excel= new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            var workbook = excel.Workbooks.Open(fileName);
            var worksheet = workbook.Worksheets.Item[1];
            var range = worksheet.UsedRange;
            var rows = range.Rows.Count;
            var cols = range.Columns.Count;
            //前两行第一行是合并单元格后的标题，不用管，第二行是表头，第三行开始是数据
            
            for (int i = 1; i <= cols; i++)
            {
                dataGrid.Columns.Add(range.Cells[2, i].Value.ToString(), range.Cells[2, i].Value.ToString());
            }
            dataGrid.Rows.Clear();
            for (int i = 3; i <= rows; i++)
            {
                if (range.Cells[i, 1].Value == null) continue;//如果第一列为空，则不写入
                var row = new List<object>();
                for (int j = 1; j <= cols; j++)
                {
                    row.Add(range.Cells[i, j].Value);
                }
                dataGrid.Rows.Add(row.ToArray());
            }
            
            dataGrid.Columns.Insert(0, new DataGridViewTextBoxColumn() { HeaderText = "编号", Name = "编号" });//插入单位编号列,空列
            workbook.Close();
            excel.Quit();

            return dataGrid;
        }

        //DataGridView转为Excel

        public static void DataGridViewToExcel(DataGridView dataGridView, string fileName,string header,bool structureOnly=false)
        {
            //第一行为标题，需要合并单元格居中，合并格子数为dataGridView.Columns.Count
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            excel.Application.Workbooks.Add(true);
            excel.Cells[1, 1] = header;
            excel.Cells[1, 1].Font.Size = 20;
            excel.Cells[1, 1].Font.Bold = true;
            excel.Cells[1, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            excel.Cells[1, 1].ColumnWidth = 20;
            excel.Cells[1, 1].RowHeight = 30;
            //合并单元格
            var workbook = excel.Application.Workbooks.Item[1];
            var worksheet =workbook.Worksheets.Item[1];
            worksheet.Name = header;
            var range_string = "A1:" + (char)('A' + dataGridView.Columns.Count - 1) + "1";

            var range = worksheet.Range[range_string];
            range.Merge();
            //range.Value2 = header;
            //合并后的单元格居中
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Borders.LineStyle = 1;
            //写入数据


            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                excel.Cells[2, i + 1] = dataGridView.Columns[i].HeaderText;
                excel.Cells[2, i + 1].Font.Bold = true;
                excel.Cells[2, i + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excel.Cells[2, i + 1].Borders.LineStyle = 1;
            }
            if (!structureOnly)
            {
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView.Columns.Count; j++)
                    {
                        excel.Cells[i + 3, j + 1] = dataGridView.Rows[i].Cells[j].Value;
                        excel.Cells[i + 3, j + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        excel.Cells[i + 3, j + 1].Borders.LineStyle = 1;
                    }
                }
            }
            excel.Columns.AutoFit();
            excel.ActiveWorkbook.SaveAs(fileName);
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

        }
    
        public static void MultiDataGridViewToExcel(Dictionary<string,DataGridView> grids, string fileName, bool structureOnly = false)
        {
            //第一行为标题，需要合并单元格居中，合并格子数为dataGridView.Columns.Count
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            excel.Application.Workbooks.Add(true);
            var Workbook= excel.Application.Workbooks.Item[1];
            var idx=1;
            foreach (var grid in grids)
            {
                if (excel.Worksheets.Count < idx)
                {
                    //在最后添加一个sheet，注意是追加，不是插入
                    excel.Worksheets.Add(After: Workbook.Sheets[Workbook.Sheets.Count]);

                }
                var cnt=excel.Worksheets.Count;
                excel.Worksheets.Item[idx].Select();

                //选择第idx张sheet

                //excel.Sheets.Add();

                var dataGridView = grid.Value;
                var header = grid.Key;
                
                excel.Cells[1, 1] = header;
                excel.Cells[1, 1].Font.Size = 20;
                excel.Cells[1, 1].Font.Bold = true;
                excel.Cells[1, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                excel.Cells[1, 1].ColumnWidth = 20;
                excel.Cells[1, 1].RowHeight = 30;

                //合并单元格
                var worksheet = Workbook.Sheets.Item[idx];
                worksheet.Name = header;

                var range_string = "A1:" + (char)('A' + dataGridView.Columns.Count - 1) + "1";
                var range = worksheet.Range[range_string];
                range.Merge();
                //range.Value2 = header;
                //合并后的单元格居中
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                range.Borders.LineStyle = 1;
                //写入数据
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    excel.Cells[2, i + 1] = dataGridView.Columns[i].HeaderText;
                    excel.Cells[2, i + 1].Font.Bold = true;
                    excel.Cells[2, i + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    excel.Cells[2, i + 1].Borders.LineStyle = 1;
                }

                if (!structureOnly)
                {
                    for (int i = 0; i < dataGridView.Rows.Count; i++)
                    {
                        //如果第一列为空，则不写入
                        if (dataGridView.Rows[i].Cells[0].Value != null)
                        {
                            for (int j = 0; j < dataGridView.Columns.Count; j++)
                            {
                                excel.Cells[i + 3, j + 1] = dataGridView.Rows[i].Cells[j].Value;
                                excel.Cells[i + 3, j + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                                excel.Cells[i + 3, j + 1].Borders.LineStyle = 1;
                            }
                        }
                    }
                }
                idx++;
            }



            excel.Columns.AutoFit();
            excel.ActiveWorkbook.SaveAs(fileName);
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

        }
    }
}
