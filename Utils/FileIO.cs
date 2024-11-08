using Microsoft.EntityFrameworkCore.Internal;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using 考核系统.Entity;
using 考核系统.Mapper;
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
        public static void ExportEmptyCompletionTable(int manager_id,string path)
        {
            //todo:导出空的考核表
            var manager= CommonData.ManagerInfo[manager_id];
            var fileName = path + "\\" + manager.manager_name + "考核表.xlsx";
            var duties= CommonData.DutyInfo.Values.Where(duty => duty.manager_id == manager_id).ToList();
            var indexes_idx=duties.Select(duty=>duty.index_id).Distinct().ToList();
            var indexes = CommonData.IndexInfo.Values.Where(index => indexes_idx.Contains(index.id)).ToList();
            var groupedIndexes=indexes.GroupBy(index => index.identifier_id).ToDictionary(group => group.Key, group => group.ToList());
            var depts = CommonData.DeptInfo.Values;
            var excel = new Microsoft.Office.Interop.Excel.Application();
            var completions = CommonData.CompletionInfo;
            //每个指标分组一个sheet，一个sheet包含该指标组下的所有指标
            excel.Visible = false;
            excel.Application.Workbooks.Add(true);
            var workbook = excel.Application.Workbooks.Item[1];
            var idx = 1;
            var currentYear=CommonData.CurrentYear;
            foreach (var group in groupedIndexes)
            {
                var currentIndexes = group.Value;
                if (excel.Worksheets.Count < idx)
                {
                    //在最后添加一个sheet，注意是追加，不是插入
                    excel.Worksheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
                }
                var cnt = excel.Worksheets.Count;
                excel.Worksheets.Item[idx].Select();
                var sheet=excel.Worksheets.Item[idx];
                var worksheet=excel.Worksheets.Item[idx];
                //选择第idx张sheet
                //excel.Sheets.Add();
                var header = CommonData.IdentifierInfo[group.Key].id+"-"+CommonData.IdentifierInfo[group.Key].identifier_name;
                excel.Worksheets.Item[idx].Name = header;
                var tableMergeWidth = 2 + 2 * currentIndexes.Count;

                var range_string = "A1:" + (char)('A' + tableMergeWidth - 1) + "1";
                var range = worksheet.Range[range_string];
                range.Merge();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                range.Font.Size = 26;
                range.Font.Bold = true;

                range.Borders.LineStyle = 1;
                range.value2 = header;
                


                range=worksheet.Range["A2:B2"];//左上角
                range.Merge();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                range.Borders.LineStyle = 1;
                range.value2 = "教学科研单位";
                range.Font.Size = 16;
                range.Font.Bold = true;
                sheet.Cells[3, 1].Value = "单位代码";
                sheet.Cells[3, 1].Font.Bold = true;
                sheet.Cells[3, 1].Font.Size = 14;
                sheet.Cells[3, 2].Value = "单位名称";
                sheet.Cells[3, 2].Font.Size = 14;
                sheet.Cells[3, 2].Font.Bold = true;
                int dept_idx = 4;
                foreach(var dept in depts)
                {
                    sheet.Cells[dept_idx, 1].Value = dept.Item1.dept_code;
                    sheet.Cells[dept_idx, 1].Font.Bold = true;
                    sheet.Cells[dept_idx, 2].Value = dept.Item1.dept_name;
                    sheet.Cells[dept_idx, 2].Font.Bold = true;
                    dept_idx++;
                }
                dept_idx--;

                int index_col_idx = 3;//指标列的起始列
                foreach (var index in currentIndexes)
                {
                    sheet.Cells[3, index_col_idx].Value = currentYear.ToString() + "年目标";
                    sheet.Cells[3, index_col_idx].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    sheet.Cells[3, index_col_idx].Font.Bold = true;
                    sheet.Cells[3, index_col_idx + 1].Value = currentYear.ToString() + "年完成";
                    sheet.Cells[3, index_col_idx + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    sheet.Cells[3, index_col_idx + 1].Font.Bold = true;


                    string col_char1 = ((char)('A' + index_col_idx - 1)).ToString();
                    string col_char2 = ((char)('A' + index_col_idx - 1 + 1)).ToString();
                    range =worksheet.Range[col_char1+"2:"+ col_char2 + "2"];
                    range.Merge();
                    range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    range.Borders.LineStyle = 1;
                    string valueString= index.identifier_id.ToString() + "." + index.secondary_identifier.ToString();
                    valueString += index.tertiary_identifier == "0" ? "" : "." + index.tertiary_identifier;//三级指标
                    valueString += ":" + index.index_name;
                    range.value2 = valueString;


                    range.Font.Bold = true;
                    range.Font.Size = 12;
                    range.WrapText = true;
                    range.Rows.AutoFit();
                    range.Rows.RowHeight = range.Rows.RowHeight * 1.5;
                    range.Columns.AutoFit();
                    range.Columns.ColumnWidth = range.Columns.ColumnWidth*1.5;
                    int dept_idx2 = 4;
                    foreach (var dept in depts)
                    {
                        var cur_completion= completions.Values.FirstOrDefault(completion => completion.dept_id == dept.Item1.id && completion.index_id == index.id);
                        if(cur_completion==null) { continue; }
                        if (cur_completion.target != 0)
                        {
                            sheet.Cells[dept_idx2, index_col_idx].Value = cur_completion.target;

                            sheet.Cells[dept_idx2, index_col_idx + 1].Value = cur_completion.completed== 0 ? "" : cur_completion.completed.ToString();
                        }
                        else
                        {
                            var bypass_range= worksheet.Range[col_char1 + dept_idx2.ToString() + ":" + col_char2 + dept_idx2.ToString()];
                            bypass_range.Merge();
                            bypass_range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            bypass_range.Value2 = "以三年为周期进行考核";
                            bypass_range.Font.Bold = true;
                            bypass_range.Font.Italic = true;
                        }
                        
                        dept_idx2++;
                    }




                    index_col_idx += 2;
                }
                //调整表格行、列宽，自适应
                range = worksheet.Range["A3:" + (char)('A' + tableMergeWidth - 1) + (dept_idx).ToString()];
                range.Columns.AutoFit();
                range.Rows.AutoFit();

                idx++;
            }

            excel.ActiveWorkbook.SaveAs(fileName);
            excel.Quit();

        }
        public static void ImportCompletionTable(string fileName)
        {
            var excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            var workbook = excel.Workbooks.Open(fileName);
            var sheet_count = workbook.Sheets.Count;
            var completions = CommonData.CompletionInfo;
            var compleitonMapper = CompletionMapper.GetInstance();
            var currentYear = CommonData.CurrentYear;
            var depts = CommonData.DeptInfo.Values;
            for (int i = 1; i <= sheet_count; i++)
            {
                var worksheet = workbook.Sheets.Item[i];
                workbook.Sheets.Item[i].Select();
                var currentDepts=new List<Department>();

                int row_idx = 4;
                while(true)
                {
                    var dept_code = worksheet.Cells[row_idx, 1].Value;
                    if (dept_code == null) break;
                    var department = depts.FirstOrDefault(dept => dept.Item1.dept_code == dept_code.ToString());
                    if (department != null)
                    {
                        currentDepts.Add(department.Item1);
                    }
                    row_idx++;
                }//读取单位信息
                int index_col_idx = 3;

                while (true)
                {
                    var index_name_raw = worksheet.Cells[2, index_col_idx].Value;
                    if (index_name_raw == null) break;
                    string code= index_name_raw.ToString().Split(':')[0];
                    int identifier_id = Convert.ToInt32(code.Split('.')[0]);
                    int secondary_identifier = Convert.ToInt32(code.Split('.')[1]);

                    string tertiary_identifier = code.Split('.').Length <= 2 ? "0" : code.Split('.')[2];//三级指标
                    var index = CommonData.IndexInfo.Values.FirstOrDefault
                        (idx => idx.identifier_id == identifier_id 
                        && idx.secondary_identifier == secondary_identifier
                        && idx.tertiary_identifier == tertiary_identifier
                        );
                    if (index != null)
                    {
                        for (int j = 0; j < currentDepts.Count; j++)
                        {
                            var cur_completion = new Completion();
                            cur_completion.dept_id = currentDepts[j].id;
                            cur_completion.index_id = index.id;
                            

                            cur_completion.target = worksheet.Cells[j + 4, index_col_idx].Value == null ? 0 : Convert.ToDouble(worksheet.Cells[j + 4, index_col_idx].Value);
                            
                            if(worksheet.Cells[j + 4, index_col_idx].Value.ToString().Contains("以三年为周期进行考核"))
                            {
                                continue;//跳过
                            }
                            cur_completion.completed = worksheet.Cells[j + 4, index_col_idx + 1].Value == null ? 0 : Convert.ToInt32(worksheet.Cells[j + 4, index_col_idx + 1].Value);
                            cur_completion.year = currentYear;
                            if (
                                completions.Any(com=> com.Value.dept_id==cur_completion.dept_id&&
                                            com.Value.index_id == cur_completion.index_id &&
                                            com.Value.year == cur_completion.year)
                                )
                            {
                                var completion = completions.Values.FirstOrDefault(com => com.dept_id == cur_completion.dept_id &&
                                            com.index_id == cur_completion.index_id &&
                                            com.year == cur_completion.year);
                                cur_completion.id = completion.id;
                                cur_completion.target = completion.target;

                                compleitonMapper.Update(cur_completion);
                                Logger.Log($"在{currentYear}年，部门{currentDepts[j].dept_name}的{index.index_name}的完成数为{cur_completion.completed}");
                            }
                            else
                            {
                                compleitonMapper.Add(cur_completion);//按理来说不会执行到这里
                            }
                        }
                    }
                    index_col_idx += 2;
                }



            }

        }
    }
}
