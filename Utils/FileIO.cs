using Microsoft.EntityFrameworkCore.Internal;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using 考核系统.Dialogs;
using 考核系统.Entity;
using 考核系统.Mapper;
using 考核系统.Utils.CalcUtils;
namespace 考核系统.Utils
{
    internal class FileIO
    {
        private static string Num2Column(int num)
        {
            string result = "";
            while (num > 0)
            {
                int remainder = num % 26;
                if (remainder == 0)
                {
                    remainder = 26;
                    num -= 1;
                }
                result = (char)('A' + remainder - 1) + result;
                num = num / 26;
            }
            return result;
            //如果列数大于26，那么就是26的倍数，需要进位
        }
        public static Dictionary<string,DataGridView> ImportMultiSheets(string fileName)
        {
            var dict=new Dictionary<string, DataGridView>();

            
            var excel = new Microsoft.Office.Interop.Excel.Application();
            excel.DisplayAlerts = false;

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
            excel.DisplayAlerts = false;
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
            var range_string = "A1:" + Num2Column(dataGridView.Columns.Count) + "1";

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
            excel.DisplayAlerts = false;
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

                var range_string = "A1:" + Num2Column(dataGridView.Columns.Count) + "1";
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
            
            
            var deptsRaw = CommonData.DeptInfo.Values;
            //depts按照dept_code排序，使用NaturalComparer
            var depts = deptsRaw.OrderBy(dept => dept.Item1.dept_code, new NaturalComparer()).ToList();


            var excel = new Microsoft.Office.Interop.Excel.Application();
            excel.DisplayAlerts = false;
            var completions = CommonData.CompletionInfo;
            //每个指标分组一个sheet，一个sheet包含该指标组下的所有指标



            excel.Visible = false;


            excel.Application.Workbooks.Add(true);
            var workbook = excel.Application.Workbooks.Item[1];
            var idx = 1;
            var currentYear=CommonData.CurrentYear;
            var info_row_idx = 4;
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

                var range_string = "A1:" + Num2Column(tableMergeWidth) + "1";
                var range = worksheet.Range[range_string];
                range.Merge();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                range.Font.Size = 26;
                range.Font.Bold = true;

                range.Borders.LineStyle = 1;
                range.value2 = header;
                


                range=worksheet.Range["A2:B4"];//左上角
                range.Merge();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                range.Borders.LineStyle = 1;
                range.value2 = "教学科研单位";
                range.Font.Size = 16;
                range.Font.Bold = true;
                
                sheet.Cells[info_row_idx, 1].Value = "单位代码";
                sheet.Cells[info_row_idx, 1].Font.Bold = true;
                sheet.Cells[info_row_idx, 1].Font.Size = 14;
                sheet.Cells[info_row_idx, 1].Borders.LineStyle = 1;
                sheet.Cells[info_row_idx, 2].Value = "单位名称";
                sheet.Cells[info_row_idx, 2].Font.Size = 14;
                sheet.Cells[info_row_idx, 2].Font.Bold = true;
                sheet.Cells[info_row_idx, 2].Borders.LineStyle = 1;
                int dept_idx = 5;
                foreach(var dept in depts)
                {
                    sheet.Cells[dept_idx, 1].Value = dept.Item1.dept_code;
                    sheet.Cells[dept_idx, 1].Font.Bold = true;
                    sheet.Cells[dept_idx, 1].Borders.LineStyle = 1;
                    sheet.Cells[dept_idx, 2].Value = dept.Item1.dept_name;
                    sheet.Cells[dept_idx, 2].Font.Bold = true;
                    sheet.Cells[dept_idx, 2].Borders.LineStyle = 1;
                    dept_idx++;
                }
                dept_idx--;

                int index_col_idx = 3;//指标列的起始列
                HashSet<Tuple<int,int>> processedGroups = new HashSet<Tuple<int, int>>();//如果是同一组的指标，只需要写一次
                foreach (var index in currentIndexes)
                {
                    if (processedGroups.Contains(new Tuple<int, int>(index.identifier_id, index.secondary_identifier)))
                    {
                        continue;//如果是同一组的指标，只需要写一次
                    }


                    var same_group= currentIndexes.Where(i => i.identifier_id == index.identifier_id && i.secondary_identifier == index.secondary_identifier).ToList();

                    if (same_group.Any(i => i.tertiary_identifier == Consts.mainIndexPlaceholder))
                    {
                        //说明有总指标，需要合并单元格
                        string main_string = index.identifier_id.ToString() +
                            "." + index.secondary_identifier.ToString() +
                            "::" + same_group.FirstOrDefault(i => i.tertiary_identifier == Consts.mainIndexPlaceholder).index_name;
                        //分隔符是::,读取时可以用split判断是否是总指标
                        string main_col_left = Num2Column(index_col_idx);
                        string main_col_right = Num2Column(index_col_idx + same_group.Count * 2 - 1);//每个子指标占两列，包括总指标也占两列
                        range = worksheet.Range[main_col_left + "2:" + main_col_right + "2"];
                        
                        range.Merge();
                        range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        range.Borders.LineStyle = 1;
                        range.value2 = main_string;
                        range.Font.Bold = true;
                        range.Font.Size = 20;
                        range.WrapText = true;
                        //先导出总指标
                        var main_index = same_group.FirstOrDefault(i => i.tertiary_identifier == Consts.mainIndexPlaceholder);
                        ExportSingleIndex(ref index_col_idx, range, worksheet, sheet, main_index, info_row_idx, true);//导出单个指标
                        same_group.Remove(main_index);
                        foreach (var group_idx in same_group)
                        {
                            ExportSingleIndex(ref index_col_idx, range, worksheet, sheet, group_idx,info_row_idx ,true);//导出单个指标
                        }
                        processedGroups.Add(new Tuple<int, int>(index.identifier_id, index.secondary_identifier));
                    }
                    else
                    {
                        ExportSingleIndex(ref index_col_idx, range, worksheet, sheet, index, info_row_idx, false);//导出单个指标
                    }
                    //processedGroups.Add(new Tuple<int, int>(index.identifier_id, index.secondary_identifier));
                }
                //调整表格行、列宽，自适应
                range = worksheet.Range["A4:" + Num2Column(tableMergeWidth) + (dept_idx).ToString()];
                range.Columns.AutoFit();
                range.Rows.AutoFit();

                idx++;
            }
            try
            {
                excel.ActiveWorkbook.SaveAs(fileName);
            }
            catch(Exception e)
            {
                
            }
            finally
            {
                excel.Quit();
            }
        }
        private static void ExportSingleIndex(ref int index_col_idx,dynamic range,dynamic worksheet,dynamic sheet,Entity.Index index,int info_row_idx, bool subIndex = false)
        {
            int currentYear= CommonData.CurrentYear;
            sheet.Cells[info_row_idx, index_col_idx].Value = currentYear.ToString() + "年目标";
            sheet.Cells[info_row_idx, index_col_idx].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            sheet.Cells[info_row_idx, index_col_idx].Font.Bold = true;
            sheet.Cells[info_row_idx, index_col_idx].Font.size = 13;
            sheet.Cells[info_row_idx, index_col_idx].Borders.LineStyle = 1;
            sheet.Cells[info_row_idx, index_col_idx + 1].Value = currentYear.ToString() + "年完成";
            sheet.Cells[info_row_idx, index_col_idx + 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            sheet.Cells[info_row_idx, index_col_idx + 1].Font.Bold = true;
            sheet.Cells[info_row_idx, index_col_idx + 1].Font.size = 13;
            sheet.Cells[info_row_idx, index_col_idx + 1].Borders.LineStyle = 1;
            string col_char1 = Num2Column(index_col_idx);
            string col_char2 = Num2Column(index_col_idx + 1);
            range = worksheet.Range[col_char1 + (subIndex?"3":"2")+":" + col_char2 + "3"];//todo
            range.Merge();
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Borders.LineStyle = 1;
            string valueString = index.identifier_id.ToString() + "." + index.secondary_identifier.ToString();
            valueString += index.tertiary_identifier == Consts.singleIndexPlaceholder ? "" : ".#" + index.id;//三级指标
            valueString += ":" + index.index_name;
            if (index.tertiary_identifier == Consts.mainIndexPlaceholder)
            {
                valueString = "总计";
            }
            range.value2 = valueString;
            range.Font.Bold = true;
            range.Font.Size = 12;
            range.WrapText = true;

            //range.Rows.RowHeight = range.Rows.RowHeight * 1.5;
            //range.Columns.ColumnWidth = range.Columns.ColumnWidth * 1.5;
            int dept_idx2 = 5;
            var deptsRaw = CommonData.DeptInfo.Values;

            //depts按照dept_code排序，使用NaturalComparer
            var depts = deptsRaw.OrderBy(dept => dept.Item1.dept_code, new NaturalComparer()).ToList();

            var completions = CommonData.CompletionInfo;
            var group_completions = CommonData.GroupCompletionInfo;
            var visitedDepts=new HashSet<int>();
            var groupMapper = GroupsMapper.GetInstance();
            int dept_iter_idx = 0;
            foreach (var dept in depts)
            {
                if(visitedDepts.Contains(dept.Item1.id))
                {
                    dept_iter_idx++;
                    continue;
                }
                var group=groupMapper.GetGroupByDeptCode(dept.Item1.dept_code, index.id);
                if (group!=null)
                {
                    var cur_groupCompletion= group_completions.Values.FirstOrDefault(completion => completion.group_id == group.id);
                    int group_start_row_idx = dept_idx2;
                    int group_end_row_idx = dept_idx2;
                    if (cur_groupCompletion != null)
                    {
                        int cnt = 0;
                        int search_idx = dept_iter_idx;
                        var naturalComparer = new NaturalComparer();
                        while (
                            search_idx<depts.Count
                            &&
                            naturalComparer.Between(group.l_bound, group.r_bound, depts[search_idx++].Item1.dept_code)
                            )
                        {
                            visitedDepts.Add(depts[search_idx - 1].Item1.id);
                            ++cnt;
                        }
                        group_end_row_idx = group_start_row_idx + cnt - 1;
                        if(cur_groupCompletion.target != 0)
                        {
                            var group_range = worksheet.Range[col_char1 + group_start_row_idx.ToString() + ":" + col_char1 + group_end_row_idx.ToString()];
                            group_range.Merge();
                            group_range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            group_range.Borders.LineStyle = 1;
                            group_range.Value2 = cur_groupCompletion.target.ToString();

                            group_range = worksheet.Range[col_char2 + group_start_row_idx.ToString() + ":" + col_char2 + group_end_row_idx.ToString()];
                            group_range.Merge();
                            group_range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            group_range.Borders.LineStyle = 1;
                            group_range.Value2 = cur_groupCompletion.completed == 0 ? "" : cur_groupCompletion.completed.ToString();
                        }
                        else
                        {
                            //以三年为周期进行考核
                            var bypass_range = worksheet.Range[col_char1 + group_start_row_idx.ToString() + ":" + col_char2 + group_end_row_idx.ToString()];
                            bypass_range.Merge();
                            bypass_range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            bypass_range.Value2 = "以三年为周期进行考核";
                            bypass_range.Borders.LineStyle = 1;
                            bypass_range.Font.Bold = true;
                            //bypass_range.Font.Italic = true;

                        }
                        dept_idx2 = group_end_row_idx + 1;
                    }
                    else
                    {
                        //按理来说不会执行到这里
                        
                    }


                }
                else
                {
                    var cur_completion = completions.Values.FirstOrDefault(completion => completion.dept_id == dept.Item1.id && completion.index_id == index.id);
                    if (cur_completion == null) { continue; }

                    if (cur_completion.target != 0)
                    {
                        sheet.Cells[dept_idx2, index_col_idx].Value = cur_completion.target;
                        sheet.Cells[dept_idx2, index_col_idx].Borders.LineStyle = 1;
                        sheet.Cells[dept_idx2, index_col_idx + 1].Value = cur_completion.completed == 0 ? "" : cur_completion.completed.ToString();
                        sheet.Cells[dept_idx2, index_col_idx + 1].Borders.LineStyle = 1;
                    }
                    else
                    {
                        var bypass_range = worksheet.Range[col_char1 + dept_idx2.ToString() + ":" + col_char2 + dept_idx2.ToString()];
                        bypass_range.Merge();
                        bypass_range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        bypass_range.Value2 = "以三年为周期进行考核";
                        bypass_range.Font.Bold = true;
                        //bypass_range.Font.Italic = true;
                        bypass_range.Borders.LineStyle = 1;
                    }

                    dept_idx2++;
                    visitedDepts.Add(dept.Item1.id);
                }
                dept_iter_idx++;
            }
            index_col_idx += 2;
        }

        public static void ImportCompletionTable(string fileName,ImportCallback callback, ref string errorInfo)
        {
            var excel = new Microsoft.Office.Interop.Excel.Application();
            excel.DisplayAlerts = false;
            excel.Visible = false;
            var workbook = excel.Workbooks.Open(fileName);
            var sheet_count = workbook.Sheets.Count;
            var completions = CommonData.CompletionInfo;
            var completionMapper = CompletionMapper.GetInstance();
            var currentYear = CommonData.CurrentYear;
            var deptsRaw = CommonData.DeptInfo.Values;
            var depts= deptsRaw.OrderBy(dept => dept.Item1.dept_code, new NaturalComparer()).ToList();
            int progress = 10;
            callback("读取元数据", progress);
            int progress_per_sheet = 90 / sheet_count;
            for (int i = 1; i <= sheet_count; i++)
            {
                progress = 10 + i * progress_per_sheet;
                callback("正在导入第" + i + "个sheet", progress);
                var worksheet = workbook.Sheets.Item[i];
                workbook.Sheets.Item[i].Select();
                var currentDepts=new List<Department>();

                int row_idx = 5;
                callback("正在导入第" + i + "个sheet:读取单位信息", progress);
                while (true)
                {
                    var dept_code = worksheet.Cells[row_idx, 1].Value;
                    if (dept_code == null) break;
                    var department = depts.FirstOrDefault(dept => dept.Item1.dept_code == dept_code.ToString());
                    if (department != null)
                    {
                        currentDepts.Add(department.Item1);//由于导出的时候是按照dept_code排序的，所以这里也是按照dept_code排序的
                    }
                    else
                    {

                    }
                    row_idx++;
                }//读取单位信息
                int index_col_idx = 3;

                while (true)
                {
                    var index_name_raw = worksheet.Cells[2, index_col_idx].Value;
                    if (index_name_raw == null) break;
                    callback("正在导入第" + i + "个sheet:读取指标"+index_name_raw, progress);
                    if (((string)index_name_raw).Contains("::")){
                        
                        string code = index_name_raw.ToString().Split("::".ToCharArray())[0];



                        var main_index = CommonData.IndexInfo.Values.FirstOrDefault
                            (idx => idx.identifier_id == Convert.ToInt32(code.Split('.')[0])
                            && idx.secondary_identifier == Convert.ToInt32(code.Split('.')[1])
                            && idx.tertiary_identifier == Consts.mainIndexPlaceholder//todo:修改 11.19
                            );
                        int group_cnt= CommonData.IndexInfo.Values.Count(idx => idx.identifier_id == main_index.identifier_id && idx.secondary_identifier == main_index.secondary_identifier);
                        for(i=0; i < group_cnt; i++)
                        {
                            index_name_raw = worksheet.Cells[3, index_col_idx].Value;
                            if(index_name_raw==null)continue;//如果是空的，说明可能有别的子指标在其他职能部门
                            ImportSingleIndex(index_name_raw, currentDepts, worksheet, ref index_col_idx,ref errorInfo, main_index);
                        }
                        
                    }
                    else
                    {
                        ImportSingleIndex(index_name_raw, currentDepts, worksheet, ref index_col_idx, ref errorInfo);
                    }

                }



            }
            workbook.Close();
            excel.Quit();
            if(errorInfo != "")
            {
                errorInfo = $"{fileName}有以下错误：\r\n" + errorInfo;
            }
            callback(fileName + ":" + "导入完成", 100);

        }
        private static void ImportSingleIndex(string index_name_raw, List<Department> currentDepts,dynamic worksheet,ref int index_col_idx,ref string errorInfo,Entity.Index importedMainIndex=null)
        {
            var currentYear = CommonData.CurrentYear;
            var completions = CommonData.CompletionInfo;
            var completionMapper = CompletionMapper.GetInstance();
            var groupCompletionMapper=GroupCompletionMapper.GetInstance();
            Entity.Index index = null;
            if (index_name_raw!="总计")
            {
                string code = index_name_raw.ToString().Split(':')[0];
                int identifier_id = Convert.ToInt32(code.Split('.')[0]);
                int secondary_identifier = Convert.ToInt32(code.Split('.')[1]);
                string tertiary_identifier = code.Split('.').Length <= 2 ? "0" : code.Split('.')[2];//三级指标
                if (tertiary_identifier.StartsWith("#"))
                {
                    tertiary_identifier=tertiary_identifier.Substring(1);//#后面的是index_id 
                    index = CommonData.IndexInfo.Values.FirstOrDefault
                        (idx => idx.identifier_id == identifier_id
                        && idx.secondary_identifier == secondary_identifier
                        && idx.id == Convert.ToInt32(tertiary_identifier)
                        );
                }
                else
                {
                    index = CommonData.IndexInfo.Values.FirstOrDefault
                        (idx => idx.identifier_id == identifier_id
                        && idx.secondary_identifier == secondary_identifier
                        );
                }




            }
            else
            {
                index = importedMainIndex;
            }
            int value_row_offset = 5;
            var groupsMapper = GroupsMapper.GetInstance();
            if (index != null)
            {
                for (int j = 0; j < currentDepts.Count; j++)
                {
                    //要判断是单独的，还是组的
                    var group= groupsMapper.GetGroupByDeptCode(currentDepts[j].dept_code, index.id);
                    if (group != null)
                    {
                        var cur_group_completion = new GroupCompletion();
                        cur_group_completion.group_id = group.id;
                        cur_group_completion.year = currentYear;
                        int delta_j = 0;

                        while(j+delta_j< currentDepts.Count && groupsMapper.GetGroupByDeptCode(currentDepts[j + delta_j].dept_code, index.id).id == group.id)
                        {
                            delta_j++;
                        }
                        delta_j--;
                        //delta_j代表某个组里面有多少个部门-1 , -1是因为下一轮循环会加1

                        if (worksheet.Cells[j + value_row_offset, index_col_idx].Value.ToString().Contains("以三年为周期进行考核")|| worksheet.Cells[j + value_row_offset, index_col_idx].Value.ToString()=="")
                        {
                            j += delta_j;
                            continue;
                        }

                        try
                        {
                            Convert.ToInt32(worksheet.Cells[j + value_row_offset, index_col_idx].Value);
                        }
                        catch (Exception)
                        {
                            string errInfo=$"表格中的 {currentYear} 年 {currentDepts[j].dept_name} 的 {index.index_name} 的目标数不合法！";
                            MessageBox.Show(errInfo, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            errorInfo+= errInfo+ "\r\n";
                            j += delta_j;
                            continue;
                        }

                        try
                        {
                            Convert.ToInt32(worksheet.Cells[j + value_row_offset, index_col_idx + 1].Value);
                        }
                        catch (Exception)
                        {
                            string errInfo = $"表格中的 {currentYear} 年 {currentDepts[j].dept_name} 的 {index.index_name} 的完成数不合法！";
                            MessageBox.Show(errInfo, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            errorInfo += errInfo + "\r\n";
                            j += delta_j;
                            continue;
                        }


                        cur_group_completion.target = worksheet.Cells[j + value_row_offset, index_col_idx].Value == null ? 0 : Convert.ToInt32(worksheet.Cells[j + value_row_offset, index_col_idx].Value);
                        cur_group_completion.completed = worksheet.Cells[j + value_row_offset, index_col_idx + 1].Value == null ? 0 : Convert.ToInt32(worksheet.Cells[j + value_row_offset, index_col_idx + 1].Value);
                        cur_group_completion.index_id = index.id;


                        if (cur_group_completion.target != 0)
                        {
                            if (CommonData.GroupCompletionInfo.Any
                                (com => com.Value.group_id == cur_group_completion.group_id &&
                            com.Value.year == cur_group_completion.year &&
                            com.Value.index_id==cur_group_completion.index_id))
                            {
                                var group_completion = CommonData.GroupCompletionInfo.Values.FirstOrDefault(com => com.group_id == cur_group_completion.group_id && com.year == cur_group_completion.year);
                                cur_group_completion.id = group_completion.id;
                                groupCompletionMapper.Update(cur_group_completion);
                                Logger.Log($"在{currentYear}年，部门{currentDepts[j].dept_name}的{index.index_name}的完成数为{cur_group_completion.completed}");
                            }
                            else
                            {
                                //groupCompletionMapper.Add(cur_group_completion);
                                //CommonData.GroupCompletionInfo[cur_group_completion.id] = cur_group_completion;
                                //按理来说不会执行到这里，因为数据库里肯定已经有了
                            }
                        }
                        j += delta_j;

                    }
                    else
                    {
                        var cur_completion = new Completion();
                        cur_completion.dept_id = currentDepts[j].id;
                        cur_completion.index_id = index.id;

                        if (worksheet.Cells[j + value_row_offset, index_col_idx].Value.ToString().Contains("以三年为周期进行考核")|| worksheet.Cells[j + value_row_offset, index_col_idx].Value.ToString()=="")
                        {
                            continue;//跳过
                        }


                        try
                        {
                            Convert.ToInt32(worksheet.Cells[j + value_row_offset, index_col_idx].Value);
                        }
                        catch (Exception)
                        {
                            string errInfo = $"表格中的 {currentYear} 年 {currentDepts[j].dept_name} 的 {index.index_name} 的目标数不合法！";
                            errorInfo += errInfo + "\r\n";
                            MessageBox.Show(errInfo, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            continue;
                            
                        }

                        try
                        {
                            Convert.ToInt32(worksheet.Cells[j + value_row_offset, index_col_idx + 1].Value);
                        }
                        catch (Exception)
                        {
                            string errInfo = $"表格中的 {currentYear} 年 {currentDepts[j].dept_name} 的 {index.index_name} 的完成数不合法！";
                            errorInfo += errInfo + "\r\n";
                            MessageBox.Show(errInfo, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            continue;

                        }


                        cur_completion.target = worksheet.Cells[j + value_row_offset, index_col_idx].Value == null ? 0 : Convert.ToInt32(worksheet.Cells[j + value_row_offset, index_col_idx].Value);
                        cur_completion.completed = worksheet.Cells[j + value_row_offset, index_col_idx + 1].Value == null ? 0 : Convert.ToInt32(worksheet.Cells[j + value_row_offset, index_col_idx + 1].Value);
                        cur_completion.year = currentYear;
                        if (
                            completions.Any(com => com.Value.dept_id == cur_completion.dept_id &&
                                        com.Value.index_id == cur_completion.index_id &&
                                        com.Value.year == cur_completion.year)
                            )
                        {
                            var completion = completions.Values.FirstOrDefault(com => com.dept_id == cur_completion.dept_id &&
                                        com.index_id == cur_completion.index_id &&
                                        com.year == cur_completion.year);
                            cur_completion.id = completion.id;
                            cur_completion.target = completion.target;

                            completionMapper.Update(cur_completion);
                            Logger.Log($"在{currentYear}年，部门{currentDepts[j].dept_name}的{index.index_name}的完成数为{cur_completion.completed}");
                        }
                        else
                        {
                            //completionMapper.Add(cur_completion);//按理来说不会执行到这里，因为数据库里肯定已经有了
                            //CommonData.CompletionInfo[cur_completion.id] = cur_completion;
                        }
                    }

                }
            }
            else
            {
                string err = $"导入的指标中有不存在的指标:{index_name_raw}，请检查!";
                var result = MessageBox.Show(err
                , "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                errorInfo += err + "\r\n";
            }
            index_col_idx += 2;
        }
        public delegate void ExportCallback(string message, int progress);
        public delegate void ImportCallback(string message, int progress);
        public static void ExportMain(string filename,ExportCallback exportCallback)
        {
            var deptFinalScores=new Dictionary<int,DeptFinalScore>();
            var globalDeptInfo = GlobalDeptInfo.GetInstance();
            var deptInfo = CommonData.DeptInfo.Values;
            
            exportCallback("初始化Excel...", 0);
            var currentYear = CommonData.CurrentYear;
            var excel= new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.Application.Workbooks.Add(true);
            var workbook = excel.Application.Workbooks.Item[1];
            var worksheet = workbook.Worksheets.Item[1];
            worksheet.Name = "汇总计算";

            var range = worksheet.Range["A1:F1"];
            range.Merge();
            range.Value2 = "指标信息";
            range.Font.Size = 20;
            range.Font.Bold = true;
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            int property_row_idx = 2;//如果后面需要更改，可以直接改这个值，不用一个一个改
            int property_col_idx_start = 1;



            exportCallback("导出元数据", 5);


            int property_col_idx = property_col_idx_start;
            worksheet.Cells[property_row_idx, property_col_idx].Value = "指标名称";
            worksheet.Cells[property_row_idx, property_col_idx].Font.Bold = true;
            worksheet.Cells[property_row_idx, property_col_idx].Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            property_col_idx++;
            worksheet.Cells[property_row_idx, property_col_idx].Value = "指标类型";
            worksheet.Cells[property_row_idx, property_col_idx].Font.Bold = true;
            worksheet.Cells[property_row_idx, property_col_idx].Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            property_col_idx++;
            worksheet.Cells[property_row_idx, property_col_idx].Value = "一级权重";
            worksheet.Cells[property_row_idx, property_col_idx].Font.Bold = true;
            worksheet.Cells[property_row_idx, property_col_idx].Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            property_col_idx++;
            worksheet.Cells[property_row_idx, property_col_idx].Value = "二级权重";
            worksheet.Cells[property_row_idx, property_col_idx].Font.Bold = true;
            worksheet.Cells[property_row_idx, property_col_idx].Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            property_col_idx++;
            worksheet.Cells[property_row_idx, property_col_idx].Value = "敏感性指标权重";
            worksheet.Cells[property_row_idx, property_col_idx].Font.Bold = true;
            worksheet.Cells[property_row_idx, property_col_idx].Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            property_col_idx++;
            worksheet.Cells[property_row_idx, property_col_idx].Value = currentYear.ToString() + "年全校总完成值";
            worksheet.Cells[property_row_idx, property_col_idx].Font.Bold = true;
            worksheet.Cells[property_row_idx, property_col_idx].Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            property_col_idx++;

            int row_idx = property_row_idx + 1;
            var indexes = CommonData.IndexInfo.Values;
            var globalIndexInfo = GlobalIndexInfo.GetInstance();
            exportCallback("写入指标元数据", 10);
            foreach (var index in indexes)
            {
                var nameString = index.identifier_id.ToString() + "." + index.secondary_identifier.ToString();
                nameString += (index.tertiary_identifier == Consts.singleIndexPlaceholder || index.tertiary_identifier == Consts.mainIndexPlaceholder)
                    ? "" : ".#" + index.identifier_id;//三级指标
                nameString += ":" + index.index_name;
                property_col_idx = property_col_idx_start;
                worksheet.Cells[row_idx, property_col_idx].Value = nameString;
                worksheet.Cells[row_idx, property_col_idx].Font.Bold = true;
                worksheet.Columns[property_col_idx].ColumnWidth = 60;
                property_col_idx++;
                worksheet.Cells[row_idx, property_col_idx++].Value = index.index_type;
                worksheet.Cells[row_idx, property_col_idx++].Value = index.weight1;
                worksheet.Cells[row_idx, property_col_idx++].Value = index.weight2;
                worksheet.Cells[row_idx, property_col_idx++].Value = index.sensitivity;
                worksheet.Cells[row_idx, property_col_idx++].Value = globalIndexInfo.GlobalCompletion(index);
                row_idx++;
            }
            worksheet.Cells[row_idx, property_col_idx_start].Value = "总分";
            worksheet.Cells[row_idx, property_col_idx_start].Font.Bold = true;
            worksheet.Cells[row_idx, property_col_idx_start].Interior.Color = ColorTranslator.ToOle(Color.LightGreen);


            int dept_col_idx_start = property_col_idx;
            int dept_col_idx = dept_col_idx_start;
            int dept_col_width = 8;//目标，完成，完成率，1,2,3,4,5



            double progress_per_dept = 90 / deptInfo.Count;
            double progress = 10;
            double progress_per_index = progress_per_dept / indexes.Count;
            int dept_idx = 0;
            foreach (var dept in deptInfo)
            {
                
                var deptAnnualInfo = dept.Item2;
                var l_col_idx=Num2Column(dept_col_idx);
                var r_col_idx = Num2Column(dept_col_idx + dept_col_width - 1);
                range = worksheet.Range[l_col_idx + "1:" + r_col_idx + "1"];
                range.Merge();
                range.Value2 = dept.Item1.dept_name+$"({deptAnnualInfo.dept_population}人)";
                range.Font.Size = 16;
                range.Font.Bold = true;
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                worksheet.Cells[property_row_idx, dept_col_idx].Value = $"{deptAnnualInfo.year}年目标值";
                worksheet.Cells[property_row_idx, dept_col_idx].WrapText = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Font.Bold = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                dept_col_idx++;

                worksheet.Cells[property_row_idx, dept_col_idx].Value = $"{deptAnnualInfo.year}年完成";
                worksheet.Cells[property_row_idx, dept_col_idx].WrapText = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Font.Bold = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                dept_col_idx++;

                worksheet.Cells[property_row_idx, dept_col_idx].Value = $"{deptAnnualInfo.year}年完成率";
                worksheet.Cells[property_row_idx, dept_col_idx].WrapText = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Font.Bold = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                dept_col_idx++;
                //基础类完成度得分
                worksheet.Cells[property_row_idx, dept_col_idx].Value = "基础类完成度得分";
                worksheet.Cells[property_row_idx, dept_col_idx].WrapText = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Font.Bold = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Interior.Color = ColorTranslator.ToOle(Color.LightGreen);
                dept_col_idx++;
                //加分类完成度得分
                worksheet.Cells[property_row_idx, dept_col_idx].Value = "加分类完成度得分";
                worksheet.Cells[property_row_idx, dept_col_idx].WrapText = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Font.Bold = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Interior.Color = ColorTranslator.ToOle(Color.Orange);
                    dept_col_idx++;
                //基础类人均贡献度得分
                worksheet.Cells[property_row_idx, dept_col_idx].Value = "基础类人均贡献度得分";
                worksheet.Cells[property_row_idx, dept_col_idx].WrapText = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Font.Bold = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Interior.Color = ColorTranslator.ToOle(Color.MediumPurple);
                dept_col_idx++;
                //敏感性指标得分
                worksheet.Cells[property_row_idx, dept_col_idx].Value = "敏感性指标得分";
                worksheet.Cells[property_row_idx, dept_col_idx].WrapText = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Font.Bold = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                dept_col_idx++;
                //基础类指标完成度理论满分
                worksheet.Cells[property_row_idx, dept_col_idx].Value = "基础类指标完成度理论满分";
                worksheet.Cells[property_row_idx, dept_col_idx].WrapText = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Font.Bold = true;
                worksheet.Cells[property_row_idx, dept_col_idx].Interior.Color = ColorTranslator.ToOle(Color.LightBlue);
                dept_col_idx++;

                row_idx = property_row_idx + 1;//从表头下一行开始
                double BasicCompletionScoreSum = 0;
                double BonusCompletionScoreSum = 0;
                double BasicCompletionScorePerCapitaSum = 0;
                double SensitivityScoreSum = 0;
                double BasicTheoreticalFullScoreSum = 0;

                
                foreach (var index in indexes)
                {
                    progress += progress_per_index;
                    exportCallback($"正在导出{dept.Item1.dept_name}的{index.index_name}", (int)progress);
                    dept_col_idx = dept_col_idx_start;
                    var completion = CommonData.CompletionInfo.Values.FirstOrDefault
                        (com => com.dept_id == dept.Item1.id &&
                    com.index_id == index.id &&
                    com.year==deptAnnualInfo.year);
                    var group= GroupsMapper.GetInstance().GetGroupByDeptCode(dept.Item1.dept_code, index.id);
                    
                    if (group != null)
                    {
                        var group_completion=CommonData.GroupCompletionInfo.Values.FirstOrDefault
                            (com => com.group_id == group.id &&
                            com.year == deptAnnualInfo.year &&
                            com.index_id == index.id);
                        if (group_completion != null)
                        {
                            if (group_completion.target != 0)
                            {
                                completion.target = group_completion.target;
                                completion.completed = group_completion.completed;
                            }
                        }
                    }

                    if (completion != null)
                    {
                        var calcUnit = new CalcUnit(index, dept.Item1, deptAnnualInfo, completion);

                        if (completion.target != 0)
                            worksheet.Cells[row_idx, dept_col_idx].Value = completion.target;
                        dept_col_idx++;

                        if(completion.completed != 0)
                            worksheet.Cells[row_idx, dept_col_idx].Value = completion.completed;
                        dept_col_idx++;

                        if (completion.completion_rate != 0)
                            worksheet.Cells[row_idx, dept_col_idx].Value = completion.completion_rate;
                        
                        dept_col_idx++;
                        //todo 等于0的话，不显示，显示空

                        if (calcUnit.BasicCompletionScore != 0)
                            worksheet.Cells[row_idx, dept_col_idx].Value = calcUnit.BasicCompletionScore;
                        BasicCompletionScoreSum += calcUnit.BasicCompletionScore;
                        dept_col_idx++;

                        if (calcUnit.BonusCompletionScore != 0)
                            worksheet.Cells[row_idx, dept_col_idx].Value = calcUnit.BonusCompletionScore;
                        BonusCompletionScoreSum += calcUnit.BonusCompletionScore;
                        dept_col_idx++;

                        if (calcUnit.BasicCompletionScorePerCapita != 0)
                            worksheet.Cells[row_idx, dept_col_idx].Value = calcUnit.BasicCompletionScorePerCapita;
                        BasicCompletionScorePerCapitaSum += calcUnit.BasicCompletionScorePerCapita;
                        dept_col_idx++;

                        if (calcUnit.SensitivityScore != 0)
                            worksheet.Cells[row_idx, dept_col_idx].Value = calcUnit.SensitivityScore;
                        SensitivityScoreSum += calcUnit.SensitivityScore;
                        dept_col_idx++;

                        if (calcUnit.BasicTheoreticalFullScore != 0)
                            worksheet.Cells[row_idx, dept_col_idx].Value = calcUnit.BasicTheoreticalFullScore;
                        BasicTheoreticalFullScoreSum += calcUnit.BasicTheoreticalFullScore;
                        dept_col_idx++;

                        row_idx++;
                    }
                }
                dept_col_idx = dept_col_idx_start + 3;//从基础类完成度得分开始，只要后5列

                worksheet.Cells[row_idx, dept_col_idx].Value = BasicCompletionScoreSum;
                worksheet.Cells[row_idx, dept_col_idx].Font.Bold = true;
                worksheet.Cells[row_idx, dept_col_idx].Interior.Color = ColorTranslator.ToOle(Color.LightGreen);
                dept_col_idx++;

                worksheet.Cells[row_idx, dept_col_idx].Value = BonusCompletionScoreSum;
                worksheet.Cells[row_idx, dept_col_idx].Font.Bold = true;
                worksheet.Cells[row_idx, dept_col_idx].Interior.Color = ColorTranslator.ToOle(Color.Orange);
                dept_col_idx++;

                worksheet.Cells[row_idx, dept_col_idx].Value = BasicCompletionScorePerCapitaSum;
                worksheet.Cells[row_idx, dept_col_idx].Font.Bold = true;
                worksheet.Cells[row_idx, dept_col_idx].Interior.Color = ColorTranslator.ToOle(Color.MediumPurple);
                dept_col_idx++;

                worksheet.Cells[row_idx, dept_col_idx].Value = SensitivityScoreSum;
                worksheet.Cells[row_idx, dept_col_idx].Font.Bold = true;
                worksheet.Cells[row_idx, dept_col_idx].Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                dept_col_idx++;

                worksheet.Cells[row_idx, dept_col_idx].Value = BasicTheoreticalFullScoreSum;
                worksheet.Cells[row_idx, dept_col_idx].Font.Bold = true;
                worksheet.Cells[row_idx, dept_col_idx].Interior.Color = ColorTranslator.ToOle(Color.LightBlue);
                dept_col_idx++;

                deptFinalScores[dept.Item1.id] = new DeptFinalScore
                (
                    dept.Item1,
                    dept.Item2,
                    BasicCompletionScoreSum,
                    BonusCompletionScoreSum,
                    BasicCompletionScorePerCapitaSum,
                    SensitivityScoreSum,
                    BasicTheoreticalFullScoreSum
                );
                dept_col_idx_start += dept_col_width;//下一个单位的起始列
            }
            //新增一个sheet


            excel.Worksheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
            excel.Worksheets.Item[2].Select();
            worksheet = workbook.Worksheets.Item[2];
            worksheet.Name = "客观分统计";
            progress = 90;
            exportCallback("正在导出客观分统计", (int)progress);
            int column_width = 16;//有16列
            string column_start_char = Num2Column(1);
            string column_end_char = Num2Column(column_width);
            range = worksheet.Range[column_start_char + "1:" + column_end_char + "1"];
            range.Merge();
            range.Value2 = "客观分统计";
            range.Font.Size = 22;
            range.Font.Bold = true;
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            int head_row_idx = 2;
            int col_idx = 1;


            range = worksheet.Range[
                Num2Column(col_idx) + head_row_idx.ToString() + ":"
                + Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "序号";
            range.Font.Bold = true;
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            col_idx++;

            range = worksheet.Range[
    Num2Column(col_idx) + head_row_idx.ToString() + ":"
    + Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "部门名称";
            range.Font.Bold = true;
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            col_idx++;

            range = worksheet.Range[
Num2Column(col_idx) + head_row_idx.ToString() + ":"
+ Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "组别";
            range.Font.Bold = true;
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            col_idx++;


            range = worksheet.Range[
Num2Column(col_idx) + head_row_idx.ToString() + ":"
+ Num2Column(col_idx + 1) + (head_row_idx).ToString()];
            range.Merge();
            range.Value2 = "完成度得分";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Bold = true;

            range = worksheet.Range[
Num2Column(col_idx) + (head_row_idx + 1).ToString() + ":"
+ Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "基础类指标";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Bold = true;
            col_idx++;
            range = worksheet.Range[
Num2Column(col_idx) + (head_row_idx + 1).ToString() + ":"
+ Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "加分类指标";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Bold = true;
            col_idx++;

            range = worksheet.Range[
Num2Column(col_idx) + head_row_idx.ToString() + ":"
+ Num2Column(col_idx + 1) + (head_row_idx).ToString()];
            range.Merge();
            range.Value2 = "贡献度得分";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Bold = true;

            range = worksheet.Range[
Num2Column(col_idx) + (head_row_idx + 1).ToString() + ":"
+ Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "基础类指标";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Bold = true;
            col_idx++;
            range = worksheet.Range[
Num2Column(col_idx) + (head_row_idx + 1).ToString() + ":"
+ Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "敏感性指标";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Bold = true;
            col_idx++;
///////////////////////////////////////////
            range = worksheet.Range[
Num2Column(col_idx) + head_row_idx.ToString() + ":"
+ Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "客观分合计";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Bold = true;
            col_idx++;

            range = worksheet.Range[
Num2Column(col_idx) + head_row_idx.ToString() + ":"
+ Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "完成度得分合计\r\n（*25%）";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Color = ColorTranslator.ToOle(Color.Red);
            range.Font.Bold = true;
            col_idx++;

            range = worksheet.Range[
Num2Column(col_idx) + head_row_idx.ToString() + ":"
+ Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "贡献度得分合计\r\n（*25%）";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Color = ColorTranslator.ToOle(Color.Red);
            range.Font.Bold = true;
            col_idx++;

            range = worksheet.Range[
Num2Column(col_idx) + head_row_idx.ToString() + ":"
+ Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "客观分合计（加权后）";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Color = ColorTranslator.ToOle(Color.Red);
            range.Font.Bold = true;
            col_idx++;

            range = worksheet.Range[
Num2Column(col_idx) + head_row_idx.ToString() + ":"
+ Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "客观分归一";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Color = ColorTranslator.ToOle(Color.Red);
            range.Font.Bold = true;
            col_idx++;
            /////////////////////////////////
            range = worksheet.Range[
Num2Column(col_idx) + head_row_idx.ToString() + ":"
+ Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "基础类指标完成度理论满分";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Bold = true;
            col_idx++;
            range = worksheet.Range[
Num2Column(col_idx) + head_row_idx.ToString() + ":"
+ Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "扣分比率";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Bold = true;
            col_idx++;
            range = worksheet.Range[
Num2Column(col_idx) + head_row_idx.ToString() + ":"
+ Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "扣分";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Bold = true;
            col_idx++;
            range = worksheet.Range[
Num2Column(col_idx) + head_row_idx.ToString() + ":"
+ Num2Column(col_idx) + (head_row_idx + 1).ToString()];
            range.Merge();
            range.Value2 = "基础指标得分-扣分";
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.Font.Bold = true;
            col_idx++;
            //宽度自适应
            for (int i = 1; i <= column_width; i++)
            {
                worksheet.Columns[i].AutoFit();
            }
            //给DeptFinalScores排序,按照FinalScore
            var deptFinalScoresList = deptFinalScores.Values.ToList();
            deptFinalScoresList.Sort((x, y) => y.FinalScore.CompareTo(x.FinalScore));
            row_idx = head_row_idx + 2;
            double highestScore = deptFinalScoresList[0].WeightedObjectiveScoreSum;
            double progress_per_dept_final = 10 / deptFinalScoresList.Count;
            //deptFinalScoresList需要按照每个dept的dept_group分组，并对每组后三分之一进行标记
            var groups = deptFinalScoresList.GroupBy(x => x.deptAnnualInfo.dept_group);

            foreach (var group in groups)
            {
                var list= group.ToList();
                for (int i = 0; i < list.Count; i++)
                {
                    progress = 90 + progress_per_dept_final * i;
                    exportCallback($"正在导出{list[i].department.dept_name}的客观分", (int)progress);
                    worksheet.Cells[row_idx, 1].Value = i + 1;
                    worksheet.Cells[row_idx, 2].Value = list[i].department.dept_name;
                    worksheet.Cells[row_idx, 3].Value = list[i].deptAnnualInfo.dept_group;
                    worksheet.Cells[row_idx, 4].Value = list[i].BasicCompletionScoreSum;
                    worksheet.Cells[row_idx, 5].Value = list[i].BonusCompletionScoreSum;
                    worksheet.Cells[row_idx, 6].Value = list[i].BasicCompletionScorePerCapitaSum;
                    worksheet.Cells[row_idx, 7].Value = list[i].SensitivityScoreSum;
                    worksheet.Cells[row_idx, 8].Value = list[i].ObjectiveScoreSum;
                    worksheet.Cells[row_idx, 9].Value = list[i].WeightedCompletionScoreSum;
                    worksheet.Cells[row_idx, 10].Value = list[i].WeightedContributiveScoreSum;
                    worksheet.Cells[row_idx, 11].Value = list[i].WeightedObjectiveScoreSum;
                    worksheet.Cells[row_idx, 12].Value = list[i].NormalizedObjectiveScoreSum(highestScore);
                    worksheet.Cells[row_idx, 13].Value = list[i].BasicTheoreticalFullScoreSum;
                    worksheet.Cells[row_idx, 14].Value = list[i].PunishRate;
                    worksheet.Cells[row_idx, 15].Value = list[i].Punishment;
                    worksheet.Cells[row_idx, 16].Value = list[i].FinalScore;

                    //判断是否是后三分之一（向下取整，例如8个部门，则只取2个）
                    bool isTopThird = i >= list.Count - list.Count / 3;
                    if (isTopThird)
                    {
                        worksheet.Rows[row_idx].Interior.Color = ColorTranslator.ToOle(Color.LightYellow);
                        worksheet.Cells[row_idx, 17].Value = "不能评选B及以上";
                    }
                    row_idx++;
                }


            }


            progress = 100;
            exportCallback($"导出完成", (int)progress);
            excel.Rows[property_row_idx].RowHeight = 30;

            excel.ActiveWorkbook.SaveAs(filename);
            workbook.Close();
            excel.Quit();
            
        }

        
    }
}
