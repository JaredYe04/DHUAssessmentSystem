using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using 考核系统.Dialogs;
using 考核系统.Utils;
using 考核系统.Entity;
using 考核系统.Mapper;
using Newtonsoft.Json;
using System.Runtime.CompilerServices;
using static 考核系统.Form1;
namespace 考核系统
{
    public partial class Form1 : Form
    {
        public enum DeptInfoColumns
        {
            id= 0,
            dept_code=1,
            dept_name=2,
            dept_population=3,
            dept_punishment=4,
            dept_group=5
        }//部门信息表的列
        public enum IndexInfoColumns
        {
            id = 0,
            index_code = 1,
            index_name = 2,
            index_type = 3,
            weight1 = 4,
            weight2 = 5,
            enable_sensitivity = 6,
            sensitivity = 7
        }//指标信息表的列
        public enum ManagerInfoColumns
        {
            id = 0,
            manager_code = 1,
            manager_name = 2
        }//职能部门信息表的列
        public Form1()
        {
            InitializeComponent();
            //订阅年份变更事件
            EventBus.YearChanged += EventBus_YearChanged;
        }
        private void EventBus_YearChanged(int year)
        {
            Logger.Log("年份变更为" + year);
            labelCurrentYear.Text = "当前年份:" + year;
            //todo:更新界面
            //labelYear.Text = year.ToString();
        }
        private void fetchDepartmentInfo()
        {
            if(CommonData.DeptInfo== null)
            {
                CommonData.DeptInfo = new Dictionary<int, Tuple<Department, DeptAnnualInfo>>();
            }//初始化部门信息字典

            Logger.Log("开始获取教学科研单位信息");
            //鼠标指针变为等待状态
            Cursor.Current = Cursors.WaitCursor;


            //清空原有数据
            deptDataGrid.Rows.Clear();
            //获取教学科研单位信息
            var deptMapper =DepartmentMapper.GetInstance();
            var deptList = deptMapper.GetAllObjects();
            
            //根据当前年份，获取部门每年的信息
            var year=CommonData.CurrentYear;
            var deptAnnualInfoMapper = DeptAnnualInfoMapper.GetInstance();

            for (int i = 0; i < deptList.Count; i++)
            {
               
                deptDataGrid.Rows.Add();
                deptDataGrid.Rows[i].Cells[(int)DeptInfoColumns.id].Value = deptList[i].id;
                deptDataGrid.Rows[i].Cells[(int)DeptInfoColumns.dept_code].Value = deptList[i].dept_code;
                deptDataGrid.Rows[i].Cells[(int)DeptInfoColumns.dept_name].Value = deptList[i].dept_name;
                var deptAnnualInfo = deptAnnualInfoMapper.GetDeptAnnualInfo(deptList[i].id, year);
                if(deptAnnualInfo == null)
                {

                    deptAnnualInfo = new DeptAnnualInfo(deptList[i].id, year, 0, 0, "1-1");
                    deptAnnualInfoMapper.Add(deptAnnualInfo);
                }
                deptDataGrid.Rows[i].Cells[(int)DeptInfoColumns.dept_population].Value = deptAnnualInfo.dept_population;
                deptDataGrid.Rows[i].Cells[(int)DeptInfoColumns.dept_punishment].Value = deptAnnualInfo.dept_punishment;
                deptDataGrid.Rows[i].Cells[(int)DeptInfoColumns.dept_group].Value = deptAnnualInfo.dept_group;


                CommonData.DeptInfo[deptList[i].id] = new Tuple<Department, DeptAnnualInfo>(Department.Copy(deptList[i]), DeptAnnualInfo.Copy(deptAnnualInfo));
                //后续根据数据表更新数据库时，参照DeptInfo中来判断是否有更新，如果有更新则更新数据库，否则不更新
            }


            //鼠标指针恢复默认状态
            Cursor.Current = Cursors.Default;
            Logger.Log("获取教学科研单位信息成功");
        }
        void fetchIndexInfo()
        {
            if (CommonData.IndexInfo == null)
            {
                CommonData.IndexInfo = new Dictionary<int, Index>();
            }//初始化部门信息字典

            Logger.Log("开始获取指标信息");
            //鼠标指针变为等待状态
            Cursor.Current = Cursors.WaitCursor;


            //清空原有数据
            indexDataGrid.Rows.Clear();
            //获取指标信息
            var indexMapper = IndexMapper.GetInstance();
            var indexList = indexMapper.GetAllObjects();


            for (int i = 0; i < indexList.Count; i++)
            {

                indexDataGrid.Rows.Add();
                indexDataGrid.Rows[i].Cells[(int)IndexInfoColumns.id].Value = indexList[i].id;
                indexDataGrid.Rows[i].Cells[(int)IndexInfoColumns.index_code].Value = indexList[i].index_code;
                indexDataGrid.Rows[i].Cells[(int)IndexInfoColumns.index_name].Value = indexList[i].index_name;
                indexDataGrid.Rows[i].Cells[(int)IndexInfoColumns.index_type].Value = indexList[i].index_type;
                indexDataGrid.Rows[i].Cells[(int)IndexInfoColumns.weight1].Value = indexList[i].weight1;
                indexDataGrid.Rows[i].Cells[(int)IndexInfoColumns.weight2].Value = indexList[i].weight2;
                indexDataGrid.Rows[i].Cells[(int)IndexInfoColumns.enable_sensitivity].Value = indexList[i].enable_sensitivity==1;
                indexDataGrid.Rows[i].Cells[(int)IndexInfoColumns.sensitivity].Value = indexList[i].sensitivity;

                CommonData.IndexInfo[indexList[i].id] = Index.Copy(indexList[i]);

            }


            //鼠标指针恢复默认状态
            Cursor.Current = Cursors.Default;
            Logger.Log("获取指标信息成功");
        }
        void fetchManagerInfo()
        {
            if (CommonData.ManagerInfo == null)
            {
                CommonData.ManagerInfo = new Dictionary<int, Manager>();
            }//初始化部门信息字典

            Logger.Log("开始获取职能部门信息");
            //鼠标指针变为等待状态
            Cursor.Current = Cursors.WaitCursor;


            //清空原有数据
            managerDataGrid.Rows.Clear();
            //获取指标信息
            var managerMapper = ManagerMapper.GetInstance();
            var managerList = managerMapper.GetAllObjects();


            for (int i = 0; i < managerList.Count; i++)
            {

                managerDataGrid.Rows.Add();
                managerDataGrid.Rows[i].Cells[(int)ManagerInfoColumns.id].Value = managerList[i].id;
                managerDataGrid.Rows[i].Cells[(int)ManagerInfoColumns.manager_code].Value = managerList[i].manager_code;
                managerDataGrid.Rows[i].Cells[(int)ManagerInfoColumns.manager_name].Value = managerList[i].manager_name;

                CommonData.ManagerInfo[managerList[i].id] = Manager.Copy(managerList[i]);

            }

            //鼠标指针恢复默认状态
            Cursor.Current = Cursors.Default;
            Logger.Log("获取职能部门信息成功");
        }

        private void fetchDutyInfo()
        {
            Logger.Log("开始获取职责分配信息");
            //鼠标指针变为等待状态
            Cursor.Current = Cursors.WaitCursor;


            fetchManagerInfo();
            fetchIndexInfo();
            if (CommonData.DutyInfo == null)
            {
                CommonData.DutyInfo = new Dictionary<int, IndexDuty>();
            }//初始化部门信息字典

            var indexDutyMapper = IndexDutyMapper.GetInstance();
            var indexDutyList = indexDutyMapper.GetAllObjects();
            for (int i = 0; i < indexDutyList.Count; i++)
            {
                CommonData.DutyInfo[indexDutyList[i].id] = IndexDuty.Copy(indexDutyList[i]);
            }

            var managers = CommonData.ManagerInfo;
            listManager.Items.Clear();
            foreach (var manager in managers.Values)
            {
                listManager.Items.Add(manager);
                listManager.DisplayMember = "manager_name";
                listManager.ValueMember = "id";

            }

            var unallocatedIndexes = CommonData.UnallocatedIndexes;
            listUnallocatedIndexes.Items.Clear();
            foreach (var index in unallocatedIndexes.Values)
            {
                listUnallocatedIndexes.Items.Add(index);
                listUnallocatedIndexes.DisplayMember = "index_name";
                listUnallocatedIndexes.ValueMember = "id";
            }

            listAllocatedIndexes.Items.Clear();
            textSelectedManager.Text = "";
            buttonCancelSelectManager.Enabled = false;
            CommonData.selectedManager = null;
            //鼠标指针恢复默认状态


            Cursor.Current = Cursors.Default;
            Logger.Log("获取职责分配信息成功");

        }
        private void 保存ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Logger.logger = textLogger;
            Logger.Log("欢迎使用DHU考核系统");
            mainContainer.SizeMode = TabSizeMode.Fixed;//用户只能从菜单栏切换视图
            fetchDepartmentInfo();

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void splitContainer15_SplitterMoved(object sender, SplitterEventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void 部门视图ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainContainer.SelectedIndex = 0;
            labelView.Text = "教学科研单位视图";
            fetchDepartmentInfo();
        }

        private void 指标视图ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainContainer.SelectedIndex = 1;
            labelView.Text = "指标视图";
            fetchIndexInfo();
        }

        private void 完成度视图ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainContainer.SelectedIndex = 2;
            labelView.Text = "职能部门视图";
            fetchManagerInfo();
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void mainContainer_Click(object sender, EventArgs e)
        {
           
        }

        private void mainContainer_Selecting(object sender, TabControlCancelEventArgs e)
        {
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void 修改年份ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainContainer.SelectedIndex = 3;
            labelView.Text = "职责分配视图";
            fetchDutyInfo();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            fetchDepartmentInfo();
        }

        private void deptDataGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void deptDataGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            object cellValue = deptDataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            //获取当前格子字段名，根据DeptInfoColumns的枚举值来获取
            string columnName = Enum.GetName(typeof(DeptInfoColumns), e.ColumnIndex);
            if (e.RowIndex>=CommonData.DeptInfo.Count)
            {
                //新增行时，写入数据库
                var deptMapper = DepartmentMapper.GetInstance();
                var deptAnnualInfoMapper = DeptAnnualInfoMapper.GetInstance();

                var newDeptInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(new Department()));
                var newDeptAnnualInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(new DeptAnnualInfo(-1, CommonData.CurrentYear, 0, 0, "1-1")));

                if (newDeptInfo.Keys.Contains(columnName))
                {
                    newDeptInfo[columnName] = cellValue;
                }
                if(newDeptAnnualInfo.Keys.Contains(columnName))
                {
                    newDeptAnnualInfo[columnName] = cellValue;
                }
                //由于不知道部门id，所以只能先插入部门信息，然后从数据库中获取id，再插入年度信息
                var newDeptInfoObj = JsonConvert.DeserializeObject<Department>(JsonConvert.SerializeObject(newDeptInfo));
                deptMapper.Add(newDeptInfoObj);

                newDeptInfo.Remove("id");//移除id字段
                newDeptInfoObj = deptMapper.GetObject(newDeptInfo);//获取刚插入的部门信息，带有id
                
                newDeptAnnualInfo["dept_id"] = newDeptInfoObj.id;
                var newDeptAnnualInfoObj = JsonConvert.DeserializeObject<DeptAnnualInfo>(JsonConvert.SerializeObject(newDeptAnnualInfo));
                deptAnnualInfoMapper.Add(newDeptAnnualInfoObj);


                CommonData.DeptInfo[newDeptInfoObj.id] = new Tuple<Department, DeptAnnualInfo>(newDeptInfoObj, newDeptAnnualInfoObj);

                deptDataGrid.Rows[e.RowIndex].Cells[(int)DeptInfoColumns.id].Value = newDeptInfoObj.id;
                deptDataGrid.Rows[e.RowIndex].Cells[(int)DeptInfoColumns.dept_code].Value = newDeptInfoObj.dept_code;
                deptDataGrid.Rows[e.RowIndex].Cells[(int)DeptInfoColumns.dept_name].Value = newDeptInfoObj.dept_name;
                deptDataGrid.Rows[e.RowIndex].Cells[(int)DeptInfoColumns.dept_population].Value = newDeptAnnualInfoObj.dept_population;
                deptDataGrid.Rows[e.RowIndex].Cells[(int)DeptInfoColumns.dept_punishment].Value = newDeptAnnualInfoObj.dept_punishment;
                deptDataGrid.Rows[e.RowIndex].Cells[(int)DeptInfoColumns.dept_group].Value = newDeptAnnualInfoObj.dept_group;
                //将新增的部门信息写入数据表
                Logger.Log($"新增部门{newDeptInfoObj.id}");
                return;

            }
            int dept_id = (int)deptDataGrid.Rows[e.RowIndex].Cells[0].Value;

            var deptInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(CommonData.DeptInfo[dept_id].Item1));
            var deptAnnualInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(CommonData.DeptInfo[dept_id].Item2));

            
            
            if (deptInfo.Keys.Contains(columnName)&& deptInfo[columnName].ToString() != cellValue.ToString())
            {
                //写数据库
                var deptMapper = DepartmentMapper.GetInstance();
                Logger.Log($"部门{dept_id}的{columnName}由{deptInfo[columnName]}变更为{cellValue}");
                deptInfo[columnName] = cellValue;
                
                var deptInfoObj=JsonConvert.DeserializeObject<Department>(JsonConvert.SerializeObject(deptInfo));

                //更新内存中的数据
                CommonData.DeptInfo[dept_id]= new Tuple<Department, DeptAnnualInfo>(deptInfoObj, CommonData.DeptInfo[dept_id].Item2);

                deptMapper.Update(deptInfoObj);
                
            }

            if(deptAnnualInfo.Keys.Contains(columnName) && deptAnnualInfo[columnName].ToString() != cellValue.ToString())
            {
                //写数据库
                var deptAnnualInfoMapper = DeptAnnualInfoMapper.GetInstance();
                Logger.Log($"部门{dept_id}在{CommonData.CurrentYear}年的{columnName}由{deptAnnualInfo[columnName]}变更为{cellValue}");
                deptAnnualInfo[columnName] = cellValue;
                var deptAnnualInfoObj = JsonConvert.DeserializeObject<DeptAnnualInfo>(JsonConvert.SerializeObject(deptAnnualInfo));

                //更新内存中的数据
                CommonData.DeptInfo[dept_id] = new Tuple<Department, DeptAnnualInfo>(CommonData.DeptInfo[dept_id].Item1, deptAnnualInfoObj);

                deptAnnualInfoMapper.Update(deptAnnualInfoObj);
               
            }

        }

        private void deptDataGrid_ColumnRemoved(object sender, DataGridViewColumnEventArgs e)
        {

        }

        private void deptDataGrid_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {

        }

        private void deptDataGrid_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {

            //如果数据表被清空，不做任何操作
            if (deptDataGrid.Rows.Count == 0)
            {
                return;
            }


            List<int> removedDeptIds = new List<int>();

            //把删除后的数据表与内存中的数据进行比对，找出被删除的部门
            foreach (var deptId in CommonData.DeptInfo.Keys)
            {
                bool isRemoved = true;
                for (int i = 0; i < deptDataGrid.Rows.Count; i++)
                {
                    if (deptDataGrid.Rows[i].Cells[0].Value == null) continue;
                    if (deptId == (int)deptDataGrid.Rows[i].Cells[0].Value)
                    {
                        isRemoved = false;
                        break;
                    }
                }
                if (isRemoved)
                {
                    removedDeptIds.Add(deptId);
                }
            }

            //删除数据库中的数据
            var deptMapper = DepartmentMapper.GetInstance();
            var deptAnnualInfoMapper = DeptAnnualInfoMapper.GetInstance();
            foreach (var deptId in removedDeptIds)
            {
                deptMapper.Remove(deptId.ToString());
                deptAnnualInfoMapper.Remove(deptId.ToString());
                Logger.Log($"删除部门{deptId}");
                CommonData.DeptInfo.Remove(deptId);
            }
        }
        private void dump2Sheet(DataGridView dataGridView)
        {
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                //excel导出
                if (saveDialog.FilterIndex == 1)
                {
                    FileIO.DataGridViewToExcel(dataGridView, saveDialog.FileName);
                }

                //csv导出
                if (saveDialog.FilterIndex == 2)
                {
                    FileIO.DataGridViewToCSV(dataGridView, saveDialog.FileName);
                }

            }
        }
        private void buttonDeptDump_Click(object sender, EventArgs e)
        {
            dump2Sheet(deptDataGrid);
        }

        private void 修改年份ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ChangeYear changeYear = new ChangeYear();
            changeYear.ShowDialog();
        }

        private void 完成度视图ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            mainContainer.SelectedIndex = 4;
            labelView.Text = "完成度视图";
        }

        private void buttonIndexRefresh_Click(object sender, EventArgs e)
        {
            fetchIndexInfo();
        }

        private bool HasSelectedManager//是否选中了职能部门
        {
            get
            {
                return textSelectedManager.Text != "";
            }
        }
        private void indexDataGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            object cellValue = indexDataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;

            string columnName = Enum.GetName(typeof(IndexInfoColumns), e.ColumnIndex);
            if (e.RowIndex >= CommonData.IndexInfo.Count)
            {
                //新增行时，写入数据库
                var indexMapper = IndexMapper.GetInstance();
                var newIndexInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(new Index()));
                
                newIndexInfo[columnName] = cellValue;//更新字段值

                var newIndexInfoObj = JsonConvert.DeserializeObject<Index>(JsonConvert.SerializeObject(newIndexInfo));
                indexMapper.Add(newIndexInfoObj);
                newIndexInfo.Remove("id");//移除id字段

                newIndexInfoObj = indexMapper.GetObject(newIndexInfo);//获取刚插入的部门信息，带有id
                
                CommonData.IndexInfo[newIndexInfoObj.id] = Index.Copy(newIndexInfoObj);

                indexDataGrid.Rows[e.RowIndex].Cells[(int)IndexInfoColumns.id].Value = newIndexInfoObj.id;
                indexDataGrid.Rows[e.RowIndex].Cells[(int)IndexInfoColumns.index_code].Value = newIndexInfoObj.index_code;
                indexDataGrid.Rows[e.RowIndex].Cells[(int)IndexInfoColumns.index_name].Value = newIndexInfoObj.index_name;
                indexDataGrid.Rows[e.RowIndex].Cells[(int)IndexInfoColumns.index_type].Value = newIndexInfoObj.index_type;
                indexDataGrid.Rows[e.RowIndex].Cells[(int)IndexInfoColumns.weight1].Value = newIndexInfoObj.weight1;
                indexDataGrid.Rows[e.RowIndex].Cells[(int)IndexInfoColumns.weight2].Value = newIndexInfoObj.weight2;
                indexDataGrid.Rows[e.RowIndex].Cells[(int)IndexInfoColumns.enable_sensitivity].Value = newIndexInfoObj.enable_sensitivity;
                indexDataGrid.Rows[e.RowIndex].Cells[(int)IndexInfoColumns.sensitivity].Value = newIndexInfoObj.sensitivity;
                
                //将新增的指标信息写入数据表
                Logger.Log($"新增指标{newIndexInfoObj.id}");
                return;

            }
            int index_id = (int)indexDataGrid.Rows[e.RowIndex].Cells[0].Value;
            var indexInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(CommonData.IndexInfo[index_id]));


            if (indexInfo[columnName].ToString() != cellValue.ToString())
            {
                //写数据库
                var indexMapper = IndexMapper.GetInstance();
                Logger.Log($"指标{index_id}的{columnName}由{indexInfo[columnName]}变更为{cellValue}");

                if(cellValue is bool)//如果是bool类型，转为int,因为SQLite不支持bool类型
                {
                    cellValue = (bool)cellValue ? 1 : 0;
                }

                indexInfo[columnName] = cellValue;

                var indexInfoObj = JsonConvert.DeserializeObject<Index>(JsonConvert.SerializeObject(indexInfo));
                
                //更新内存中的数据
                CommonData.IndexInfo[index_id] = Index.Copy(indexInfoObj);
                indexMapper.Update(indexInfoObj);

            }
        }

        private void buttonIndexDump_Click(object sender, EventArgs e)
        {
            dump2Sheet(indexDataGrid);
        }

        private void indexDataGrid_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {

            //如果数据表被清空，不做任何操作
            if (indexDataGrid.Rows.Count == 0)
            {
                return;
            }


            List<int> removedIndexIds = new List<int>();

            //把删除后的数据表与内存中的数据进行比对，找出被删除的部门
            foreach (var indexId in CommonData.IndexInfo.Keys)
            {
                bool isRemoved = true;
                for (int i = 0; i < indexDataGrid.Rows.Count; i++)
                {
                    if (indexDataGrid.Rows[i].Cells[0].Value == null) continue;
                    if (indexId == (int)indexDataGrid.Rows[i].Cells[0].Value)
                    {
                        isRemoved = false;
                        break;
                    }
                }
                if (isRemoved)
                {
                    removedIndexIds.Add(indexId);
                }
            }

            //删除数据库中的数据
            var indexMapper = IndexMapper.GetInstance();
            foreach (var indexId in removedIndexIds)
            {
                indexMapper.Remove(indexId.ToString());
                Logger.Log($"删除指标{indexId}");
                CommonData.IndexInfo.Remove(indexId);
            }
        }

        private void buttonManagerDump_Click(object sender, EventArgs e)
        {
            dump2Sheet(managerDataGrid);
        }

        private void buttonManagerRefresh_Click(object sender, EventArgs e)
        {
            fetchManagerInfo();
        }

        private void managerDataGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            object cellValue = managerDataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;

            string columnName = Enum.GetName(typeof(ManagerInfoColumns), e.ColumnIndex);
            if (e.RowIndex >= CommonData.ManagerInfo.Count)
            {
                //新增行时，写入数据库
                var managerMapper = ManagerMapper.GetInstance();
                var newManagerInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(new Manager()));
                newManagerInfo[columnName] = cellValue;//更新字段值
                var newManagerInfoObj = JsonConvert.DeserializeObject<Manager>(JsonConvert.SerializeObject(newManagerInfo));
                managerMapper.Add(newManagerInfoObj);
                newManagerInfo.Remove("id");//移除id字段
                newManagerInfoObj = managerMapper.GetObject(newManagerInfo);//获取刚插入的职能部门信息，带有id
                CommonData.ManagerInfo[newManagerInfoObj.id] = Manager.Copy(newManagerInfoObj); 


                managerDataGrid.Rows[e.RowIndex].Cells[(int)ManagerInfoColumns.id].Value = newManagerInfoObj.id;
                managerDataGrid.Rows[e.RowIndex].Cells[(int)ManagerInfoColumns.manager_code].Value = newManagerInfoObj.manager_code;
                managerDataGrid.Rows[e.RowIndex].Cells[(int)ManagerInfoColumns.manager_name].Value = newManagerInfoObj.manager_name;
                //将新增的职能部门信息写入数据表
                Logger.Log($"新增职能部门{newManagerInfoObj.id}");
                return;
            }
            int manager_id = (int)managerDataGrid.Rows[e.RowIndex].Cells[0].Value;
            var managerInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(CommonData.ManagerInfo[manager_id]));

            if (managerInfo[columnName].ToString() != cellValue.ToString())
            {
                //写数据库
                var managerMapper = ManagerMapper.GetInstance();
                Logger.Log($"职能部门{manager_id}的{columnName}由{managerInfo[columnName]}变更为{cellValue}");
                
                managerInfo[columnName] = cellValue;
                
                var managerInfoObj = JsonConvert.DeserializeObject<Manager>(JsonConvert.SerializeObject(managerInfo));

                //更新内存中的数据
                CommonData.ManagerInfo[manager_id] = Manager.Copy(managerInfoObj);
                managerMapper.Update(managerInfoObj);

            }


        }

        private void managerDataGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void managerDataGrid_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            //如果数据表被清空，不做任何操作
            if (managerDataGrid.Rows.Count == 0)
            {
                return;
            }


            List<int> removedManagerIds = new List<int>();

            //把删除后的数据表与内存中的数据进行比对，找出被删除的部门
            foreach (var managerId in CommonData.ManagerInfo.Keys)
            {
                bool isRemoved = true;
                for (int i = 0; i < managerDataGrid.Rows.Count; i++)
                {
                    if (managerDataGrid.Rows[i].Cells[0].Value == null) continue;
                    if (managerId == (int)managerDataGrid.Rows[i].Cells[0].Value)
                    {
                        isRemoved = false;
                        break;
                    }
                }
                if (isRemoved)
                {
                    removedManagerIds.Add(managerId);
                }
            }

            //删除数据库中的数据
            var managerMapper = ManagerMapper.GetInstance();
            foreach (var managerId in removedManagerIds)
            {
                managerMapper.Remove(managerId.ToString());
                Logger.Log($"删除职能部门{managerId}");
                CommonData.ManagerInfo.Remove(managerId);
            }
        }
        
        private void refreshDutyAllocateButtonState()
        {
            //如果没有选中职能部门，不允许分配职责
            //如果没有选中指标，不允许分配职责
            //如果没有选中已经分配的指标，不允许取消分配
            buttonAllocateDuty.Enabled = HasSelectedManager && listUnallocatedIndexes.SelectedItem != null;
            buttonAllocatedDutyAll.Enabled= HasSelectedManager&&listUnallocatedIndexes.Items.Count > 0;

            buttonUnallocateDuty.Enabled = HasSelectedManager && listAllocatedIndexes.SelectedItem != null;
            buttonUnallocateDutyAll.Enabled = HasSelectedManager && listAllocatedIndexes.Items.Count > 0;
        }
        private void refreshCurrentManager()
        {
            var manager = (Manager)listManager.SelectedItem;
            //如果没有选中任何职能部门，不做任何操作
            if (manager == null)
            {
                //todo，要清空别的表单
                listAllocatedIndexes.Items.Clear();

                textSelectedManager.Text = "";
                buttonCancelSelectManager.Enabled = false;
                CommonData.selectedManager = null;
            }
            else
            {
                textSelectedManager.Text = manager.manager_name;
                buttonCancelSelectManager.Enabled = true;
                CommonData.selectedManager= manager;
                var indexDutyMapper = IndexDutyMapper.GetInstance();
                var indexDutyList = indexDutyMapper.GetIndexDutyByManagerId(manager.id);
                listAllocatedIndexes.Items.Clear();
                listAllocatedIndexes.DisplayMember = "index_name";
                listAllocatedIndexes.ValueMember = "id";
                foreach (var indexDuty in indexDutyList)
                {
                    var index = CommonData.IndexInfo[indexDuty.index_id];
                    listAllocatedIndexes.Items.Add(index);
                }
            }
            refreshDutyAllocateButtonState();
        }
        private void listManager_DoubleClick(object sender, EventArgs e)
        {
            refreshCurrentManager();
        }

        private void listManager_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textSelectedManager_TextChanged(object sender, EventArgs e)
        {

        }

        private void listUnallocatedIndexes_SelectedIndexChanged(object sender, EventArgs e)
        {
            refreshDutyAllocateButtonState();
        }

        private void listAllocatedIndexes_SelectedIndexChanged(object sender, EventArgs e)
        {
            refreshDutyAllocateButtonState();
        }

        private void buttonCancelSelectManager_Click(object sender, EventArgs e)
        {
            
            listManager.ClearSelected();    
            refreshCurrentManager();
        }
        private void allocateDuty(ListBox.SelectedObjectCollection selectedList)
        {
            var manager = CommonData.selectedManager;
            List<Index> indexes = new List<Index>();
            foreach (var index in selectedList)
            {
                indexes.Add((Index)index);
            }
            var indexDutyMapper = IndexDutyMapper.GetInstance();
            foreach (var index in indexes)
            {
                var indexDuty = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(new IndexDuty(-1, manager.id, index.id, 1)));

                indexDuty.Remove("id");//移除id字段
                indexDutyMapper.Add(JsonConvert.DeserializeObject<IndexDuty>(JsonConvert.SerializeObject(indexDuty)));
                var indexDutyObj = indexDutyMapper.GetObject(indexDuty);//获取刚插入的职责信息，带有id
                CommonData.DutyInfo[indexDutyObj.id] = IndexDuty.Copy(indexDutyObj);

                //更新列表
                listAllocatedIndexes.Items.Add(index);
                listUnallocatedIndexes.Items.Remove(index);
                Logger.Log($"为职能部门\"{manager.manager_name}\"分配指标\"{index.index_name}\"");
            }
        }
        private void unallocateDuty(ListBox.SelectedObjectCollection selectedList)
        {
            var manager = CommonData.selectedManager;
            List<Index> indexes = new List<Index>();
            foreach (var index in selectedList)
            {
                indexes.Add((Index)index);
            }
            var indexDutyMapper = IndexDutyMapper.GetInstance();
            foreach (var index in indexes)
            {
                var indexDuty = JsonConvert.DeserializeObject<Dictionary<string, object>>
                    (
                    JsonConvert.SerializeObject(
                        indexDutyMapper.GetIndexDutyByIndexAndManagerId(index.id, manager.id)
                        )
                    );
                indexDutyMapper.Remove(indexDuty["id"].ToString());//删除数据库中的数据，根据id
                Logger.Log($"为职能部门\"{manager.manager_name}\"取消指标\"{index.index_name}\"的分配");
                CommonData.DutyInfo.Remove(CommonData.DutyInfo.First(x => x.Value.manager_id == manager.id && x.Value.index_id == index.id).Key);
                listAllocatedIndexes.Items.Remove(index);
                listUnallocatedIndexes.Items.Add(index);
                listUnallocatedIndexes.DisplayMember = "index_name";
                listUnallocatedIndexes.ValueMember = "id";
            }
        }
        private void buttonAllocateDuty_Click(object sender, EventArgs e)
        {
            allocateDuty(listUnallocatedIndexes.SelectedItems);
        }

        private void buttonAllocatedDutyAll_Click(object sender, EventArgs e)
        {
            listUnallocatedIndexes.SelectedItems.Clear();
            for(int i = 0; i < listUnallocatedIndexes.Items.Count; i++)
            {
                listUnallocatedIndexes.SetSelected(i, true);
            }
            allocateDuty(listUnallocatedIndexes.SelectedItems);
        }

        private void buttonUnallocateDuty_Click(object sender, EventArgs e)
        {
            unallocateDuty(listAllocatedIndexes.SelectedItems);
        }

        private void buttonUnallocateDutyAll_Click(object sender, EventArgs e)
        {
            listAllocatedIndexes.SelectedItems.Clear();
            for (int i = 0; i < listAllocatedIndexes.Items.Count; i++)
            {
                listAllocatedIndexes.SetSelected(i, true);
            }
            unallocateDuty(listAllocatedIndexes.SelectedItems);
        }
    }
}
