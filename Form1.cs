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

        public Form1()
        {
            InitializeComponent();
            //订阅年份变更事件
            EventBus.YearChanged += EventBus_YearChanged;
        }
        private DataGridView indexIdentifierDataGrid
        {
            get
            {
                return this.editIndexIdentifier.indexIdentifierDataGrid;

            }
        }
        private void EventBus_YearChanged(int year)
        {
            Logger.Log("年份变更为" + year);
            labelCurrentYear.Text = "当前年份:" + year;
            fetchDepartmentInfo();
            //部门信息会随年份变更而变更
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

                    deptAnnualInfo = new DeptAnnualInfo(deptList[i].id, year, 0, 0, "组1");
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
        void fetchIndexIdentifierInfo()
        {
            if (CommonData.IdentifierInfo == null)
            {
                CommonData.IdentifierInfo = new Dictionary<int, IndexIdentifier>();
            }

            Logger.Log("开始获取指标分类信息");
            //鼠标指针变为等待状态
            Cursor.Current = Cursors.WaitCursor;


            indexIdentifierDataGrid.Rows.Clear();
            //获取指标信息
            var indexIdentifierMapper = IndexIdentifierMapper.GetInstance();
            var indexIdentifierList = indexIdentifierMapper.GetAllObjects();

            for (int i = 0; i < indexIdentifierList.Count; i++)
            {
                indexIdentifierDataGrid.Rows.Add();
                indexIdentifierDataGrid.Rows[i].Cells[0].Value = indexIdentifierList[i].id;
                indexIdentifierDataGrid.Rows[i].Cells[1].Value = indexIdentifierList[i].identifier_name;
                CommonData.IdentifierInfo[indexIdentifierList[i].id] = IndexIdentifier.Copy(indexIdentifierList[i]);
            }
                
            //鼠标指针恢复默认状态
            Cursor.Current = Cursors.Default;
            Logger.Log("获取职能指标分类成功");
            updateComboIndexIdentifier();
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

            CommonData.IndexInfo.Clear();
            foreach (var index in indexList)
            {
                CommonData.IndexInfo[index.id] = Index.Copy(index);
            }
            //指标信息不显示，只有在选中第一级类别时才显示


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
            fetchIndexIdentifierInfo();
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
            if (cellValue == null) cellValue = "0";
            //获取当前格子字段名，根据DeptInfoColumns的枚举值来获取
            string columnName = Enum.GetName(typeof(DeptInfoColumns), e.ColumnIndex);
            if (e.RowIndex>=CommonData.DeptInfo.Count)
            {
                //新增行时，写入数据库
                var deptMapper = DepartmentMapper.GetInstance();
                var deptAnnualInfoMapper = DeptAnnualInfoMapper.GetInstance();

                var newDeptInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(new Department()));
                var newDeptAnnualInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(new DeptAnnualInfo(-1, CommonData.CurrentYear, 0, 0, "组1")));

                if (newDeptInfo.Keys.Contains(columnName))
                {
                    newDeptInfo[columnName] = cellValue.ToString();
                }
                if(newDeptAnnualInfo.Keys.Contains(columnName))
                {
                    newDeptAnnualInfo[columnName] = cellValue.ToString();
                }
                //由于不知道部门id，所以只能先插入部门信息，然后从数据库中获取id，再插入年度信息
                var newDeptInfoObj = JsonConvert.DeserializeObject<Department>(JsonConvert.SerializeObject(newDeptInfo));
                deptMapper.Add(newDeptInfoObj);

                newDeptInfo.Remove("id");//移除id字段
                newDeptInfoObj = deptMapper.GetObject(newDeptInfo);//获取刚插入的部门信息，带有id
                

                newDeptAnnualInfo["dept_id"] = newDeptInfoObj.id;
                var newDeptAnnualInfoObj = JsonConvert.DeserializeObject<DeptAnnualInfo>(JsonConvert.SerializeObject(newDeptAnnualInfo));
                deptAnnualInfoMapper.Add(newDeptAnnualInfoObj);
                newDeptAnnualInfo.Remove("id");//移除id字段
                newDeptAnnualInfoObj = deptAnnualInfoMapper.GetDeptAnnualInfo(newDeptInfoObj.id, CommonData.CurrentYear);//获取刚插入的部门信息，带有id

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

            
            
            if (deptInfo.Keys.Contains(columnName)&& (deptInfo[columnName]==null||deptInfo[columnName].ToString() != cellValue.ToString()))
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

            if(deptAnnualInfo.Keys.Contains(columnName) && (deptAnnualInfo[columnName]==null||deptAnnualInfo[columnName].ToString() != cellValue.ToString()))
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
        private void dump2Sheet(Dictionary<string, DataGridView> sheets, Dictionary<string, bool> reserveId = null,bool structureOnly=false)
        {
            if (reserveId == null)
            {
                reserveId = new Dictionary<string, bool>();
                foreach (var keypair in sheets)
                {
                    reserveId[keypair.Key] = false;
                }
            }
            foreach (var keypair in reserveId)
            {
                if (keypair.Value == true) continue;

                var header = keypair.Key;
                DataGridView dataGridView = sheets[keypair.Key]; ;
                //复制一份数据表，去掉id列
                DataGridView dataGridViewCopy = new DataGridView();
                dataGridViewCopy.ColumnCount = dataGridView.ColumnCount - 1;
                for (int i = 1; i < dataGridView.ColumnCount; i++)
                {
                    dataGridViewCopy.Columns[i - 1].Name = dataGridView.Columns[i].Name;
                    //复制标题
                    dataGridViewCopy.Columns[i - 1].HeaderText = dataGridView.Columns[i].HeaderText;
                }

                for (int i = 0; i < dataGridView.RowCount; i++)
                {
                    dataGridViewCopy.Rows.Add();
                    for (int j = 1; j < dataGridView.ColumnCount; j++)
                    {
                        dataGridViewCopy.Rows[i].Cells[j - 1].Value = dataGridView.Rows[i].Cells[j].Value;
                    }
                }
                dataGridView = dataGridViewCopy;

                sheets[header] = dataGridView;


            }

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileIO.MultiDataGridViewToExcel(sheets, saveDialog.FileName, structureOnly);
                    Logger.Log(saveDialog.FileName + "导出成功");
                }
                catch (Exception ex)
                {
                    Logger.Log(saveDialog.FileName + "导出失败:" + ex.Message);

                }
            }

        }
        private void buttonDeptDump_Click(object sender, EventArgs e)
        {
            saveDialog.FileName= DateTime.Now.ToString("yyyy-MM-dd")+"教学科研单位信息.xlsx";
            var dict = new Dictionary<string, DataGridView>();
            dict.Add("教学科研单位信息表", deptDataGrid);
            dump2Sheet(dict);
        }

        private void 修改年份ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ChangeYear changeYear = new ChangeYear();
            changeYear.ShowDialog();
        }

        private void 完成度视图ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            mainContainer.SelectedIndex = 4;
            labelView.Text = "完成情况视图";
            fetchIndexInfo();
            fetchManagerInfo();
            fetchDepartmentInfo();
            fetchCompletionInfo();
        }

        private void buttonIndexRefresh_Click(object sender, EventArgs e)
        {
            fetchIndexInfo();
            fetchIndexIdentifierInfo(); 
            comboIndexIdentifier.Text = "";//清空选中的一级类别
            indexDataGrid.Enabled=false;
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

            if (cellValue == null && e.ColumnIndex == (int)IndexInfoColumns.enable_sensitivity) return;//应对bug:当新的行选中checkbox，然后取消选中，cellValue为null
            if (cellValue == null) cellValue = "0";

            string columnName = Enum.GetName(typeof(IndexInfoColumns), e.ColumnIndex);
            if (e.RowIndex >= CommonData.currentCategoryIndexes.Count)
            {
                //新增行时，写入数据库
                var indexMapper = IndexMapper.GetInstance();
                var newIndexInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(new Index()));


                newIndexInfo[columnName] = cellValue;//更新字段值

                newIndexInfo["identifier_id"] = CommonData.selectedIdentifier.id;//新增的指标一定属于当前选中的一级类别

                int nextSecondaryIdentifier = 1;
                if (CommonData.IndexInfo.Values.Any(index => index.identifier_id == CommonData.selectedIdentifier.id))
                {
                    nextSecondaryIdentifier = CommonData.IndexInfo.Values.Where(index => index.identifier_id == CommonData.selectedIdentifier.id).Max(index => index.secondary_identifier) + 1;
                }//获取当前一级类别下的最大二级类别，新的二级类别为最大二级类别+1
                newIndexInfo["secondary_identifier"] = nextSecondaryIdentifier;
                newIndexInfo["tertiary_identifier"] = 0;


                var newIndexInfoObj = JsonConvert.DeserializeObject<Index>(JsonConvert.SerializeObject(newIndexInfo));
                indexMapper.Add(newIndexInfoObj);
                newIndexInfo.Remove("id");//移除id字段

                newIndexInfoObj = indexMapper.GetObject(newIndexInfo);//获取刚插入的部门信息，带有id

                CommonData.IndexInfo[newIndexInfoObj.id] = Index.Copy(newIndexInfoObj);

                indexDataGrid.Rows[e.RowIndex].Cells[(int)IndexInfoColumns.id].Value = newIndexInfoObj.id;
                indexDataGrid.Rows[e.RowIndex].Cells[(int)IndexInfoColumns.identifier_id].Value = newIndexInfoObj.identifier_id;
                indexDataGrid.Rows[e.RowIndex].Cells[(int)IndexInfoColumns.secondary_identifier].Value = newIndexInfoObj.secondary_identifier;
                indexDataGrid.Rows[e.RowIndex].Cells[(int)IndexInfoColumns.tertiary_identifier].Value = newIndexInfoObj.tertiary_identifier;
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


            if (indexInfo[columnName] == null || indexInfo[columnName].ToString() != cellValue.ToString())
            {
                //写数据库
                var indexMapper = IndexMapper.GetInstance();
                Logger.Log($"指标{index_id}的{columnName}由{indexInfo[columnName]}变更为{cellValue}");

                if (cellValue is bool)//如果是bool类型，转为int,因为SQLite不支持bool类型
                {
                    cellValue = (bool)cellValue ? 1 : 0;
                }
                //if (columnName == "secondary_identifier")
                //{
                //    //如果是二级类别，判断一下是否有重复
                //    if (CommonData.IndexInfo.Values.Any(index =>
                //    index.secondary_identifier == Int32.Parse(cellValue.ToString())
                //    && index.identifier_id == CommonData.selectedIdentifier.id
                //    ))
                //    {
                //        MessageBox.Show("二级类别不能重复", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //        indexDataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = indexInfo[columnName];
                //        //如果有重复，不做任何操作
                //        return;
                //    }
                //}//因为有了三级类别，所以不再需要判断二级类别是否重复
                indexInfo[columnName] = cellValue;

                var indexInfoObj = JsonConvert.DeserializeObject<Index>(JsonConvert.SerializeObject(indexInfo));

                //更新内存中的数据
                CommonData.IndexInfo[index_id] = Index.Copy(indexInfoObj);
                indexMapper.Update(indexInfoObj);

            }
            //特判三级类别
            if (e.ColumnIndex == (int)IndexInfoColumns.tertiary_identifier)
            {
                //该行为总指标，除了基本信息外，其他信息高亮显示
                for (int i = 0; i < indexDataGrid.ColumnCount; i++)
                {
                    indexDataGrid.Rows[e.RowIndex].Cells[i].Style.BackColor = cellValue.ToString() == "-1" ? Color.LightYellow : Color.White;
                    //if (i == (int)IndexInfoColumns.secondary_identifier || i == (int)IndexInfoColumns.tertiary_identifier || i == (int)IndexInfoColumns.index_name)
                    //{
                    //}
                    //else
                    //{
                    //    //indexDataGrid.Rows[e.RowIndex].Cells[i].ReadOnly = (cellValue.ToString() == "-1");
                        
                    //}
                }
                
            }
        }
        private void buttonIndexDump_Click(object sender, EventArgs e)
        {
            saveDialog.FileName = DateTime.Now.ToString("yyyy-MM-dd") + "考核指标信息.xlsx";
            var dict = new Dictionary<string, DataGridView>();
            dict.Add("指标一级类别信息表", indexIdentifierDataGrid);
            var fullIndexDataGrid = new DataGridView();
            fullIndexDataGrid.ColumnCount = indexDataGrid.ColumnCount;
            for (int i = 0; i < indexDataGrid.ColumnCount; i++)
            {
                fullIndexDataGrid.Columns[i].Name = indexDataGrid.Columns[i].Name;
                fullIndexDataGrid.Columns[i].HeaderText = indexDataGrid.Columns[i].HeaderText;
            }
            foreach(var index in CommonData.IndexInfo)
            {
                fullIndexDataGrid.Rows.Add();
                fullIndexDataGrid.Rows[fullIndexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.id].Value = index.Value.id;
                fullIndexDataGrid.Rows[fullIndexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.identifier_id].Value = index.Value.identifier_id;
                fullIndexDataGrid.Rows[fullIndexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.secondary_identifier].Value = index.Value.secondary_identifier;
                fullIndexDataGrid.Rows[fullIndexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.tertiary_identifier].Value = index.Value.tertiary_identifier;
                fullIndexDataGrid.Rows[fullIndexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.index_name].Value = index.Value.index_name;
                fullIndexDataGrid.Rows[fullIndexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.index_type].Value = index.Value.index_type;
                fullIndexDataGrid.Rows[fullIndexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.weight1].Value = index.Value.weight1;
                fullIndexDataGrid.Rows[fullIndexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.weight2].Value = index.Value.weight2;
                fullIndexDataGrid.Rows[fullIndexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.enable_sensitivity].Value = index.Value.enable_sensitivity;
                fullIndexDataGrid.Rows[fullIndexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.sensitivity].Value = index.Value.sensitivity;

            }
            dict.Add("考核指标信息表", fullIndexDataGrid);
            var reserveId = new Dictionary<string, bool>
            {
                { "指标一级类别信息表", true },
                { "考核指标信息表", false }
            };
            dump2Sheet(dict,reserveId);
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
            foreach (var indexId in CommonData.currentCategoryIndexes.Keys)
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
            saveDialog.FileName = DateTime.Now.ToString("yyyy-MM-dd") + "职能部门信息.xlsx";
            var dict = new Dictionary<string, DataGridView>();
            dict.Add("职能部门信息表", managerDataGrid);
            dump2Sheet(dict);
        }

        private void buttonManagerRefresh_Click(object sender, EventArgs e)
        {
            fetchManagerInfo();
        }
        private void fetchCompletionInfo()
        {
            
            fetchManagerInfo();
            fetchDepartmentInfo();
            fetchDutyInfo();

            labelCurrentIndexCompletion.Text = "当前指标";


            //1.处理树形结构视图
            treeDuty.Nodes.Clear();
            foreach (var managerPair in CommonData.ManagerInfo)
            {
                var manager = managerPair.Value;
                var managerNode = new TreeNode(manager.manager_name);
                managerNode.Tag = manager;
                treeDuty.Nodes.Add(managerNode);
                var indexDutyMapper = IndexDutyMapper.GetInstance();
                var indexDutyList = indexDutyMapper.GetIndexDutyByManagerId(manager.id);
                foreach (var indexDuty in indexDutyList)
                {
                    if (!CommonData.IndexInfo.ContainsKey(indexDuty.index_id)) continue;
                    var nodeText =
                        CommonData.IndexInfo[indexDuty.index_id].identifier_id.ToString() + "." +
                        CommonData.IndexInfo[indexDuty.index_id].secondary_identifier.ToString() +
                        (CommonData.IndexInfo[indexDuty.index_id].tertiary_identifier != "0"
                        && CommonData.IndexInfo[indexDuty.index_id].tertiary_identifier != "-1"
                        ? ("." + CommonData.IndexInfo[indexDuty.index_id].tertiary_identifier) : "")
                        + " " +
                        CommonData.IndexInfo[indexDuty.index_id].index_name;
                    var indexDutyNode = new TreeNode(nodeText);
                    indexDutyNode.Tag = CommonData.IndexInfo[indexDuty.index_id];
                    managerNode.Nodes.Add(indexDutyNode);
                }
            }

            //2.处理数据绑定
            unbindCompletionIndex();

            //3.获取所有指标完成情况
            if (CommonData.CompletionInfo == null) CommonData.CompletionInfo = new Dictionary<int, Completion>();
            CommonData.CompletionInfo.Clear();
            var completionMapper = CompletionMapper.GetInstance();
            var indexCompletionList = completionMapper.GetIndexCompletionByYear(CommonData.CurrentYear);
            CommonData.CompletionInfo = new Dictionary<int, Completion>();
            foreach (var indexCompletion in indexCompletionList)
            {
                CommonData.CompletionInfo[indexCompletion.id] = Completion.Copy(indexCompletion);
            }

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

            if (managerInfo[columnName]==null||managerInfo[columnName].ToString() != cellValue.ToString())
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
                    try
                    {
                        var index = CommonData.IndexInfo[indexDuty.index_id];
                        listAllocatedIndexes.Items.Add(index);
                    }
                    catch(Exception ex)
                    {
                        Logger.Log($"指标{indexDuty.index_id}不存在;"+ex.Message);
                    }
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

        private void indexIdentifierDataGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
            => editIndexIdentifier.indexIdentifierDataGrid_CellEndEdit(sender, e);
        public void updateComboIndexIdentifier()
        {
            var currentSelectedIndexIdentifier = (IndexIdentifier)comboIndexIdentifier.SelectedItem;
            comboIndexIdentifier.Items.Clear();
            foreach (var indexIdentifier in CommonData.IdentifierInfo.Values)
            {
                comboIndexIdentifier.Items.Add(indexIdentifier);
            }
            comboIndexIdentifier.DisplayMember = "identifier_name";
            comboIndexIdentifier.ValueMember = "id";
            if (currentSelectedIndexIdentifier != null)
            {
                comboIndexIdentifier.SelectedItem = currentSelectedIndexIdentifier;
            }
        }
        private void indexIdentifierDataGrid_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        => editIndexIdentifier.indexIdentifierDataGrid_RowsRemoved(sender, e);

        private void comboIndexIdentifier_SelectedIndexChanged(object sender, EventArgs e)
        {
            //indexDataGrid显示选中的指标分类下的指标
            var indexIdentifier = (IndexIdentifier)comboIndexIdentifier.SelectedItem;
            CommonData.selectedIdentifier = indexIdentifier;//更新当前选中的指标分类
            if (indexIdentifier == null)
            {
                indexDataGrid.Enabled = false;
                return;
            }
            indexDataGrid.Enabled = true;
            indexDataGrid.Rows.Clear();

            var selectedIndexes = CommonData.currentCategoryIndexes;
            foreach (var keyValuePair in selectedIndexes)
            {
                var index = keyValuePair.Value;
                indexDataGrid.Rows.Add();
                indexDataGrid.Rows[indexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.id].Value = index.id;
                indexDataGrid.Rows[indexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.identifier_id].Value = index.identifier_id;
                indexDataGrid.Rows[indexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.secondary_identifier].Value = index.secondary_identifier;
                indexDataGrid.Rows[indexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.tertiary_identifier].Value = index.tertiary_identifier;
                indexDataGrid.Rows[indexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.index_name].Value = index.index_name;
                indexDataGrid.Rows[indexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.index_type].Value = index.index_type;
                indexDataGrid.Rows[indexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.weight1].Value = index.weight1;
                indexDataGrid.Rows[indexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.weight2].Value = index.weight2;
                indexDataGrid.Rows[indexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.enable_sensitivity].Value = index.enable_sensitivity;
                indexDataGrid.Rows[indexDataGrid.Rows.Count - 2].Cells[(int)IndexInfoColumns.sensitivity].Value = index.sensitivity;

                if (index.tertiary_identifier == "-1")
                {
                    foreach(DataGridViewCell cell in indexDataGrid.Rows[indexDataGrid.Rows.Count - 2].Cells)
                    {
                        cell.Style.BackColor = Color.LightYellow;
                    }
                }
            }
        }

        private void comboIndexIdentifier_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void buttonDeptTemplateDump_Click(object sender, EventArgs e)
        {
            saveDialog.FileName = "教学科研单位信息模板.xlsx";
            var dict = new Dictionary<string, DataGridView>();
            dict.Add("教学科研单位信息表", deptDataGrid);
            dump2Sheet(dict,null,true);

        }

        private void buttonManagerTemplateDump_Click(object sender, EventArgs e)
        {
            saveDialog.FileName = "职能部门信息模板.xlsx";
            var dict = new Dictionary<string, DataGridView>();
            dict.Add("职能部门信息表", managerDataGrid);
            dump2Sheet(dict, null, true);
        }

        private void 导出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainContainer.SelectedIndex = 5;
            labelView.Text = "导出向导";
        }

        private void 导出向导ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainContainer.SelectedIndex = 5;
            labelView.Text = "导出向导";
        }

        private void buttonIndexTemplateDump_Click(object sender, EventArgs e)
        {
            saveDialog.FileName = "考核指标信息模板.xlsx";
            var dict = new Dictionary<string, DataGridView>();
            dict.Add("指标一级类别信息表", indexIdentifierDataGrid);
            dict.Add("考核指标信息表",indexDataGrid);
            var reserveId = new Dictionary<string, bool>
            {
                { "指标一级类别信息表", true },
                { "考核指标信息表", false }
            };
            dump2Sheet(dict, reserveId,true);
        }

        private void buttonDutyClear_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定要清空所有职责分配吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                var indexDutyMapper = IndexDutyMapper.GetInstance();
                foreach (var indexDuty in CommonData.DutyInfo.Values)
                {
                    indexDutyMapper.Remove(indexDuty.id.ToString());
                }
                CommonData.DutyInfo.Clear();
                listAllocatedIndexes.Items.Clear();
                listUnallocatedIndexes.Items.Clear();
                Logger.Log("清空所有职责分配");
                fetchDutyInfo();
            }
        }

        private void buttonDeptImport_Click(object sender, EventArgs e)
        {
            var importMode = new ImportMode();
            //importMode.ShowDialog();
            //if (importMode.DialogResult != DialogResult.OK) return;
            var importModeValue = importMode.ModeFlag;//获取导入模式,true为追加模式，false为覆盖模式
            

            openDialog.Title = "请选择要导入的部门信息文件";
            openDialog.ShowDialog();
            if (openDialog.FileName == "")
            {
                return;
            }
            //fetchDepartmentInfo();
            var importedDeptInfo = FileIO.ImportSingleSheet(openDialog.FileName);

            for (int i = 0; i < importedDeptInfo.Rows.Count - 1; i++)//最后一行是未提交的新数据，不参与比对
            {
                var exists = false;
                int exist_row_idx = -1;
                if (!importModeValue)
                {
                    //检查一下原本的DeptInfo中，是否存在DeptCode和新数据匹配的数据
                    //注意最后一行是未提交的新数据，不参与比对
                    for (int j=0;j<deptDataGrid.Rows.Count;j++)
                    {
                        if (deptDataGrid.Rows[j].Cells[(int)DeptInfoColumns.id].Value == null) continue;
                        var dept = CommonData.DeptInfo[(int)deptDataGrid.Rows[j].Cells[(int)DeptInfoColumns.id].Value].Item1;
                        

                        if (
                            dept.dept_code ==
                            importedDeptInfo.Rows[i].Cells[(int)DeptInfoColumns.dept_code].Value?.ToString()
                            )
                        {
                            //如果存在，更新数据
                            exists = true;
                            exist_row_idx = j;
                            break;
                        }
                    }
                }
                int cur_idx=(!importModeValue && exists)? exist_row_idx : deptDataGrid.Rows.Add();

                deptDataGrid.Rows[cur_idx].Cells[(int)DeptInfoColumns.dept_name].Value = importedDeptInfo.Rows[i].Cells[(int)DeptInfoColumns.dept_name].Value;
                deptDataGrid_CellEndEdit(sender, new DataGridViewCellEventArgs((int)DeptInfoColumns.dept_name, cur_idx));

                deptDataGrid.Rows[cur_idx].Cells[(int)DeptInfoColumns.dept_code].Value = importedDeptInfo.Rows[i].Cells[(int)DeptInfoColumns.dept_code].Value;
                deptDataGrid_CellEndEdit(sender, new DataGridViewCellEventArgs((int)DeptInfoColumns.dept_code, cur_idx));

                deptDataGrid.Rows[cur_idx].Cells[(int)DeptInfoColumns.dept_population].Value = importedDeptInfo.Rows[i].Cells[(int)DeptInfoColumns.dept_population].Value;
                deptDataGrid_CellEndEdit(sender, new DataGridViewCellEventArgs((int)DeptInfoColumns.dept_population, cur_idx));

                deptDataGrid.Rows[cur_idx].Cells[(int)DeptInfoColumns.dept_punishment].Value = importedDeptInfo.Rows[i].Cells[(int)DeptInfoColumns.dept_punishment].Value;
                deptDataGrid_CellEndEdit(sender, new DataGridViewCellEventArgs((int)DeptInfoColumns.dept_punishment, cur_idx));

                deptDataGrid.Rows[cur_idx].Cells[(int)DeptInfoColumns.dept_group].Value = importedDeptInfo.Rows[i].Cells[(int)DeptInfoColumns.dept_group].Value;
                deptDataGrid_CellEndEdit(sender, new DataGridViewCellEventArgs((int)DeptInfoColumns.dept_group, cur_idx));

                if ((!importModeValue && exists))
                {
                    Logger.Log("部门" + importedDeptInfo.Rows[i].Cells[(int)DeptInfoColumns.dept_code].Value.ToString() + "已存在，更新数据");
                }
                else
                {
                    Logger.Log("新增部门" + importedDeptInfo.Rows[i].Cells[(int)DeptInfoColumns.dept_code].Value.ToString());
                }

            }
            fetchDepartmentInfo();
        }

        private void buttonManagerImport_Click(object sender, EventArgs e)
        {
            var importMode = new ImportMode();
            //importMode.ShowDialog();
            //if (importMode.DialogResult != DialogResult.OK) return;
            var importModeValue = importMode.ModeFlag;//获取导入模式,true为追加模式，false为覆盖模式


            openDialog.Title = "请选择要导入的职能部门信息文件";
            openDialog.ShowDialog();
            if (openDialog.FileName == "")
            {
                return;
            }
            //fetchDepartmentInfo();
            var importedManagerInfo = FileIO.ImportSingleSheet(openDialog.FileName);

            for (int i = 0; i < importedManagerInfo.Rows.Count - 1; i++)//最后一行是未提交的新数据，不参与比对
            {
                var exists = false;
                int exist_row_idx = -1;
                if (!importModeValue)
                {
                    //检查一下原本的DeptInfo中，是否存在DeptCode和新数据匹配的数据
                    //注意最后一行是未提交的新数据，不参与比对
                    for (int j = 0; j < managerDataGrid.Rows.Count; j++)
                    {
                        if (managerDataGrid.Rows[j].Cells[(int)ManagerInfoColumns.id].Value == null) continue;
                        var manager = CommonData.ManagerInfo[(int)managerDataGrid.Rows[j].Cells[(int)ManagerInfoColumns.id].Value];
                        

                        if (
                            manager.manager_code ==
                            importedManagerInfo.Rows[i].Cells[(int)ManagerInfoColumns.manager_code].Value?.ToString()
                            )
                        {
                            //如果存在，更新数据
                            exists = true;
                            exist_row_idx = j;
                            break;
                        }
                    }
                }
                int cur_idx = (!importModeValue && exists) ? exist_row_idx : managerDataGrid.Rows.Add();

                managerDataGrid.Rows[cur_idx].Cells[(int)ManagerInfoColumns.manager_name].Value = importedManagerInfo.Rows[i].Cells[(int)ManagerInfoColumns.manager_name].Value;
                managerDataGrid_CellEndEdit(sender, new DataGridViewCellEventArgs((int)ManagerInfoColumns.manager_name, cur_idx));

                managerDataGrid.Rows[cur_idx].Cells[(int)ManagerInfoColumns.manager_code].Value = importedManagerInfo.Rows[i].Cells[(int)ManagerInfoColumns.manager_code].Value;
                managerDataGrid_CellEndEdit(sender, new DataGridViewCellEventArgs((int)ManagerInfoColumns.manager_code, cur_idx));

                if ((!importModeValue && exists))
                {
                    Logger.Log("职能部门" + importedManagerInfo.Rows[i].Cells[(int)ManagerInfoColumns.manager_code].Value.ToString() + "已存在，更新数据");
                }
                else
                {
                    Logger.Log("新增职能部门" + importedManagerInfo.Rows[i].Cells[(int)ManagerInfoColumns.manager_code].Value.ToString());
                }

            }
            fetchManagerInfo();
        }

        private void buttonIndexImport_Click(object sender, EventArgs e)
        {
            var importMode = new ImportMode();
            //importMode.ShowDialog();
            //if (importMode.DialogResult != DialogResult.OK) return;
            var importModeValue = importMode.ModeFlag;//获取导入模式,true为追加模式，false为覆盖模式


            openDialog.Title = "请选择要导入的考核指标信息文件";
            openDialog.ShowDialog();
            //如果用户在openDialog点了取消，就不打开

            

            if (openDialog.FileName == "")
            {
                return;
            }
            //fetchDepartmentInfo();
            var importedDict = FileIO.ImportMultiSheets(openDialog.FileName);
            var importedIndexInfo=importedDict["考核指标信息表"];
            var importedIndexIdentifierInfo = importedDict["指标一级类别信息表"];
            importedIndexIdentifierInfo.Columns.Remove("编号");
            for (int i = 0; i < importedIndexIdentifierInfo.Rows.Count - 1; i++)//最后一行是未提交的新数据，不参与比对
            {
                var exists = false;
                int exist_row_idx = -1;
                if (!importModeValue)
                {
                    //检查一下原本的DeptInfo中，是否存在DeptCode和新数据匹配的数据
                    //注意最后一行是未提交的新数据，不参与比对
                    for (int j = 0; j < indexIdentifierDataGrid.Rows.Count; j++)
                    {
                        if (indexIdentifierDataGrid.Rows[j].Cells[(int)IndexIdentifierInfoColumns.id].Value == null) continue;
                        var identifier = CommonData.IdentifierInfo[Int32.Parse(indexIdentifierDataGrid.Rows[j].Cells[(int)IndexIdentifierInfoColumns.id].Value.ToString())];

                        if (
                            identifier.id.ToString() ==
                            indexIdentifierDataGrid.Rows[i].Cells[(int)IndexIdentifierInfoColumns.id].Value?.ToString()
                            )
                        {
                            //如果存在，更新数据
                            exists = true;
                            exist_row_idx = j;
                            break;
                        }
                    }
                }
                int cur_idx = (!importModeValue && exists) ? exist_row_idx : indexIdentifierDataGrid.Rows.Add();



                indexIdentifierDataGrid.Rows[cur_idx].Cells[(int)IndexIdentifierInfoColumns.id].Value = importedIndexIdentifierInfo.Rows[i].Cells[(int)IndexIdentifierInfoColumns.id].Value;
                indexIdentifierDataGrid_CellEndEdit(true, new DataGridViewCellEventArgs((int)IndexIdentifierInfoColumns.id, cur_idx));

                indexIdentifierDataGrid.Rows[cur_idx].Cells[(int)IndexIdentifierInfoColumns.identifier_name].Value = importedIndexIdentifierInfo.Rows[i].Cells[(int)IndexIdentifierInfoColumns.identifier_name].Value;
                indexIdentifierDataGrid_CellEndEdit(true, new DataGridViewCellEventArgs((int)IndexIdentifierInfoColumns.identifier_name, cur_idx));

                //if ((!importModeValue && exists))
                //{
                //    Logger.Log("指标分类" + importedIndexIdentifierInfo.Rows[i].Cells[(int)IndexIdentifierInfoColumns.id].Value.ToString() + "已存在，更新数据");

                //}
                //else
                //{
                //    Logger.Log("新增指标分类" + importedIndexIdentifierInfo.Rows[i].Cells[(int)IndexIdentifierInfoColumns.id].Value.ToString());
                //}

            }


            //指标信息的修改不参照原本方法
            for(int i = 0; i < importedIndexInfo.Rows.Count - 1; i++)
            {
                var indexObj = new Index();
                var row=importedIndexInfo.Rows[i];
                //indexObj.id = row.Cells[(int)IndexInfoColumns.id].Value == null ? -1 : Int32.Parse(row.Cells[(int)IndexInfoColumns.id].Value.ToString());
                indexObj.identifier_id = Int32.Parse(row.Cells[(int)IndexInfoColumns.identifier_id].Value.ToString());
                indexObj.secondary_identifier = Int32.Parse(row.Cells[(int)IndexInfoColumns.secondary_identifier].Value.ToString());
                indexObj.tertiary_identifier = row.Cells[(int)IndexInfoColumns.tertiary_identifier].Value == null ? "": row.Cells[(int)IndexInfoColumns.tertiary_identifier].Value.ToString();
                indexObj.index_name = row.Cells[(int)IndexInfoColumns.index_name].Value.ToString();
                indexObj.index_type = row.Cells[(int)IndexInfoColumns.index_type].Value.ToString();
                indexObj.weight1 = row.Cells[(int)IndexInfoColumns.weight1].Value == null ? 0 : double.Parse(row.Cells[(int)IndexInfoColumns.weight1].Value.ToString());
                indexObj.weight2 = row.Cells[(int)IndexInfoColumns.weight2].Value == null ? 0 : double.Parse(row.Cells[(int)IndexInfoColumns.weight2].Value.ToString());
                if (row.Cells[(int)IndexInfoColumns.enable_sensitivity].Value == null) row.Cells[(int)IndexInfoColumns.enable_sensitivity].Value = 0;
                indexObj.enable_sensitivity = Int32.Parse(row.Cells[(int)IndexInfoColumns.enable_sensitivity].Value.ToString());
                indexObj.sensitivity = row.Cells[(int)IndexInfoColumns.sensitivity].Value == null ? 0 : double.Parse(row.Cells[(int)IndexInfoColumns.sensitivity].Value.ToString());
                
                var exists = CommonData.IndexInfo.Values.Any(x => x.identifier_id == indexObj.identifier_id && x.secondary_identifier == indexObj.secondary_identifier);

                var indexMapper = IndexMapper.GetInstance();
                if (exists)
                {
                    //id列不参与比对，要从identifier_id列和secondary_identifier列找到对应的id
                    var currentRow = CommonData.IndexInfo.FirstOrDefault(x => x.Value.identifier_id == indexObj.identifier_id && x.Value.secondary_identifier == indexObj.secondary_identifier);
                    indexObj.id = currentRow.Key;
                    indexMapper.Update(indexObj);
                }
                else
                {
                    indexMapper.Add(indexObj);
                }

            }



            fetchIndexInfo();
            fetchIndexIdentifierInfo();
        }

        private void buttonChangeYear_Click(object sender, EventArgs e)
        {
            ChangeYear changeYear = new ChangeYear();
            changeYear.ShowDialog();
        }
        private EditIndexIdentifier editIndexIdentifier = new EditIndexIdentifier();
        private void buttonEditIdentifier_Click(object sender, EventArgs e)
        {
            editIndexIdentifier.ShowDialog();
            updateComboIndexIdentifier();
        }

        private void buttonCompletionRefresh_Click(object sender, EventArgs e)
        {
            fetchCompletionInfo();
        }
        private void calcCompletionRate(int row_idx)
        {
            var row = completionDataGrid.Rows[row_idx];
            if(
                row.Cells[(int)CompletionColumns.target].Value== null ||
                row.Cells[(int)CompletionColumns.completed].Value == null||
                Int32.Parse(row.Cells[(int)CompletionColumns.target].Value.ToString()) == 0||
                Int32.Parse(row.Cells[(int)CompletionColumns.completed].Value.ToString()) == 0
                )
            {
                row.Cells[(int)CompletionColumns.completion_rate].Value = "";
            }
            int target = Int32.Parse(row.Cells[(int)CompletionColumns.target].Value.ToString());
            int completion = Int32.Parse(row.Cells[(int)CompletionColumns.completed].Value.ToString());
            double rate = (double)completion / target;
            row.Cells[(int)CompletionColumns.completion_rate].Value = rate.ToString("P");
        }
        private void bindCompletionIndex()
        {
            if (CommonData.currentIndexCompletion == null)
            {
                CommonData.currentIndexCompletion = new Dictionary<int, Completion>();
            }
            
            if (treeDuty.SelectedNode != null && treeDuty.SelectedNode.Tag is Index)
            {
                labelCurrentIndexCompletion.Text = "当前指标:" + treeDuty.SelectedNode.Text;
                var index = (Index)treeDuty.SelectedNode.Tag;
                initCompletion(index);
                CommonData.currentCompletionIndex = index;
                var completionList = CompletionMapper.GetInstance().GetCompletionByIndexId(index.id);
                foreach (var completion in completionList)
                {
                    if (!CommonData.currentIndexCompletion.ContainsKey(completion.dept_id))
                    {
                        CommonData.currentIndexCompletion.Add(completion.dept_id, completion);
                    }
                }
                var completions = CommonData.currentIndexCompletion;
                completionDataGrid.Rows.Clear();
                foreach (var item in completions)
                {
                    var completion = item.Value;
                    var dept= CommonData.DeptInfo[completion.dept_id].Item1;
                    var row= completionDataGrid.Rows[completionDataGrid.Rows.Add()];
                    row.Cells[(int)CompletionColumns.id].Value = completion.id;
                    row.Cells[(int)CompletionColumns.dept_code].Value= dept.dept_code;
                    row.Cells[(int)CompletionColumns.dept_name].Value = dept.dept_name;
                    row.Cells[(int)CompletionColumns.target].Value = completion.target;
                    row.Cells[(int)CompletionColumns.completed].Value = completion.completed;
                    calcCompletionRate(row.Index);
                }

            }
            var lastRow = completionDataGrid.Rows[completionDataGrid.Rows.Count - 1];
            lastRow.Cells[(int)CompletionColumns.target].ReadOnly = true;
            lastRow.Cells[(int)CompletionColumns.completed].ReadOnly = true;
        }
        private void unbindCompletionIndex()
        {
            completionDataGrid.Rows.Clear();
            labelCurrentIndexCompletion.Text = "当前指标";
            //treeDuty.SelectedNode = null;
            CommonData.currentIndexCompletion = null;
            CommonData.currentCompletionIndex = null;
        }
        private void initCompletion(Index index)
        {
            //由于某些指标在数据库中没有完成度信息，所以要先创建空的完成度信息
            var completionMapper = CompletionMapper.GetInstance();
            var completionList = completionMapper.GetCompletionByIndexId(index.id);
            var currentYear = CommonData.CurrentYear;
            if (completionList.Count < CommonData.DeptInfo.Count)
            {
                var depts= CommonData.DeptInfo.Values.Select(x => x.Item1);
                foreach (var dept in depts)
                {
                    var completion = new Completion();
                    completion.index_id = index.id;
                    completion.dept_id = dept.id;
                    completion.year = CommonData.CurrentYear;
                    //如果没有完成度信息，就创建一个，有的话就不创建
                    if(!completionList.Any(x => x.dept_id == dept.id && x.index_id==index.id && x.year==currentYear))
                    {
                        completionMapper.Add(completion);
                    }
                    if (CommonData.currentIndexCompletion != null)
                    {
                        //当目前选中了指标，并正在初始化时，要把新创建的完成度信息加入到内存中，但如果没有选中指标，就不需要加入内存
                        CommonData.currentIndexCompletion[dept.id] = completion;
                    }
                    CommonData.CompletionInfo[completion.id] = completion;
                }
                Logger.Log($"为指标{index.index_name}创建了{depts.Count()}个部门的完成度信息");
            }
        }
        private void treeDuty_DoubleClick(object sender, EventArgs e)
        {
            unbindCompletionIndex();
            bindCompletionIndex();
        }

        private void treeDuty_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void completionDataGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            object cellValue = completionDataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;

            string columnName = Enum.GetName(typeof(CompletionColumns), e.ColumnIndex);
            
            int completion_id = (int)completionDataGrid.Rows[e.RowIndex].Cells[0].Value;
            var completionInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(CommonData.CompletionInfo[completion_id]));

            if (completionInfo[columnName] == null || completionInfo[columnName].ToString() != cellValue.ToString())
            {
                //写数据库
                var completionMapper = CompletionMapper.GetInstance();
                Logger.Log($"部门{completion_id}的{columnName}由{completionInfo[columnName]}变更为{cellValue}");

                completionInfo[columnName] = cellValue;

                var completionInfoObj = JsonConvert.DeserializeObject<Completion>(JsonConvert.SerializeObject(completionInfo));

                CommonData.currentIndexCompletion[completion_id] = Completion.Copy(completionInfoObj);

                completionMapper.Update(completionInfoObj);

                calcCompletionRate(e.RowIndex);

            }
        }

        private void buttonCompletionExport_Click(object sender, EventArgs e)
        {
            fetchIndexInfo();
            fetchDepartmentInfo();
            foreach (var index in CommonData.IndexInfo.Values)
            {
                initCompletion(index);
            }
            fetchManagerInfo();
            fetchDutyInfo();
            fetchIndexIdentifierInfo();
            fetchCompletionInfo();
            var formExportWizard=new FormExportWizard();
            formExportWizard.ShowDialog();
        }

        private void buttonCompletionImport_Click(object sender, EventArgs e)
        {
            multiOpenDialog.ShowDialog();
            if(multiOpenDialog.FileNames.Length == 0)
            {
                return;
            }
            int idx = 1;
            foreach (var filename in multiOpenDialog.FileNames)
            {
                FileIO.ImportCompletionTable(filename);
                Logger.Log($"{idx++}/{multiOpenDialog.FileNames.Length}导入{filename}已完成");
            }
        }
    }
}
