using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using 考核系统.Entity;
using 考核系统.Mapper;
using 考核系统.Utils;

namespace 考核系统.Dialogs
{
    public partial class EditGroups : Form
    {
        public EditGroups()
        {
            InitializeComponent();
        }
        public void fetchGroupsInfo()
        {
            if (CommonData.GroupInfo == null)
            {
                CommonData.GroupInfo = new Dictionary<int, Groups>();
            }

            Logger.Log("开始获取单位组别信息");
            //鼠标指针变为等待状态
            Cursor.Current = Cursors.WaitCursor;


            groupDataGrid.Rows.Clear();
            //获取信息
            var groupMapper = GroupsMapper.GetInstance();
            var groupList = groupMapper.GetAllObjects();
            for( int i = 0; i < groupList.Count; i++)
            {
                groupDataGrid.Rows.Add();
                groupDataGrid.Rows[i].Cells[0].Value = groupList[i].id;
                groupDataGrid.Rows[i].Cells[1].Value = groupList[i].group_name;
                CommonData.GroupInfo[groupList[i].id] = Groups.Copy(groupList[i]);
            }

            //鼠标指针恢复默认状态
            Cursor.Current = Cursors.Default;
            Logger.Log("获取单位组别成功");
        }
        private void EditGroups_Load(object sender, EventArgs e)
        {
            fetchGroupsInfo();
        }

        private void groupDataGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            object cellValue = groupDataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;

            string columnName = groupDataGrid.Columns[e.ColumnIndex].Name;
            if (e.RowIndex >= CommonData.GroupInfo.Count)
            {
                //新增行时，写入数据库
                var groupMapper = GroupsMapper.GetInstance();
                var newGroupInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(new Groups()));
                newGroupInfo[columnName] = cellValue;//更新字段值
                var newGroupInfoObj = JsonConvert.DeserializeObject<Groups>(JsonConvert.SerializeObject(newGroupInfo));
                groupMapper.Add(newGroupInfoObj);
                newGroupInfo.Remove("id");//移除id字段
                newGroupInfoObj = groupMapper.GetObject(newGroupInfo);//获取刚插入的单位组别信息，带有id
                CommonData.GroupInfo[newGroupInfoObj.id] = Groups.Copy(newGroupInfoObj);


                groupDataGrid.Rows[e.RowIndex].Cells[(int)GroupColumns.id].Value = newGroupInfoObj.id;
                groupDataGrid.Rows[e.RowIndex].Cells[(int)GroupColumns.group_name].Value = newGroupInfoObj.group_name;
                //将新增的单位组别信息写入数据表
                Logger.Log($"新增单位组别{newGroupInfoObj.id}");
                return;
            }
            int groupId = (int)groupDataGrid.Rows[e.RowIndex].Cells[(int)GroupColumns.id].Value;
            var groupInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(CommonData.GroupInfo[groupId]));

            if (groupInfo[columnName] == null || groupInfo[columnName].ToString() != cellValue.ToString())
            {
                //写数据库
                var groupMapper = GroupsMapper.GetInstance();
                Logger.Log($"单位组别{groupId}的{columnName}由{groupInfo[columnName]}变更为{cellValue}");

                groupInfo[columnName] = cellValue;

                var groupInfoObj = JsonConvert.DeserializeObject<Groups>(JsonConvert.SerializeObject(groupInfo));

                //更新内存中的数据
                CommonData.GroupInfo[groupId] = Groups.Copy(groupInfoObj);
                groupMapper.Update(groupInfoObj);
            }


        }

        private void groupDataGrid_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            //如果数据表被清空，不做任何操作
            if (groupDataGrid.Rows.Count == 0)
            {
                return;
            }

            List<int> removedGroupIds = new List<int>();

            //把删除后的数据表与内存中的数据进行比对，找出被删除的单位组别
            foreach (var groupId in CommonData.GroupInfo.Keys)
            {
                bool isRemoved = true;
                for (int i = 0; i < groupDataGrid.Rows.Count; i++)
                {
                    if (groupDataGrid.Rows[i].Cells[0].Value == null) continue;
                    if (groupId == (int)groupDataGrid.Rows[i].Cells[0].Value)
                    {
                        isRemoved = false;
                        break;
                    }
                }
                if (isRemoved)
                {
                    removedGroupIds.Add(groupId);
                }
            }

            //删除数据库中的数据
            var groupMapper = GroupsMapper.GetInstance();
            foreach (var groupId in removedGroupIds)
            {
                groupMapper.Remove(groupId.ToString());
                Logger.Log($"删除单位组别{groupId}");
                CommonData.GroupInfo.Remove(groupId);
            }
        }
    }
}
