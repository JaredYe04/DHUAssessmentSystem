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
    public partial class EditIndexIdentifier: Form
    {
        public EditIndexIdentifier()
        {
            InitializeComponent();
        }


        public void indexIdentifierDataGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            object cellValue = indexIdentifierDataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            if (cellValue == null) cellValue = "";
            string columnName = Enum.GetName(typeof(IndexIdentifierInfoColumns), e.ColumnIndex);
            if (e.RowIndex >= CommonData.IdentifierInfo.Count)
            {
                //新增行时，写入数据库
                var indexIdentifierMapper = IndexIdentifierMapper.GetInstance();
                var newIndexIdentifierInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(new IndexIdentifier()));
                newIndexIdentifierInfo[columnName] = cellValue.ToString();//更新字段值
                IndexIdentifier newIndexIdentifierInfoObj = null;
                try
                {
                    newIndexIdentifierInfoObj = JsonConvert.DeserializeObject<IndexIdentifier>(JsonConvert.SerializeObject(newIndexIdentifierInfo));
                }

                catch (Exception ex)
                {
                    MessageBox.Show("输入数据格式错误", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    indexIdentifierDataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                    return;
                }


                indexIdentifierMapper.Add(newIndexIdentifierInfoObj, false);
                newIndexIdentifierInfoObj = indexIdentifierMapper.GetObject(newIndexIdentifierInfo);//获取刚插入的部门信息，带有id
                CommonData.IdentifierInfo[newIndexIdentifierInfoObj.id] = IndexIdentifier.Copy(newIndexIdentifierInfoObj);

                indexIdentifierDataGrid.Rows[e.RowIndex].Cells[(int)IndexIdentifierInfoColumns.id].Value = newIndexIdentifierInfoObj.id;
                indexIdentifierDataGrid.Rows[e.RowIndex].Cells[(int)IndexIdentifierInfoColumns.identifier_name].Value = newIndexIdentifierInfoObj.identifier_name;
                //将新增的指标分类信息写入数据表

                Logger.Log($"新增指标分类{newIndexIdentifierInfoObj.id}");
                return;
            }


            ////////以上是新增行的操作，以下是修改行的操作

            int identifier_id = Int32.Parse(indexIdentifierDataGrid.Rows[e.RowIndex].Cells[0].Value.ToString());

            //由于改的是主键，所以要特殊处理
            var obj = new IndexIdentifier();
            if (CommonData.IdentifierInfo.ContainsKey(identifier_id) == false)
            {
                //从第二列获取该行的数据
                var identifierName = (string)indexIdentifierDataGrid.Rows[e.RowIndex].Cells[1].Value;
                var currentRow = CommonData.IdentifierInfo.FirstOrDefault(x => x.Value.identifier_name == identifierName);
                obj = currentRow.Value;
                var oldIdentifierId = currentRow.Key;
                CommonData.IdentifierInfo.Remove(oldIdentifierId);
                var indexIdentifierMapper = IndexIdentifierMapper.GetInstance();
                indexIdentifierMapper.Remove(oldIdentifierId.ToString());
                obj.id = identifier_id;
                indexIdentifierMapper.Add(obj, false);
                CommonData.IdentifierInfo[identifier_id] = IndexIdentifier.Copy(obj);
                Logger.Log($"更改指标{obj.identifier_name}分类编号至:{identifier_id}");
                return;//由于改的是主键，因此要删了再加
            }
            else if (e.ColumnIndex == 0 && CommonData.IdentifierInfo.ContainsKey(identifier_id))
            {
                //说明是修改主键，但是主键已经存在
                if (!(sender is Boolean))
                    MessageBox.Show("指标分类编号不能重复", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                var identifierName = (string)indexIdentifierDataGrid.Rows[e.RowIndex].Cells[1].Value;
                var oldId = CommonData.IdentifierInfo.FirstOrDefault(x => x.Value.identifier_name == identifierName).Value.id;
                indexIdentifierDataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = oldId;
                return;
            }
            else
            {
                obj = CommonData.IdentifierInfo[identifier_id];
            }
            var identifierInfo = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(obj));

            if (identifierInfo[columnName].ToString() != cellValue.ToString())
            {
                //写数据库
                var indexIdentifierMapper = IndexIdentifierMapper.GetInstance();
                Logger.Log($"指标分类{identifier_id}的{columnName}由{identifierInfo[columnName]}变更为{cellValue}");
                identifierInfo[columnName] = cellValue;

                IndexIdentifier identifierInfoObj = null;
                try
                {
                    identifierInfoObj = JsonConvert.DeserializeObject<IndexIdentifier>(JsonConvert.SerializeObject(identifierInfo));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("输入数据格式错误", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    indexIdentifierDataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                    return;
                }
                //更新内存中的数据

                CommonData.IdentifierInfo[identifier_id] = IndexIdentifier.Copy(identifierInfoObj);
                indexIdentifierMapper.Update(identifierInfoObj);
            }

            //updateComboIndexIdentifier();
        }

        public void indexIdentifierDataGrid_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            //如果数据表被清空，不做任何操作
            if (indexIdentifierDataGrid.Rows.Count == 0)
            {
                return;
            }

            List<int> removedIndexIdentifierIds = new List<int>();

            foreach (var indexIdentifierId in CommonData.IdentifierInfo.Keys)
            {
                bool isRemoved = true;
                for (int i = 0; i < indexIdentifierDataGrid.Rows.Count; i++)
                {
                    if (indexIdentifierDataGrid.Rows[i].Cells[0].Value == null) continue;
                    if (indexIdentifierId == (int)indexIdentifierDataGrid.Rows[i].Cells[0].Value)
                    {
                        isRemoved = false;
                        break;
                    }
                }
                if (isRemoved)
                {
                    removedIndexIdentifierIds.Add(indexIdentifierId);
                }
            }

            var indexIdentifierMapper = IndexIdentifierMapper.GetInstance();
            foreach (var indexIdentifierId in removedIndexIdentifierIds)
            {
                indexIdentifierMapper.Remove(indexIdentifierId.ToString());
                Logger.Log($"删除指标分类{indexIdentifierId}");
                CommonData.IdentifierInfo.Remove(indexIdentifierId);
            }
            //updateComboIndexIdentifier();
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
        }
        private void EditIndexIdentifier_Load(object sender, EventArgs e)
        {
            fetchIndexIdentifierInfo();
        }
    }
}
