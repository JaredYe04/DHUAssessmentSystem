using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using 考核系统.Utils;
using 考核系统.Entity;
namespace 考核系统.Dialogs
{
    public partial class FormExportWizard : Form
    {
        public FormExportWizard()
        {
            InitializeComponent();
        }

        private void buttonSelectAll_Click(object sender, EventArgs e)
        {
            for(int i = 0; i < checkListManager.Items.Count; i++)
            {
                checkListManager.SetItemChecked(i, true);
            }
           
        }

        private void buttonClearSelect_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkListManager.Items.Count; i++)
            {
                checkListManager.SetItemChecked(i, false);
            }
        }

        private void FormExportWizard_Load(object sender, EventArgs e)
        {
            var managers = CommonData.ManagerInfo;
            int idx = 0;
            foreach (var manager in managers)
            {
                
                if(!CommonData.DutyInfo.Any(x => x.Value.manager_id == manager.Value.id))
                {
                    continue;
                }
                else
                {
                    checkListManager.Items.Add(manager.Value.manager_name);
                    checkListManager.SetItemChecked(idx++, true);
                }
            }
            checkListManager.Tag = managers;//绑定数据
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            folderBrowser.Description = "请选择导出路径";
            folderBrowser.ShowDialog();
            
            var path = folderBrowser.SelectedPath;
            if(path == "")
            {
                return;
            }
            var selectedManagers = new List<Manager>();
            var managers = (checkListManager.Tag as Dictionary<int, Manager>).ToList();
            Cursor = Cursors.WaitCursor;
            for (int i = 0; i < checkListManager.Items.Count; i++)
            {
                progressBar.Value = (i + 1) * 100 / checkListManager.Items.Count;
                if (checkListManager.GetItemChecked(i))
                {
                    var manager=managers.Where(m => m.Value.manager_name == checkListManager.Items[i].ToString()).FirstOrDefault();
                    FileIO.ExportEmptyCompletionTable(manager.Value.id, path);
                    
                    Logger.Log(path + "\\" + manager.Value.manager_name + "考核表.xlsx 导出成功");
                }
            }
            Cursor = Cursors.Default;
            MessageBox.Show("导出成功","提示",MessageBoxButtons.OK,MessageBoxIcon.Information);
            this.Hide();
        }
    }
}
