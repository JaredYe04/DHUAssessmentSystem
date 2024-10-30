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
        private void EventBus_YearChanged(int year)
        {
            Logger.Log("年份变更为" + year);
            labelCurrentYear.Text = "当前年份:" + year;
            //todo:更新界面
            //labelYear.Text = year.ToString();
        }
        private async Task fetchDepartmentInfo()
        {
            Logger.Log("开始获取教学科研单位信息");
            //鼠标指针变为等待状态
            Cursor.Current = Cursors.WaitCursor;


            //清空原有数据
            dataGridView1.Rows.Clear();
            //获取教学科研单位信息
            var deptMapper =DepartmentMapper.GetInstance();
            var deptList = await deptMapper.FindAll();
            




            //鼠标指针恢复默认状态
            Cursor.Current = Cursors.Default;
            Logger.Log("获取教学科研单位信息成功");
        }
        private void 保存ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Logger.logger = textLogger;
            Logger.Log("欢迎使用DHU考核系统");
            
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

        private async void 部门视图ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainContainer.SelectedIndex = 0;
            labelView.Text = "教学科研单位视图";
            await fetchDepartmentInfo();
        }

        private void 指标视图ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainContainer.SelectedIndex = 1;
            labelView.Text = "指标视图";
        }

        private void 完成度视图ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mainContainer.SelectedIndex = 2;
            labelView.Text = "完成度视图";
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
            ChangeYear changeYear = new ChangeYear();
            changeYear.ShowDialog();
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            await fetchDepartmentInfo();
        }
    }
}
