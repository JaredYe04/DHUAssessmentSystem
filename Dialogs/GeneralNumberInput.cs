using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 考核系统.Dialogs
{
    public partial class GeneralNumberInput : Form
    {
        public GeneralNumberInput(string caption)
        {
           
            InitializeComponent();
            this.Text = caption;
        }

        private void textNumber_TextChanged(object sender, EventArgs e)
        {
            //如果不是数字，把非数字字符去掉
            if (!System.Text.RegularExpressions.Regex.IsMatch(textNumber.Text, "^[0-9]*$"))
            {
                textNumber.Text = System.Text.RegularExpressions.Regex.Replace(textNumber.Text, "[^0-9]", "");
            }
        }

        private void GeneralNumberInput_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textNumber.Text == "")
            {
                MessageBox.Show("请输入数字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            this.DialogResult = DialogResult.OK;
            this.Hide();
        }
    }
}
