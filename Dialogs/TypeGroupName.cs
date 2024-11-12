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
    public partial class TypeGroupName : Form
    {
        public TypeGroupName()
        {
            InitializeComponent();
        }

        private void buttonConfirm_Click(object sender, EventArgs e)
        {
            if(textGroupName.Text == "")
            {
                MessageBox.Show("请输入组名","提示",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                return;
            }
            this.DialogResult = DialogResult.OK;
            this.Hide();
        }
    }
}
