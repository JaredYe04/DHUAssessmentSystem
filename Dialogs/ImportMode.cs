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
    public partial class ImportMode : Form
    {
        public ImportMode()
        {
            InitializeComponent();
        }
        public bool ModeFlag
        {
            get
            {
                return radioAppendMode.Checked;//true为追加模式，false为覆盖模式
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            this.Hide();
        }
    }
}
