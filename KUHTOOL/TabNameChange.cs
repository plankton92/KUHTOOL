using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace KUHTOOL
{
    public partial class TabNameChange : Form
    {
        static private String tab = "";
        public TabNameChange(String tabIndex)
        {
            InitializeComponent();
            tab = tabIndex;
        }

        private void ultraButton2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ultraButton1_Click(object sender, EventArgs e)
        {
            TabPage tappage = (TabPage)this.Owner.Controls.Find("tabPage" + tab, true)[0];
            tappage.Text = ultraTextEditor1.Text;
            File.WriteAllText(@"C:\HIS\temp\TABName" + tab + ".txt", tappage.Text, Encoding.UTF8);
            this.Close();
        }
    }
}
