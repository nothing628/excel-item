using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IE182
{
    public partial class Setting : Form
    {
        public Setting()
        {
            InitializeComponent();
        }

        private void LoadSettings()
        {
            numericUpDown1.Value = Properties.Settings.Default.RowStart;

            listBox1.Items.Clear();

            foreach (var str in Properties.Settings.Default.IgnoreColumn)
            {
                listBox1.Items.Add(str);
            }
        }

        private void SaveSettings()
        {
            Properties.Settings.Default.RowStart = Convert.ToInt32(numericUpDown1.Value);
            Properties.Settings.Default.IgnoreColumn.Clear();

            foreach (string item in listBox1.Items)
            {
                Properties.Settings.Default.IgnoreColumn.Add(item);
            }

            Properties.Settings.Default.Save();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            SaveSettings();
            Close();
        }

        private void Setting_Load(object sender, EventArgs e)
        {
            LoadSettings();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex > -1)
            {
                //Remove from list
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Add to list
        }
    }
}
