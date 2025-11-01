using System;
using System.Windows.Forms;

namespace BFeiertage
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape || e.KeyCode == Keys.Enter || e.KeyCode == Keys.F1)
            {
                Close();
            }
        }

        private void Form2_Click(object sender, EventArgs e) => Close();

        private void Label1_Click(object sender, EventArgs e) => Close();

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try { System.Diagnostics.Process.Start("https://de.wikipedia.org/wiki/Gesetzliche_Feiertage_in_Deutschland"); }
            catch (InvalidOperationException ex) { MessageBox.Show(ex.Message, "Fehler"); }
        }
    }
}
