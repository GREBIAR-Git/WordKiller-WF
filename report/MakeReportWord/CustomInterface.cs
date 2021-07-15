using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace MakeReportWord
{
    public partial class CustomInterface : Form
    {
        public CustomInterface()
        {
            InitializeComponent();
        }

        private void tableLayoutPanel1_CellPaint(object sender, TableLayoutCellPaintEventArgs e)
        {
            if (e.Row == 0 || e.Row == 1)
                using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 253, 219, 124)))
                    e.Graphics.FillRectangle(brush, e.CellBounds);
            else if (e.Row == 3 || e.Row == 4)
                using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 208, 117, 252)))
                    e.Graphics.FillRectangle(brush, e.CellBounds);
            else if (e.Row == 6 || e.Row == 7)
                using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 84, 213, 245)))
                    e.Graphics.FillRectangle(brush, e.CellBounds);
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            this.BackColor = Color.FromArgb(255, 50, 39, 62);
            label1.BackColor = Color.FromArgb(255, 253, 219, 124);
            label2.BackColor = Color.FromArgb(255, 253, 219, 124);
            label3.BackColor = Color.FromArgb(255, 208, 117, 252);
            label4.BackColor = Color.FromArgb(255, 208, 117, 252);
            label5.BackColor = Color.FromArgb(255, 84, 213, 245);
            label6.BackColor = Color.FromArgb(255, 84, 213, 245);
            label7.BackColor = Color.FromArgb(255, 253, 219, 124);
            label8.BackColor = Color.FromArgb(255, 253, 219, 124);
            label9.BackColor = Color.FromArgb(255, 208, 117, 252);
            label10.BackColor = Color.FromArgb(255, 84, 213, 245);
            label11.BackColor = Color.FromArgb(255, 84, 213, 245);
            label12.BackColor = Color.FromArgb(255, 50, 39, 62);
            label13.BackColor = Color.FromArgb(255, 50, 39, 62);
            button1.BackColor = Color.FromArgb(255, 238, 230, 246);
            button2.BackColor = Color.FromArgb(255, 238, 230, 246);
            button3.BackColor = Color.FromArgb(255, 238, 230, 246);
            button4.BackColor = Color.FromArgb(255, 238, 230, 246);
            button5.BackColor = Color.FromArgb(255, 238, 230, 246);
            button6.BackColor = Color.FromArgb(255, 238, 230, 246);
            button7.BackColor = Color.FromArgb(255, 238, 230, 246);
            button8.BackColor = Color.FromArgb(255, 238, 230, 246);
            button9.BackColor = Color.FromArgb(255, 238, 230, 246);
            comboBox2.BackColor = Color.FromArgb(255, 238, 230, 246);
            comboBox3.BackColor = Color.FromArgb(255, 238, 230, 246);
            comboBox4.BackColor = Color.FromArgb(255, 238, 230, 246);
            comboBox5.BackColor = Color.FromArgb(255, 238, 230, 246);
            tableLayoutPanel1.BackColor = Color.FromArgb(255, 50, 39, 62);
            tableLayoutPanel2.BackColor = Color.FromArgb(255, 50, 39, 62);
            tableLayoutPanel3.BackColor = Color.FromArgb(255, 50, 39, 62);
            label12.ForeColor = Color.FromArgb(255, 238, 230, 246);
            label13.ForeColor = Color.FromArgb(255, 238, 230, 246);
            label1.Focus();
            showTop(sender, e);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            showBottom(sender, e);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            showTop(sender, e);
        }

        private void showBottom(object sender, EventArgs e)
        {
            button2.Visible = false;
            tableLayoutPanel1.Visible = false;
            button3.Visible = true;
            tableLayoutPanel3.Visible = true;
        }

        private void showTop(object sender, EventArgs e)
        {
            tableLayoutPanel3.Visible = false;
            button3.Visible = false;
            tableLayoutPanel1.Visible = true;
            button2.Visible = true;
        }

        private void CloseWindow_Click(object sender, EventArgs e)
        {
            CloseWindow.Checked = !CloseWindow.Checked;
        }

        private async void ReadScroll_Click(object sender, EventArgs e)
        {
            MakeReport report = new MakeReport();
            string faculty = comboBox1.Text;
            string numberLab = maskedTextBox1.Text;
            string theme = textBox1.Text;
            string discipline = textBox2.Text;
            string professor = textBox3.Text;
            string year = textBox4.Text;
            await Task.Run(() => report.CreateReportLab(faculty, numberLab, theme, discipline, professor, year));
            if (CloseWindow.Checked)
            {
                Application.Exit();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (button8.Text == "К тексту")
            {
                label7.Visible = false;
                label8.Visible = false;
                label9.Visible = false;
                label11.Visible = false;
                label10.Visible = false;
                comboBox2.Visible = false;
                comboBox4.Visible = false;
                comboBox5.Visible = false;
                comboBox3.Visible = false;
                button1.Visible = false;
                button4.Visible = false;
                button5.Visible = false;
                button6.Visible = false;
                button7.Visible = false;
                button9.Visible = false;
                button8.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
                button8.Text = "К подстановкам";
                label13.Text = "текст";
            }
            else
            {
                label7.Visible = true;
                label8.Visible = true;
                label9.Visible = true;
                label11.Visible = true;
                label10.Visible = true;
                comboBox2.Visible = true;
                comboBox4.Visible = true;
                comboBox5.Visible = true;
                comboBox3.Visible = true;
                button1.Visible = true;
                button4.Visible = true;
                button5.Visible = true;
                button6.Visible = true;
                button7.Visible = true;
                button9.Visible = true;
                button8.Margin = new System.Windows.Forms.Padding(3, 13, 3, 3);
                button8.Text = "К тексту";
                label13.Text = "ничто";
            }
        }

        private void Lab1_Click(object sender, EventArgs e)
        {
            Lab1.Checked = !Lab1.Checked;
        }
    }
}
