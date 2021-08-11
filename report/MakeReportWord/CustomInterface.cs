using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace MakeReportWord
{
    public partial class CustomInterface : Form
    {
        string text;

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
            label11.BackColor = Color.FromArgb(255, 84, 213, 245);
            label12.BackColor = Color.FromArgb(255, 50, 39, 62);
            label13.BackColor = Color.FromArgb(255, 50, 39, 62);
            button1.BackColor = Color.FromArgb(255, 238, 230, 246);
            button2.BackColor = Color.FromArgb(255, 238, 230, 246);
            button3.BackColor = Color.FromArgb(255, 238, 230, 246);
            button4.BackColor = Color.FromArgb(255, 238, 230, 246);
            button5.BackColor = Color.FromArgb(255, 238, 230, 246);
            button6.BackColor = Color.FromArgb(255, 238, 230, 246);
            button8.BackColor = Color.FromArgb(255, 238, 230, 246);
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
            try
            {
                string text = richTextBox1.Text;
                await Task.Run(() => report.CreateReportLab(faculty, numberLab, theme, discipline, professor, year, "Текст1☺h1Текст2☺h1Текст3"));
            }
            catch
            {
                MessageBox.Show("Что-то пошло не так :(", "Ошибка",MessageBoxButtons.OK,MessageBoxIcon.Error,MessageBoxDefaultButton.Button1,MessageBoxOptions.DefaultDesktopOnly);
            }

            if (CloseWindow.Checked)
            {
                Application.Exit();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (button8.Text == "К тексту")
            {
                tableLayoutPanel4.Visible = false;
                button8.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
                button8.Text = "К подстановкам";
                label13.Text = "текст";
                richTextBox1.Text = text;
            }
            else
            {
                tableLayoutPanel4.Visible = true;
                button8.Margin = new System.Windows.Forms.Padding(3, 13, 3, 3);
                button8.Text = "К тексту";
                label13.Text = "нечто";
                richTextBox1.Text = "";
            }
        }

        private void Lab1_Click(object sender, EventArgs e)
        {
            Lab1.Checked = !Lab1.Checked;
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }


        private void AddToComboBox(ComboBox comboBox, string element)
        {
            if (!comboBox.Items.Contains(element))
            {
                comboBox.Items.Add(element);
                comboBox.SelectedIndex = comboBox.Items.IndexOf(element);
            }
        }

        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            
            if (comboBox.SelectedIndex != -1)
            {
                for (int i = 4; i < 8; i++)
                {
                    ComboBox comboBoxToDeselect;
                    if (i != tableLayoutPanel4.Controls.IndexOf(comboBox))
                    {
                        comboBoxToDeselect = (ComboBox)(tableLayoutPanel4.Controls[i]);
                        comboBoxToDeselect.SelectedIndex = -1;
                    }
                }
                Label13StartText(sender);
                label13.Text += (comboBox.Items.IndexOf(comboBox.SelectedItem) + 1).ToString();
                richTextBox1.Text = comboBox.SelectedItem.ToString();
            }
            else
            {
                label13.Text = "нечто";
                richTextBox1.Text = "";
            }
        }

        private void Label13StartText(object sender)
        {
            Control senderControl = (Control)sender;
            label13.Text = tableLayoutPanel4.Controls[tableLayoutPanel4.Controls.IndexOf(senderControl) - 4].Text + ": ";
        }

        private void ComboBox_MouseDown(object sender, MouseEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            if (e.Button == MouseButtons.Right)
            {
                if (Control.ModifierKeys != Keys.Shift && Control.ModifierKeys != Keys.Control && Control.ModifierKeys != Keys.Alt)
                {
                    for (int i = 4; i < 8; i++)
                    {
                        comboBox = (ComboBox)(tableLayoutPanel4.Controls[i]);
                        comboBox.SelectedIndex = -1;
                    }
                }
                else if (Control.ModifierKeys == Keys.Shift)
                {
                    if (comboBox.SelectedIndex > 0)
                    {
                        int cursorSave = richTextBox1.SelectionStart;
                        string save = comboBox.Items[comboBox.SelectedIndex].ToString();
                        comboBox.Items[comboBox.SelectedIndex] = comboBox.Items[comboBox.SelectedIndex - 1];
                        comboBox.Items[comboBox.SelectedIndex - 1] = save;
                        comboBox.SelectedIndex--;
                        richTextBox1.SelectionStart = cursorSave;
                    }
                }
                else if (Control.ModifierKeys == Keys.Control)
                {
                    if (comboBox.SelectedIndex < comboBox.Items.Count - 1)
                    {
                        int cursorSave = richTextBox1.SelectionStart;
                        string save = comboBox.Items[comboBox.SelectedIndex].ToString();
                        comboBox.Items[comboBox.SelectedIndex] = comboBox.Items[comboBox.SelectedIndex + 1];
                        comboBox.Items[comboBox.SelectedIndex + 1] = save;
                        comboBox.SelectedIndex++;
                        richTextBox1.SelectionStart = cursorSave;
                    }
                }
                else if (Control.ModifierKeys == Keys.Alt)
                {
                    comboBox.Items.RemoveAt(comboBox.SelectedIndex);
                    ComboBox_SelectedIndexChanged(sender, e);
                }
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            if (label13.Text != "нечто" && label13.Text != "текст")
            {
                ComboBox comboBox = new ComboBox();
                comboBox.Visible = false;
                if (label13.Text.StartsWith("Заголовок 1"))
                {
                    comboBox = comboBox2;
                }
                else if (label13.Text.StartsWith("Заголовок 2"))
                {
                    comboBox = comboBox4;
                }
                else if (label13.Text.StartsWith("Список"))
                {
                    comboBox = comboBox5;
                }
                else if (label13.Text.StartsWith("Картинка"))
                {
                    comboBox = comboBox3;
                }
                if (comboBox.Visible == true)
                {
                    int cursorSave = richTextBox1.SelectionStart;
                    comboBox.Items[comboBox.SelectedIndex] = richTextBox1.Text;
                    richTextBox1.SelectionStart = cursorSave;
                }
            }
            if (label13.Text == "текст")
            {
                text = richTextBox1.Text;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AddToComboBox(comboBox2, richTextBox1.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            AddToComboBox(comboBox4, richTextBox1.Text);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            AddToComboBox(comboBox5, richTextBox1.Text);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            AddToComboBox(comboBox3, richTextBox1.Text);
            // picture
        }

        private void Save1_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "|*.wordkiller;";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileStream fileStream = File.Open(saveFileDialog.FileName, FileMode.Create);
                StreamWriter output = new StreamWriter(fileStream);

                output.WriteLine("comboBox1==" + comboBox1.SelectedItem.ToString());
                output.WriteLine("maskedTextBox1==" + maskedTextBox1.Text);
                output.WriteLine("textBox1==" + textBox1.Text);
                output.WriteLine("textBox2==" + textBox2.Text);
                output.WriteLine("textBox3==" + textBox3.Text);
                output.WriteLine("textBox4==" + textBox4.Text);

                for (int i = 0; i < comboBox2.Items.Count; i++)
                {
                    output.WriteLine("comboBox2.Items[i]==" + comboBox2.Items[i].ToString());
                }
                for (int i = 0; i < comboBox4.Items.Count; i++)
                {
                    output.WriteLine("comboBox4.Items[i]==" + comboBox4.Items[i].ToString());
                }
                for (int i = 0; i < comboBox5.Items.Count; i++)
                {
                    output.WriteLine("comboBox5.Items[i]==" + comboBox5.Items[i].ToString());
                }
                for (int i = 0; i < comboBox3.Items.Count; i++)
                {
                    output.WriteLine("comboBox3.Items[i]==" + comboBox3.Items[i].ToString());
                }
                output.WriteLine("###textstart");
                output.WriteLine(text);
                output.WriteLine("###textend");

                output.Close();
            }
        }

        private void Open1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "|*.Аллянов;";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileStream file = new FileStream(openFileDialog.FileName, FileMode.Open);
                StreamReader reader = new StreamReader(file);
                try
                {
                    string data = reader.ReadToEnd();
                    for (int i = 1; i < data.Length; i++)
                    {
                        if (data[i - 1] == '\r')
                        {
                            data = data.Remove(i, 1);
                        }
                    }
                    string[] lines = data.Split('\r');
                }
                catch
                {
                    MessageBox.Show("Файл повреждён");
                }
                reader.Close();
            }
        }

        public string ProcessSpecial(int i,string special)
        {
            string text = string.Empty;
            if(special=="h1")
            {
                text=comboBox2.Items[i-1].ToString();
            }
            else if (special == "h2")
            {
                text=comboBox4.Items[i-1].ToString();
            }
            else if (special == "l")
            {
                text=comboBox5.Items[i-1].ToString();
            }
            else if (special == "p")
            {
                text=comboBox3.Items[i-1].ToString();
            }
            return text;
        }
    }
}
