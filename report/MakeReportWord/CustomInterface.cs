using System;
using System.Drawing;
using System.IO;
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
            facultyLabel.BackColor = Color.FromArgb(255, 253, 219, 124);
            numberLabLabel.BackColor = Color.FromArgb(255, 253, 219, 124);
            themeLabel.BackColor = Color.FromArgb(255, 208, 117, 252);
            disciplineLabel.BackColor = Color.FromArgb(255, 208, 117, 252);
            professorLabel.BackColor = Color.FromArgb(255, 84, 213, 245);
            yearLabel.BackColor = Color.FromArgb(255, 84, 213, 245);
            heading1Label.BackColor = Color.FromArgb(255, 253, 219, 124);
            heading2Label.BackColor = Color.FromArgb(255, 253, 219, 124);
            listLabel.BackColor = Color.FromArgb(255, 208, 117, 252);
            pictureLabel.BackColor = Color.FromArgb(255, 84, 213, 245);
            label1.BackColor = Color.FromArgb(255, 50, 39, 62);
            label2.BackColor = Color.FromArgb(255, 50, 39, 62);
            buttonHeading1.BackColor = Color.FromArgb(255, 238, 230, 246);
            buttonDown.BackColor = Color.FromArgb(255, 238, 230, 246);
            buttonUp.BackColor = Color.FromArgb(255, 238, 230, 246);
            buttonHeading2.BackColor = Color.FromArgb(255, 238, 230, 246);
            buttonList.BackColor = Color.FromArgb(255, 238, 230, 246);
            buttonPicture.BackColor = Color.FromArgb(255, 238, 230, 246);
            button1.BackColor = Color.FromArgb(255, 238, 230, 246);
            heading1ComboBox.BackColor = Color.FromArgb(255, 238, 230, 246);
            pictureComboBox.BackColor = Color.FromArgb(255, 238, 230, 246);
            heading2ComboBox.BackColor = Color.FromArgb(255, 238, 230, 246);
            listComboBox.BackColor = Color.FromArgb(255, 238, 230, 246);
            tableLayoutPanel1.BackColor = Color.FromArgb(255, 50, 39, 62);
            tableLayoutPanel2.BackColor = Color.FromArgb(255, 50, 39, 62);
            tableLayoutPanel3.BackColor = Color.FromArgb(255, 50, 39, 62);
            label1.ForeColor = Color.FromArgb(255, 238, 230, 246);
            label2.ForeColor = Color.FromArgb(255, 238, 230, 246);
            facultyLabel.Focus();
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
            buttonDown.Visible = false;
            tableLayoutPanel1.Visible = false;
            buttonUp.Visible = true;
            tableLayoutPanel3.Visible = true;
        }

        private void showTop(object sender, EventArgs e)
        {
            tableLayoutPanel3.Visible = false;
            buttonUp.Visible = false;
            tableLayoutPanel1.Visible = true;
            buttonDown.Visible = true;
        }

        private void CloseWindow_Click(object sender, EventArgs e)
        {
            CloseWindow.Checked = !CloseWindow.Checked;
        }

        private async void ReadScroll_Click(object sender, EventArgs e)
        {
            MakeReport report = new MakeReport();
            string faculty = facultyComboBox.Text;
            string numberLab = numberLabTextBox.Text;
            string theme = themeTextBox.Text;
            string discipline = disciplineTextBox.Text;
            string professor = professorTextBox.Text;
            string year = yearTextBox.Text;
            try
            {
                string text = richTextBox.Text;
                UserInput userInput = new UserInput();
                userInput.ComboBoxH1 = new string[heading1ComboBox.Items.Count];
                for (int i = 0; i < heading1ComboBox.Items.Count; i++)
                {
                    userInput.ComboBoxH1[i] = heading1ComboBox.Items[i].ToString();
                }
                userInput.ComboBoxH2 = new string[heading2ComboBox.Items.Count];
                for (int i = 0; i < heading2ComboBox.Items.Count; i++)
                {
                    userInput.ComboBoxH2[i] = heading2ComboBox.Items[i].ToString();
                }
                userInput.ComboBoxL = new string[listComboBox.Items.Count];
                for (int i = 0; i < listComboBox.Items.Count; i++)
                {
                    userInput.ComboBoxL[i] = listComboBox.Items[i].ToString();
                }
                userInput.ComboBoxP = new string[pictureComboBox.Items.Count];
                for (int i = 0; i < pictureComboBox.Items.Count; i++)
                {
                    userInput.ComboBoxP[i] = pictureComboBox.Items[i].ToString();
                }
                userInput.Text = this.text;
                await Task.Run(() => report.CreateReportLab(faculty, numberLab, theme, discipline, professor, year, userInput));
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
            if (button1.Text == "К тексту")
            {
                tableLayoutPanel4.Visible = false;
                button1.Text = "К подстановкам";
                label2.Text = "текст";
                richTextBox.Text = text;
                buttonSpecial.Visible = true;
            }
            else
            {
                tableLayoutPanel4.Visible = true;
                button1.Text = "К тексту";
                label2.Text = "нечто";
                richTextBox.Text = "";
                buttonSpecial.Visible = false;
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
                label2.Text += (comboBox.Items.IndexOf(comboBox.SelectedItem) + 1).ToString();
                richTextBox.Text = comboBox.SelectedItem.ToString();
            }
            else
            {
                label2.Text = "нечто";
                richTextBox.Text = "";
            }
        }

        private void Label13StartText(object sender)
        {
            Control senderControl = (Control)sender;
            label2.Text = tableLayoutPanel4.Controls[tableLayoutPanel4.Controls.IndexOf(senderControl) - 4].Text + ": ";
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
                        int cursorSave = richTextBox.SelectionStart;
                        string save = comboBox.Items[comboBox.SelectedIndex].ToString();
                        comboBox.Items[comboBox.SelectedIndex] = comboBox.Items[comboBox.SelectedIndex - 1];
                        comboBox.Items[comboBox.SelectedIndex - 1] = save;
                        comboBox.SelectedIndex--;
                        richTextBox.SelectionStart = cursorSave;
                    }
                }
                else if (Control.ModifierKeys == Keys.Control)
                {
                    if (comboBox.SelectedIndex < comboBox.Items.Count - 1)
                    {
                        int cursorSave = richTextBox.SelectionStart;
                        string save = comboBox.Items[comboBox.SelectedIndex].ToString();
                        comboBox.Items[comboBox.SelectedIndex] = comboBox.Items[comboBox.SelectedIndex + 1];
                        comboBox.Items[comboBox.SelectedIndex + 1] = save;
                        comboBox.SelectedIndex++;
                        richTextBox.SelectionStart = cursorSave;
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
            if (label2.Text != "нечто" && label2.Text != "текст")
            {
                ComboBox comboBox = new ComboBox();
                comboBox.Visible = false;
                if (label2.Text.StartsWith("Заголовок 1"))
                {
                    comboBox = heading1ComboBox;
                }
                else if (label2.Text.StartsWith("Заголовок 2"))
                {
                    comboBox = heading2ComboBox;
                }
                else if (label2.Text.StartsWith("Список"))
                {
                    comboBox = listComboBox;
                }
                else if (label2.Text.StartsWith("Картинка"))
                {
                    comboBox = pictureComboBox;
                }
                if (comboBox.Visible == true)
                {
                    int cursorSave = richTextBox.SelectionStart;
                    comboBox.Items[comboBox.SelectedIndex] = richTextBox.Text;
                    richTextBox.SelectionStart = cursorSave;
                }
            }
            if (label2.Text == "текст")
            {
                text = richTextBox.Text;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AddToComboBox(heading1ComboBox, richTextBox.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            AddToComboBox(heading2ComboBox, richTextBox.Text);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            AddToComboBox(listComboBox, richTextBox.Text);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            AddToComboBox(pictureComboBox, richTextBox.Text);
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

                output.WriteLine("comboBox1==" + facultyComboBox.SelectedItem.ToString());
                output.WriteLine("maskedTextBox1==" + numberLabTextBox.Text);
                output.WriteLine("textBox1==" + themeTextBox.Text);
                output.WriteLine("textBox2==" + disciplineTextBox.Text);
                output.WriteLine("textBox3==" + professorTextBox.Text);
                output.WriteLine("textBox4==" + yearTextBox.Text);

                for (int i = 0; i < heading1ComboBox.Items.Count; i++)
                {
                    output.WriteLine("comboBox2.Items[i]==" + heading1ComboBox.Items[i].ToString());
                }
                for (int i = 0; i < heading2ComboBox.Items.Count; i++)
                {
                    output.WriteLine("comboBox4.Items[i]==" + heading2ComboBox.Items[i].ToString());
                }
                for (int i = 0; i < listComboBox.Items.Count; i++)
                {
                    output.WriteLine("comboBox5.Items[i]==" + listComboBox.Items[i].ToString());
                }
                for (int i = 0; i < pictureComboBox.Items.Count; i++)
                {
                    output.WriteLine("comboBox3.Items[i]==" + pictureComboBox.Items[i].ToString());
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
            openFileDialog.Filter = "|*.wordkiller;";
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

        private void button7_Click(object sender, EventArgs e)
        {
            richTextBox.Text += "☺";
        }
    }
}
