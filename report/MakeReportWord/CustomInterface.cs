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
            if (Lab.Checked)
            {
                this.Text = "Сотворение лабораторной работы из небытия";
            }
            else if(Practice.Checked)
            {
                this.Text = "Сотворение практической работы из небытия";
            }
            else if (Kurs.Checked)
            {
                this.Text = "Сотворение курсовой работы из небытия";
            }
            else if (Ref.Checked)
            {
                this.Text = "Сотворение реферата из небытия";
            }
            else if (Diplom.Checked)
            {
                this.Text = "Сотворение дипломной работы из небытия";
            }
            else if (VKR.Checked)
            {
                this.Text = "Сотворение ВКР из небытия";
            }
            else if (RGR.Checked)
            {
                this.Text = "Сотворение РГР из небытия";
            }
        }

        void titlepagePanel_CellPaint(object sender, TableLayoutCellPaintEventArgs e)
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

        void CustomInterface_Shown(object sender, EventArgs e)
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
            displayedLabel.BackColor = Color.FromArgb(255, 50, 39, 62);
            elementLabel.BackColor = Color.FromArgb(255, 50, 39, 62);
            buttonHeading1.BackColor = Color.FromArgb(255, 238, 230, 246);
            buttonDown.BackColor = Color.FromArgb(255, 238, 230, 246);
            buttonUp.BackColor = Color.FromArgb(255, 238, 230, 246);
            buttonHeading2.BackColor = Color.FromArgb(255, 238, 230, 246);
            buttonList.BackColor = Color.FromArgb(255, 238, 230, 246);
            buttonPicture.BackColor = Color.FromArgb(255, 238, 230, 246);
            buttonText.BackColor = Color.FromArgb(255, 238, 230, 246);
            heading1ComboBox.BackColor = Color.FromArgb(255, 238, 230, 246);
            pictureComboBox.BackColor = Color.FromArgb(255, 238, 230, 246);
            heading2ComboBox.BackColor = Color.FromArgb(255, 238, 230, 246);
            listComboBox.BackColor = Color.FromArgb(255, 238, 230, 246);
            titlepagePanel.BackColor = Color.FromArgb(255, 50, 39, 62);
            MainPanel.BackColor = Color.FromArgb(255, 50, 39, 62);
            DownPanel.BackColor = Color.FromArgb(255, 50, 39, 62);
            displayedLabel.ForeColor = Color.FromArgb(255, 238, 230, 246);
            elementLabel.ForeColor = Color.FromArgb(255, 238, 230, 246);
            facultyLabel.Focus();
            showTop(sender, e);
        }

        void buttonDown_Click(object sender, EventArgs e)
        {
            showBottom(sender, e);
        }

        void buttonUp_Click(object sender, EventArgs e)
        {
            showTop(sender, e);
        }

        void showBottom(object sender, EventArgs e)
        {
            buttonUp.Visible = true;
            buttonDown.Visible = false;
            titlepagePanel.Visible = false;
            DownPanel.Visible = true;
        }

        void showTop(object sender, EventArgs e)
        {
            buttonUp.Visible = false;
            buttonDown.Visible = true;
            titlepagePanel.Visible = true;
            DownPanel.Visible = false;
        }

        void buttonText_Click(object sender, EventArgs e)
        {
            if (buttonText.Text == "К тексту")
            {
                elementPanel.Visible = false;
                buttonText.Text = "К подстановкам";
                elementLabel.Text = "текст";
                richTextBox.Text = text;
                buttonSpecialH1.Visible = true;
                buttonSpecialH2.Visible = true;
                buttonSpecialL.Visible = true;
                buttonSpecialP.Visible = true;
            }
            else
            {
                elementPanel.Visible = true;
                buttonText.Text = "К тексту";
                elementLabel.Text = "нечто";
                richTextBox.Text = string.Empty;
                buttonSpecialH1.Visible = false;
                buttonSpecialH2.Visible = false;
                buttonSpecialL.Visible = false;
                buttonSpecialP.Visible = false;
            }
        }

        void AddToComboBox(ComboBox comboBox, string element)
        {
            if (!comboBox.Items.Contains(element))
            {
                comboBox.Items.Add(element);
                comboBox.SelectedIndex = comboBox.Items.IndexOf(element);
            }
        }

        void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            
            if (comboBox.SelectedIndex != -1)
            {
                for (int i = 4; i < 8; i++)
                {
                    ComboBox comboBoxToDeselect;
                    if (i != elementPanel.Controls.IndexOf(comboBox))
                    {
                        comboBoxToDeselect = (ComboBox)(elementPanel.Controls[i]);
                        comboBoxToDeselect.SelectedIndex = -1;
                    }
                }
                LStartText(sender);
                elementLabel.Text += (comboBox.Items.IndexOf(comboBox.SelectedItem) + 1).ToString();
                richTextBox.Text = comboBox.SelectedItem.ToString();
            }
            else
            {
                elementLabel.Text = "нечто";
                richTextBox.Text = string.Empty;
            }
        }

        void LStartText(object sender)
        {
            Control senderControl = (Control)sender;
            elementLabel.Text = elementPanel.Controls[elementPanel.Controls.IndexOf(senderControl) - 4].Text + ": ";
        }

        void ComboBox_MouseDown(object sender, MouseEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            if (e.Button == MouseButtons.Right)
            {
                if (Control.ModifierKeys != Keys.Shift && Control.ModifierKeys != Keys.Control && Control.ModifierKeys != Keys.Alt)
                {
                    for (int i = 4; i < 8; i++)
                    {
                        comboBox = (ComboBox)(elementPanel.Controls[i]);
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

        void richTextBox_TextChanged(object sender, EventArgs e)
        {
            if (elementLabel.Text != "нечто" && elementLabel.Text != "текст")
            {
                ComboBox comboBox = new ComboBox();
                comboBox.Visible = false;
                if (elementLabel.Text.StartsWith("Заголовок 1"))
                {
                    comboBox = heading1ComboBox;
                }
                else if (elementLabel.Text.StartsWith("Заголовок 2"))
                {
                    comboBox = heading2ComboBox;
                }
                else if (elementLabel.Text.StartsWith("Список"))
                {
                    comboBox = listComboBox;
                }
                else if (elementLabel.Text.StartsWith("Картинка"))
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
            if (elementLabel.Text == "текст")
            {
                text = richTextBox.Text;
            }
        }

        void buttonHeading1_Click(object sender, EventArgs e)
        {
            AddToComboBox(heading1ComboBox, richTextBox.Text);
        }

        void buttonHeading2_Click(object sender, EventArgs e)
        {
            AddToComboBox(heading2ComboBox, richTextBox.Text);
        }

        void buttonList_Click(object sender, EventArgs e)
        {
            AddToComboBox(listComboBox, richTextBox.Text);
        }

        void buttonPicture_Click(object sender, EventArgs e)
        {
            AddToComboBox(pictureComboBox, richTextBox.Text);
            // picture
        }

        void buttonSpecial_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            int cursorSave = richTextBox.SelectionStart;
            richTextBox.Text=richTextBox.Text.Insert(richTextBox.SelectionStart, "☺"+ button.Text.ToLower());
            richTextBox.Focus();
            richTextBox.SelectionStart = cursorSave+ button.Text.Length+1;
        }

        async void ReadScroll_Click(object sender, EventArgs e)
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
                UserInput userInput = new UserInput();
                userInput.ComboBoxH1 = DataComboBox(heading1ComboBox);
                userInput.ComboBoxH2 = DataComboBox(heading2ComboBox);
                userInput.ComboBoxL = DataComboBox(listComboBox);
                userInput.ComboBoxP = DataComboBox(pictureComboBox);
                userInput.Text = richTextBox.Text;
                await Task.Run(() => report.CreateReportLab(faculty, numberLab, theme, discipline, professor, year, userInput));
            }
            catch
            {
                MessageBox.Show("Что-то пошло не так :(", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }

            if (CloseWindow.Checked)
            {
                Application.Exit();
            }
        }

        string[] DataComboBox(ComboBox comboBox)
        {
            string[] dataComboBox = new string[comboBox.Items.Count];
            for (int i = 0; i < comboBox.Items.Count; i++)
            {
                dataComboBox[i] = comboBox.Items[i].ToString();
            }
            return dataComboBox;
        }

        void CloseWindow_Click(object sender, EventArgs e)
        {
            CloseWindow.Checked = !CloseWindow.Checked;
        }

        void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        void Save_Click(object sender, EventArgs e)
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

        void Open_Click(object sender, EventArgs e)
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

        void work_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem toolStripMenuItem = (ToolStripMenuItem)sender;
            if(toolStripMenuItem.Checked)
            {
                return;
            }
            
            if (toolStripMenuItem.Text == "Лабораторная")
            {
                this.Text = "Сотворение лабораторной работы из небытия";
            }
            else if (toolStripMenuItem.Text == "Практическая")
            {
                this.Text = "Сотворение практической работы из небытия";
            }
            else if (toolStripMenuItem.Text == "Курсовая")
            {
                this.Text = "Сотворение курсовой работы из небытия";
            }
            else if (toolStripMenuItem.Text == "Реферат")
            {
                this.Text = "Сотворение реферата из небытия";
            }
            else if (toolStripMenuItem.Text == "Диплом")
            {
                this.Text = "Сотворение дипломной работы из небытия";
            }
            else if (toolStripMenuItem.Text == "ВКР")
            {
                this.Text = "Сотворение ВКР из небытия";
            }
            else if (toolStripMenuItem.Text == "РГР")
            {
                this.Text = "Сотворение РГР из небытия";
            }
            Lab.Checked = false;
            Practice.Checked = false;
            Kurs.Checked = false;
            Ref.Checked = false;
            Diplom.Checked = false;
            VKR.Checked = false;
            RGR.Checked = false;
            toolStripMenuItem.Checked = true;
        }
    }
}
