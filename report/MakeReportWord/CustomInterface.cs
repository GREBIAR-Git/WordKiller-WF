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
        int menuLeftIndex;
        string[] menuLabels;
        ToolStripMenuItem DownPanelMI;

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
            replaceMenu();
            facultyLabel.Focus();
            menuLeftIndex = 1;
            DownPanelMI = SubstitutionMenuItem;
            HiddenElements(SubstitutionMenuItem);
            ShowElements(TitlePageMenuItem);
        }

        void buttonText_Click(object sender, EventArgs e)
        {
            if (buttonText.Text == "К тексту")
            {
                HiddenElements(SubstitutionMenuItem);
                ShowElements(TextMenuItem);
                DownPanelMI = TextMenuItem;
                TextMenuItem.Checked = true;
                SubstitutionMenuItem.Checked = false;
            }
            else
            {
                HiddenElements(TextMenuItem);
                ShowElements(SubstitutionMenuItem);
                DownPanelMI = SubstitutionMenuItem;
                SubstitutionMenuItem.Checked = true;
                TextMenuItem.Checked = false;
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
                for (int i = elementPanel.ColumnCount - 1; i < elementPanel.ColumnCount * 2 - 1 - 3; i++)
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
            int i = elementPanel.Controls.IndexOf(senderControl) - 1 - (elementPanel.ColumnCount - 2);
            elementLabel.Text = menuLabels[i] + ": ";
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
                int index = StringArraySearcher.IndexOf(elementLabel.Text.Split(':')[0], menuLabels);
                comboBox = (ComboBox)(elementPanel.Controls[index + 1 + (elementPanel.ColumnCount - 2)]);
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

                output.WriteLine("facultyComboBox=" + facultyComboBox.SelectedItem.ToString());
                output.WriteLine("numberLabTextBox.Text=" + numberLabTextBox.Text);
                output.WriteLine("themeTextBox.Text=" + themeTextBox.Text);
                output.WriteLine("disciplineTextBox.Text=" + disciplineTextBox.Text);
                output.WriteLine("professorTextBox.Text=" + professorTextBox.Text);
                output.WriteLine("yearTextBox.Text=" + yearTextBox.Text);

                for (int i = 0; i < heading1ComboBox.Items.Count; i++)
                {
                    output.WriteLine("heading1ComboBox.Items[" + i+"]=" + heading1ComboBox.Items[i].ToString());
                }
                for (int i = 0; i < heading2ComboBox.Items.Count; i++)
                {
                    output.WriteLine("heading2ComboBox.Items[" + i + "]=" + heading2ComboBox.Items[i].ToString());
                }
                for (int i = 0; i < listComboBox.Items.Count; i++)
                {
                    output.WriteLine("listComboBox.Items[" + i + "]=" + listComboBox.Items[i].ToString());
                }
                for (int i = 0; i < pictureComboBox.Items.Count; i++)
                {
                    output.WriteLine("pictureComboBox.Items[" + i + "]=" + pictureComboBox.Items[i].ToString());
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

                    bool readingText = false;
                    for (int i = 0; i < lines.Length; i++)
                    {
                        if (lines[i].StartsWith("###textstart"))
                        {
                            readingText = true;
                        }
                        else if (readingText)
                        {
                            if (lines[i].StartsWith("###textend"))
                            {
                                readingText = false;
                            }
                            else
                            {
                                text += lines[i] + "\n";
                            }
                        }
                        else
                        {
                            string[] variable_value = lines[i].Split('=');
                            if (variable_value.Length == 2)
                            {
                                if (variable_value[0].StartsWith("facultyComboBox"))
                                {
                                    facultyComboBox.SelectedItem=variable_value[1];
                                }
                                else if (variable_value[0].StartsWith("numberLabTextBox.Text"))
                                {
                                    numberLabTextBox.Text = variable_value[1];
                                }
                                else if (variable_value[0].StartsWith("themeTextBox.Text"))
                                {
                                    themeTextBox.Text = variable_value[1];
                                }
                                else if (variable_value[0].StartsWith("disciplineTextBox.Text"))
                                {
                                    disciplineTextBox.Text = variable_value[1];
                                }
                                else if (variable_value[0].StartsWith("professorTextBox.Text"))
                                {
                                    professorTextBox.Text = variable_value[1];
                                }
                                else if (variable_value[0].StartsWith("yearTextBox.Text"))
                                {
                                    yearTextBox.Text = variable_value[1];
                                }
                                else if (variable_value[0].StartsWith("heading1ComboBox"))
                                {
                                    heading1ComboBox.Items.Add(variable_value[1]);
                                }
                                else if (variable_value[0].StartsWith("heading2ComboBox"))
                                {
                                    heading2ComboBox.Items.Add(variable_value[1]);
                                }
                                else if (variable_value[0].StartsWith("listComboBox"))
                                {
                                    listComboBox.Items.Add(variable_value[1]);
                                }
                                else if (variable_value[0].StartsWith("pictureComboBox"))
                                {
                                    pictureComboBox.Items.Add(variable_value[1]);
                                }
                            }
                        }
                    }
                    
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

        void hideMenuButtons()
        {
            for (int i = 0; i < elementPanel.Controls.Count; i++)
            {
                elementPanel.Controls[i].Visible = true;
            }
            for (int i = 0; i < elementPanel.Controls.Count; i++)
            {
                if (elementPanel.Controls[i].Width == 1)
                {
                    elementPanel.Controls[i].Visible = false;
                }
            }
        }

        void refreshMenuArrows()
        {
            if (menuLeftIndex == 1)
            {
                buttonBack.Visible = false;
            }
            else
            {
                buttonBack.Visible = true;
            }
            if (menuLeftIndex == elementPanel.ColumnStyles.Count - 1 - 4)
            {
                buttonForward.Visible = false;
            }
            else
            {
                buttonForward.Visible = true;
            }
        }

        void refreshMenu()
        {
            elementPanel.SuspendLayout();
            for (int i = 0; i < elementPanel.ColumnStyles.Count - 1; i++)
            {
                elementPanel.ColumnStyles[i].SizeType = SizeType.Percent;
                elementPanel.ColumnStyles[i].Width = 0;
            }
            elementPanel.ColumnStyles[0].Width = 6;
            elementPanel.ColumnStyles[elementPanel.ColumnStyles.Count - 1].Width = 6;

            for (int i = menuLeftIndex; i < menuLeftIndex + 4; i++)
            {
                elementPanel.ColumnStyles[i].Width = 22;
            }
            elementPanel.ResumeLayout();
            hideMenuButtons();
            refreshMenuArrows();
        }

        void DragNDropPanel_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        void DragNDropPanel_DragDrop(object sender, DragEventArgs e)
        {
            var data = e.Data.GetData(DataFormats.FileDrop);
            if (data != null)
            {
                var fileNames = data as string[];
                if (fileNames.Length > 0)
                {
                    Point controlRelatedCoords = this.DragNDropPanel.PointToClient(new Point(e.X, e.Y));
                    if (controlRelatedCoords.X < 148)
                    {
                        pictureBox.BackgroundImage = Image.FromFile(fileNames[0]);
                    }
                    else
                    {
                        Graphics g = pictureBox.CreateGraphics();
                        g.DrawImage(Properties.Resources.Code, 0, 0, pictureBox.Width, pictureBox.Height);
                        richTextBox.Text = fileNames[0];
                        string nameFile = fileNames[0].Split('\\')[fileNames[0].Split('\\').Length-1];
                        g.DrawString(nameFile, new Font("Microsoft Sans Serif", 14), new SolidBrush(Color.Black), new Point(20, pictureBox.Height/2+30));
                        
                    }
                }
            }
        }

        void CustomInterface_DragEnter(object sender, DragEventArgs e)
        {
            pictureBox.BackgroundImage = Properties.Resources.pictureCode;
        }

        void CustomInterface_DragLeave(object sender, EventArgs e)
        {
            pictureBox.BackgroundImage = Properties.Resources.DragNDrop;
        }

        void DragNDropPanel_DragOver(object sender, DragEventArgs e)
        {
            Point controlRelatedCoords = this.DragNDropPanel.PointToClient(new Point(e.X, e.Y));
            if (controlRelatedCoords.X < 148)
            {
                pictureBox.BackgroundImage = Properties.Resources.pictureCode_Picture;
            }
            else
            {
                pictureBox.BackgroundImage = Properties.Resources.pictureCode_Code;
            }
        }

        void View_MenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickMenuItem = (ToolStripMenuItem)sender;
            if (TitlePageMenuItem.Checked)
            {
                TitlePageMenuItem.Checked = false;
                ClickMenuItem.Checked = true;
                HiddenElements(TitlePageMenuItem);
                ShowElements(ClickMenuItem);
            }
            else if (SubstitutionMenuItem.Checked)
            {
                SubstitutionMenuItem.Checked = false;
                ClickMenuItem.Checked = true;
                HiddenElements(SubstitutionMenuItem);
                ShowElements(ClickMenuItem);
            }
            else if (TextMenuItem.Checked)
            {
                TextMenuItem.Checked = false;
                ClickMenuItem.Checked = true;
                HiddenElements(TextMenuItem);
                ShowElements(ClickMenuItem);
            }
        }

        void HiddenElements(ToolStripMenuItem MenuItem)
        {
            if (MenuItem == TitlePageMenuItem)
            {
                buttonDown.Visible = false;
                titlepagePanel.Visible = false;
                refreshMenu();
            }
            else if (MenuItem == SubstitutionMenuItem)
            {
                elementPanel.Visible = false;
                pictureBox.Visible = false;
            }
            else if (MenuItem == TextMenuItem)
            {
                tableLayoutPanel1.Visible = false;
            }
        }

        void ShowElements(ToolStripMenuItem MenuItem)
        {
            if (MenuItem == TitlePageMenuItem)
            {
                buttonDown.Visible = true;
                titlepagePanel.Visible = true;
                DownPanel.Visible = false;
                buttonUp.Visible = false;
            }
            else if (MenuItem == SubstitutionMenuItem)
            {
                buttonUp.Visible = true;
                DownPanel.Visible = true;
                pictureBox.Visible = true;
                elementPanel.Visible = true;
                elementLabel.Text = "нечто";
                richTextBox.Text = string.Empty;
                buttonText.Text = "К тексту";
                textPicturePanel.ColumnStyles[0].Width = 60;
                textPicturePanel.ColumnStyles[1].Width = 40;
                DownPanelMI = SubstitutionMenuItem;
            }
            else if (MenuItem == TextMenuItem)
            {
                buttonUp.Visible = true;
                DownPanel.Visible = true;
                tableLayoutPanel1.Visible = true;
                buttonText.Text = "К подстановкам";
                richTextBox.Text = text;
                elementLabel.Text = "текст";
                textPicturePanel.ColumnStyles[1].Width = 0;
                textPicturePanel.ColumnStyles[0].Width = 100;
                DownPanelMI = TextMenuItem;
            }
        }

        private void buttonForward_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonForward.BackgroundImage = Properties.Resources.arrowsRightPressed;
            }
        }

        private void buttonBack_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonBack.BackgroundImage = Properties.Resources.arrowsLeftPressed;
            }
        }

        private void buttonForward_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonForward.BackgroundImage = Properties.Resources.arrowsRightSelected;
                if (menuLeftIndex != elementPanel.ColumnStyles.Count - 2 + 1 - 4)
                {
                    menuLeftIndex++;
                }
                refreshMenu();
            }
        }

        private void buttonBack_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonBack.BackgroundImage = Properties.Resources.arrowsLeftSelected;
                if (menuLeftIndex != 1)
                {
                    menuLeftIndex--;
                }
                refreshMenu();
            }
        }

        private void buttonForward_MouseEnter(object sender, EventArgs e)
        {
            buttonForward.BackgroundImage = Properties.Resources.arrowsRightSelected;
        }

        private void buttonForward_MouseLeave(object sender, EventArgs e)
        {
            buttonForward.BackgroundImage = Properties.Resources.arrowsRight;
        }

        private void buttonBack_MouseEnter(object sender, EventArgs e)
        {
            buttonBack.BackgroundImage = Properties.Resources.arrowsLeftSelected;
        }

        private void buttonBack_MouseLeave(object sender, EventArgs e)
        {
            buttonBack.BackgroundImage = Properties.Resources.arrowsLeft;
        }

        private void buttonUp_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonUp.BackgroundImage = Properties.Resources.arrowsUpPressed;
            }

        }

        private void buttonUp_MouseEnter(object sender, EventArgs e)
        {
            buttonUp.BackgroundImage = Properties.Resources.arrowsUpSelected;
        }

        private void buttonUp_MouseLeave(object sender, EventArgs e)
        {
            buttonUp.BackgroundImage = Properties.Resources.arrowsUp;
        }

        private void buttonUp_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonUp.BackgroundImage = Properties.Resources.arrowsUpSelected;
                HiddenElements(DownPanelMI);
                ShowElements(TitlePageMenuItem);
                TitlePageMenuItem.Checked = true;
                DownPanelMI.Checked = false;
            }
        }

        private void buttonDown_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonDown.BackgroundImage = Properties.Resources.arrowsDownPressed;
            }
        }

        private void buttonDown_MouseEnter(object sender, EventArgs e)
        {
            buttonDown.BackgroundImage = Properties.Resources.arrowsDownSelected;
        }

        private void buttonDown_MouseLeave(object sender, EventArgs e)
        {
            buttonDown.BackgroundImage = Properties.Resources.arrowsDown;
        }

        private void buttonDown_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonDown.BackgroundImage = Properties.Resources.arrowsDownSelected;
                HiddenElements(TitlePageMenuItem);
                ShowElements(DownPanelMI);
                TitlePageMenuItem.Checked = false;
                DownPanelMI.Checked = true;
            }
        }

        private void replaceMenu()
        {
            globalFont.SetFont(heading1Label.Font, heading1Label.Font.Style);
            PictureBox[] menuPBarray = GetMenuLabelReplacement(elementPanel.ColumnCount - 2);
            PictureBox[] menuAddPBarray = GetMenuButtonReplacement(elementPanel.ColumnCount - 2);
            Control[] controlsSave = new Control[elementPanel.Controls.Count];
            for (int i = 0; i < elementPanel.Controls.Count; i++)
            {
                controlsSave[i] = elementPanel.Controls[i];
            }
            elementPanel.Controls.Clear();
            elementPanel.Controls.Add(controlsSave[0]);
            menuLabels = new string[menuPBarray.Length];
            for (int i = 0; i < menuPBarray.Length; i++)
            {
                elementPanel.Controls.Add(menuPBarray[i], i + 1, 0);
                menuLabels[i] = controlsSave[i + 1].Text;
            }
            for (int i = 0; i < elementPanel.ColumnCount - 2; i++)
            {
                elementPanel.Controls.Add(controlsSave[elementPanel.ColumnCount - 1 + i]);
            }
            for (int i = 0; i < menuAddPBarray.Length; i++)
            {
                elementPanel.Controls.Add(menuAddPBarray[i], i + 1, 2);
            }
            elementPanel.Controls.Add(controlsSave[controlsSave.Length - 1]);
        }
        private bool MouseIsOverControl(PictureBox pb) => pb.ClientRectangle.Contains(pb.PointToClient(Cursor.Position));

        private void menuPB_Paint(object sender, PaintEventArgs e)
        {
            PictureBox pb = (PictureBox)sender;
            Control controlPB = (Control)sender;
            int index = elementPanel.Controls.IndexOf(controlPB);
            Font fnt = globalFont.GetFont();

            if (MouseIsOverControl(pb) && Control.MouseButtons != MouseButtons.Left)
            {
                fnt = new Font(fnt.Name, 16);
            }
            else
            {
                fnt = new Font(fnt.Name, 14);
            }
            using (fnt)
            {
                SizeF stringSize = e.Graphics.MeasureString(menuLabels[index - 1], fnt);
                e.Graphics.DrawString(menuLabels[index - 1], fnt, Brushes.Black, new Point((int)(pb.Width / 2 - stringSize.Width / 2), (int)(pb.Height / 2 - stringSize.Height / 2)));
            }
        }

        private void menuPBadd_Paint(object sender, PaintEventArgs e)
        {
            string str = "Добавить";
            PictureBox pb = (PictureBox)sender;
            Font fnt = globalFont.GetFont();

            if (MouseIsOverControl(pb) && Control.MouseButtons != MouseButtons.Left)
            {
                fnt = new Font(fnt.Name, 16);
            }
            else
            {
                fnt = new Font(fnt.Name, 14);
            }
            using (fnt)
            {
                SizeF stringSize = e.Graphics.MeasureString(str, fnt);
                e.Graphics.DrawString(str, fnt, Brushes.Black, new Point((int)(pb.Width / 2 - stringSize.Width / 2), (int)(pb.Height / 2 - stringSize.Height / 2)));
            }
        }

        private PictureBox[] GetMenuButtonReplacement(int amount)
        {
            PictureBox[] menuPBarray = new PictureBox[amount];
            for (int menuPBindex = 0; menuPBindex < amount; menuPBindex++)
            {
                menuPBarray[menuPBindex] = new PictureBox();
                menuPBarray[menuPBindex].Dock = DockStyle.Fill;
                //menuPBarray[menuPBindex].TabIndex = 15 + menuPBindex;
                menuPBarray[menuPBindex].BackgroundImageLayout = ImageLayout.Stretch;
                menuPBarray[menuPBindex].BackgroundImage = Properties.Resources.AddNormal;
                menuPBarray[menuPBindex].MouseDown += menuAddPB_MouseDown;
                menuPBarray[menuPBindex].MouseUp += menuAddPB_MouseUp;
                menuPBarray[menuPBindex].MouseEnter += menuAddPB_MouseEnter;
                menuPBarray[menuPBindex].MouseLeave += menuAddPB_MouseLeave;
                menuPBarray[menuPBindex].Paint += menuPBadd_Paint;
            }
            return menuPBarray;
        }

        private PictureBox[] GetMenuLabelReplacement(int amount)
        {
            PictureBox[] menuPBarray = new PictureBox[amount];
            for (int menuPBindex = 0; menuPBindex < amount; menuPBindex++)
            {
                menuPBarray[menuPBindex] = new PictureBox();
                menuPBarray[menuPBindex].Dock = DockStyle.Fill;
                menuPBarray[menuPBindex].TabIndex = 15 + menuPBindex;
                menuPBarray[menuPBindex].BackgroundImageLayout = ImageLayout.Stretch;
                if (menuPBindex < amount / 3)
                {
                    menuPBarray[menuPBindex].BackgroundImage = Properties.Resources.exort;
                    menuPBarray[menuPBindex].MouseDown += menuExort_MouseDown;
                    menuPBarray[menuPBindex].MouseUp += menuExort_MouseUp;
                    menuPBarray[menuPBindex].MouseEnter += menuExort_MouseEnter;
                    menuPBarray[menuPBindex].MouseLeave += menuExort_MouseLeave;
                }
                else if (menuPBindex < 2 * amount / 3)
                {
                    menuPBarray[menuPBindex].BackgroundImage = Properties.Resources.wex;
                    menuPBarray[menuPBindex].MouseDown += menuWex_MouseDown;
                    menuPBarray[menuPBindex].MouseUp += menuWex_MouseUp;
                    menuPBarray[menuPBindex].MouseEnter += menuWex_MouseEnter;
                    menuPBarray[menuPBindex].MouseLeave += menuWex_MouseLeave;
                }
                else
                {
                    menuPBarray[menuPBindex].BackgroundImage = Properties.Resources.quas;
                    menuPBarray[menuPBindex].MouseDown += menuQuas_MouseDown;
                    menuPBarray[menuPBindex].MouseUp += menuQuas_MouseUp;
                    menuPBarray[menuPBindex].MouseEnter += menuQuas_MouseEnter;
                    menuPBarray[menuPBindex].MouseLeave += menuQuas_MouseLeave;
                }
                menuPBarray[menuPBindex].Paint += menuPB_Paint;
            }
            return menuPBarray;
        }

        private void buttonTable_Click(object sender, EventArgs e)
        {
            AddToComboBox(tableComboBox, richTextBox.Text);
        }

        private void buttonCode_Click(object sender, EventArgs e)
        {
            AddToComboBox(codeComboBox, richTextBox.Text);
        }

        private void menuExort_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.exortPressed;
            }
        }

        private void menuWex_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.wexPressed;
            }
        }

        private void menuQuas_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.quasPressed;
            }
        }

        private void menuExort_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.exortSelected;
            }
        }

        private void menuWex_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.wexSelected;
            }
        }

        private void menuQuas_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.quasSelected;
            }
        }

        private void menuExort_MouseEnter(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.exortSelected;
        }

        private void menuWex_MouseEnter(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.wexSelected;
        }

        private void menuQuas_MouseEnter(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.quasSelected;
        }

        private void menuExort_MouseLeave(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.exort;
        }

        private void menuWex_MouseLeave(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.wex;
        }

        private void menuQuas_MouseLeave(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.quas;
        }

        void menuAddPB_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.AddPressed;
            }
        }

        void menuAddPB_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                PictureBox control = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.AddSelected;
                ComboBox comboBox = (ComboBox)(elementPanel.Controls[elementPanel.Controls.IndexOf(control) - (elementPanel.ColumnCount - 2)]);
                AddToComboBox(comboBox, richTextBox.Text);
            }
        }

        void menuAddPB_MouseEnter(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.AddSelected;
        }
        void menuAddPB_MouseLeave(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.AddNormal;
        }
    }
}

