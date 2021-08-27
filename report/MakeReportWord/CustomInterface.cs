using System;
using System.Drawing;
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
        string[] fileNames;
        CreatedElements createdElements;

        public CustomInterface()
        {
            InitializeComponent();
            if (Lab.Checked)
            {
                this.Text = "Сотворение лабораторной работы из небытия";
            }
            else if (Practice.Checked)
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
            replaceMenu();
            facultyLabel.Focus();
            menuLeftIndex = 1;
            HiddenElements(SubstitutionMenuItem);
            ShowElements(TitlePageMenuItem);
            createdElements = new CreatedElements();
        }

        void buttonText_Click(object sender, EventArgs e)
        {
            if (elementLabel.Text != "текст")
            {
                HiddenElements(SubstitutionMenuItem);
                ShowElements(TextMenuItem);
            }
            else
            {
                HiddenElements(TextMenuItem);
                ShowElements(SubstitutionMenuItem);
            }
        }

        void richTextBox_TextChanged(object sender, EventArgs e)
        {
            if (elementLabel.Text != "нечто" && elementLabel.Text != "текст")
            {
                int index = StringArraySearcher.IndexOf(elementLabel.Text.Split(':')[0], menuLabels);
                ComboBox comboBox = (ComboBox)(elementPanel.Controls[index + 1 + (elementPanel.ColumnCount - 2)]);
                if (comboBox.Visible)
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
            UpdateTypeButton();
        }

        void UpdateTypeButton()
        {
            if (DownPanelMI == TextMenuItem && createdElements.sum() > 0)
            {
                tableTypeInserts.Visible = true;
                CountTypeText(createdElements.H1, "h1");
                CountTypeText(createdElements.H2, "h2");
                CountTypeText(createdElements.L, "l");
                CountTypeText(createdElements.P, "p");
                CountTypeText(createdElements.T, "t");
                CountTypeText(createdElements.C, "c");
            }
        }

        void CountTypeText(int countCreatedElements, string str)
        {
            if (countCreatedElements <= (richTextBox.Text.Length - richTextBox.Text.Replace(("☺" + str), "").Length) / (str.Length+1))
            {
                tableTypeInserts.Controls[str.ToUpper()].Visible = false;
            }
            else
            {
                tableTypeInserts.Controls[str.ToUpper()].Visible = true;
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

        void buttonTable_Click(object sender, EventArgs e)
        {
            AddToComboBox(tableComboBox, richTextBox.Text);
        }

        void buttonCode_Click(object sender, EventArgs e)
        {
            AddToComboBox(codeComboBox, richTextBox.Text);
        }

        void buttonSpecial_Click(object sender, EventArgs e)
        {
            PictureBox button = (PictureBox)sender;
            int cursorSave = richTextBox.SelectionStart;
            if(richTextBox.Text.Length>0 && richTextBox.Text[cursorSave-1] == '☺')
            {
                richTextBox.Text = richTextBox.Text.Insert(richTextBox.SelectionStart, button.Name.ToLower());
                richTextBox.Focus();
                richTextBox.SelectionStart = cursorSave + button.Name.Length;
            }
            else
            {
                richTextBox.Text = richTextBox.Text.Insert(richTextBox.SelectionStart, "☺" + button.Name.ToLower());
                richTextBox.Focus();
                richTextBox.SelectionStart = cursorSave + button.Name.Length + 1;
            }
        }

        void buttonForward_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonForward.BackgroundImage = Properties.Resources.arrowsRightSelected;
                if (menuLeftIndex != elementPanel.ColumnStyles.Count - 1 - 4)
                {
                    menuLeftIndex++;
                }
                refreshMenu();
            }
        }

        void buttonBack_MouseUp(object sender, MouseEventArgs e)
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

        void buttonUp_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonUp.BackgroundImage = Properties.Resources.arrowsUpSelected;
                HiddenElements(DownPanelMI);
                ShowElements(TitlePageMenuItem);
            }
        }

        void buttonDown_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonDown.BackgroundImage = Properties.Resources.arrowsDownSelected;
                HiddenElements(TitlePageMenuItem);
                ShowElements(DownPanelMI);
            }
        }

        void menuExort_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.exortSelected;
                createTemplate(sender);
            }
        }

        void menuWex_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.wexSelected;
                createTemplate(sender);
            }
        }

        void menuQuas_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.quasSelected;
                createTemplate(sender);
            }
        }

        void menuButtonPB_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.BtnSelected;
                if (element.Name == "Добавить")
                {
                    ComboBox comboBox = (ComboBox)(elementPanel.Controls[elementPanel.Controls.IndexOf(element) - (elementPanel.ColumnCount - 2)]);
                    AddToComboBox(comboBox, richTextBox.Text);
                    createdElements.Add(comboBox.Name);
                }
                else if (element.Name == "КнопкаТекст")
                {
                    buttonText_Click(sender, e);
                }
                else if (element.Name == "H1"|| element.Name == "H2"|| element.Name == "L"|| element.Name == "P"|| element.Name == "T"|| element.Name == "C")
                {
                    buttonSpecial_Click(sender, e);
                }
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
                for (int i = elementPanel.ColumnCount - 1; i < elementPanel.ColumnCount * 2 - 4; i++)
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
                    for (int i = elementPanel.ColumnCount-1;i<2*elementPanel.ColumnCount-1-2;i++)
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
                    createdElements.Del(comboBox.Name);
                }
            }
        }

        async void ReadScroll_Click(object sender, EventArgs e)
        {
            MakeReport report = new MakeReport();
            string faculty = facultyComboBox.Text;
            string numberLab = numberLabTextBox.Text;
            string theme = themeTextBox.Text;
            string discipline = disciplineTextBox.Text;
            string professor = professorComboBox.Text;
            string year = yearTextBox.Text;
            try
            {
                UserInput userInput = new UserInput();
                userInput.ComboBoxH1 = DataComboBox(heading1ComboBox);
                userInput.ComboBoxH2 = DataComboBox(heading2ComboBox);
                userInput.ComboBoxL = DataComboBox(listComboBox);
                userInput.ComboBoxP = DataComboBox(pictureComboBox);
                userInput.ComboBoxT = DataComboBox(tableComboBox);
                userInput.ComboBoxC = DataComboBox(codeComboBox);
                if(TextMenuItem.Checked)
                {
                    userInput.Text = richTextBox.Text;
                }
                else
                {
                    userInput.Text = text;
                }
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

        void View_MenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickMenuItem = (ToolStripMenuItem)sender;
            if (TitlePageMenuItem.Checked)
            {
                HiddenElements(TitlePageMenuItem);
                ShowElements(ClickMenuItem);
            }
            else if (SubstitutionMenuItem.Checked)
            {
                HiddenElements(SubstitutionMenuItem);
                ShowElements(ClickMenuItem);
            }
            else if (TextMenuItem.Checked)
            {
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
                DownPanelMI = SubstitutionMenuItem;
            }
            else if (MenuItem == TextMenuItem)
            {
                tableTypeInserts.Visible = false;
                DownPanelMI = TextMenuItem;
            }
            MenuItem.Checked = false;
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
                buttonText.Text = "К тексту";
                textPicturePanel.ColumnStyles[0].Width = 60;
                textPicturePanel.ColumnStyles[1].Width = 40;
                DownPanelMI = SubstitutionMenuItem;
                richTextBox.Text = string.Empty;
            }
            else if (MenuItem == TextMenuItem)
            {
                buttonUp.Visible = true;
                DownPanel.Visible = true;
                buttonText.Text = "К подстановкам";
                richTextBox.Text = text;
                elementLabel.Text = "текст";
                textPicturePanel.ColumnStyles[1].Width = 0;
                textPicturePanel.ColumnStyles[0].Width = 100;
                DownPanelMI = TextMenuItem;
                UpdateTypeButton();
            }
            MenuItem.Checked = true;
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

        void hideMenuButtons()
        {
            for (int i = 0; i < elementPanel.Controls.Count; i++)
            {
                elementPanel.Controls[i].Visible = true;
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

        void DragNDropPanel_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        void DragNDropPanel_DragDrop(object sender, DragEventArgs e)
        {
            var data = e.Data.GetData(DataFormats.FileDrop);
            if (data != null)
            {
                fileNames = data as string[];
                if (fileNames.Length > 0)
                {
                    Point controlRelatedCoords = this.DragNDropPanel.PointToClient(new Point(e.X, e.Y));
                    if (controlRelatedCoords.X < 148)
                    {
                        richTextBox.Text = "☺p☺\n\n☺Содержимое☺\n";
                        pictureBox.BackgroundImage = Image.FromFile(fileNames[0]);
                        fileNames = null;
                    }
                    else
                    {
                        pictureBox.Refresh();
                    }
                }
            }
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

        void CustomInterface_DragEnter(object sender, DragEventArgs e)
        {
            fileNames = null;
            pictureBox.BackgroundImage = Properties.Resources.pictureCode;
        }

        void CustomInterface_DragLeave(object sender, EventArgs e)
        {
            pictureBox.BackgroundImage = Properties.Resources.DragNDrop;
        }

        void replaceMenuSpecial()
        {
            tableTypeInserts.Controls.Clear();
            tableTypeInserts.Controls.Add(GetMenuButtonReplacement(1, "H1")[0]);
            tableTypeInserts.Controls.Add(GetMenuButtonReplacement(1, "H2")[0]);
            tableTypeInserts.Controls.Add(GetMenuButtonReplacement(1, "L")[0]);
            tableTypeInserts.Controls.Add(GetMenuButtonReplacement(1, "P")[0]);
            tableTypeInserts.Controls.Add(GetMenuButtonReplacement(1, "T")[0]);
            tableTypeInserts.Controls.Add(GetMenuButtonReplacement(1, "C")[0]);
        }

        void replaceMenu()
        {
            globalFont.SetFont(heading1Label.Font, heading1Label.Font.Style);
            PictureBox[] menuPBarray = GetMenuLabelReplacement(elementPanel.ColumnCount - 2);
            PictureBox[] menuAddPBarray = GetMenuButtonReplacement(elementPanel.ColumnCount - 2, "Добавить");
            Control[] downPanelSave = new Control[DownPanel.Controls.Count];
            for (int i = 0; i < DownPanel.Controls.Count; i++)
            {
                downPanelSave[i] = DownPanel.Controls[i];
            }
            DownPanel.Controls.Clear();
            for (int i= 0; i < downPanelSave.Length; i++)
            {
                if (downPanelSave[i].Name == "buttonText")
                {
                    DownPanel.Controls.Add(GetMenuTextBtnReplacement(), 4, 1);
                }
                else
                {
                    DownPanel.Controls.Add(downPanelSave[i]);
                }
            }
            replaceMenuSpecial();
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
                elementPanel.Controls[i + 1].Name = menuLabels[i];
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

        bool MouseIsOverControl(PictureBox pb) => pb.ClientRectangle.Contains(pb.PointToClient(Cursor.Position));

        PictureBox[] GetMenuButtonReplacement(int amount, string name)
        {
            PictureBox[] menuPBarray = new PictureBox[amount];
            for (int menuPBindex = 0; menuPBindex < amount; menuPBindex++)
            {
                menuPBarray[menuPBindex] = new PictureBox();
                menuPBarray[menuPBindex].Dock = DockStyle.Fill;
                //menuPBarray[menuPBindex].TabIndex = 15 + menuPBindex;
                menuPBarray[menuPBindex].BackgroundImageLayout = ImageLayout.Stretch;
                menuPBarray[menuPBindex].BackgroundImage = Properties.Resources.Btn;
                menuPBarray[menuPBindex].MouseDown += menuButtonPB_MouseDown;
                menuPBarray[menuPBindex].MouseUp += menuButtonPB_MouseUp;
                menuPBarray[menuPBindex].MouseEnter += menuButtonPB_MouseEnter;
                menuPBarray[menuPBindex].MouseLeave += menuButtonPB_MouseLeave;
                menuPBarray[menuPBindex].Paint += menuPBbtn_Paint;
                menuPBarray[menuPBindex].Name = name;
            }
            return menuPBarray;
        }

        PictureBox[] GetMenuLabelReplacement(int amount)
        {
            PictureBox[] menuPBarray = new PictureBox[amount];
            for (int menuPBindex = 0; menuPBindex < amount; menuPBindex++)
            {
                menuPBarray[menuPBindex] = new PictureBox();
                menuPBarray[menuPBindex].Dock = DockStyle.Fill;
                //menuPBarray[menuPBindex].TabIndex = 15 + menuPBindex;
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

        PictureBox GetMenuTextBtnReplacement()
        {
            PictureBox pictureBox = new PictureBox();
            pictureBox.Dock = DockStyle.Fill;
            pictureBox.BackgroundImageLayout = ImageLayout.Stretch;
            pictureBox.Name = "КнопкаТекст";
            pictureBox.BackgroundImage = Properties.Resources.Btn;
            pictureBox.MouseDown += menuButtonPB_MouseDown;
            pictureBox.MouseUp += menuButtonPB_MouseUp;
            pictureBox.MouseEnter += menuButtonPB_MouseEnter;
            pictureBox.MouseLeave += menuButtonPB_MouseLeave;
            pictureBox.Paint += buttonText_Paint;

            return pictureBox;
        }

        void createTemplate(object sender)
        {
            Control control = (Control)sender;
            if(control.Name == "Заголовок 1")
            {
                richTextBox.Text = "☺h1☺\n\n☺Содержимое☺\n";
            }
            else if(control.Name == "Заголовок 2")
            {
                richTextBox.Text = "☺h2☺\n\n☺Содержимое☺\n";
            }
            else if (control.Name == "Список")
            {
                richTextBox.Text = "☺l☺\n\n☺Содержимое☺\n";
            }
            else if (control.Name == "Картинка")
            {
                richTextBox.Text = "☺p☺\n\n☺Содержимое☺\n";
            }
            else if (control.Name == "Таблица")
            {
                richTextBox.Text = "☺t☺\n\n☺Содержимое☺\n";
            }
            else if (control.Name == "Код")
            {
                richTextBox.Text = "☺c☺\n\n☺Содержимое☺\n";
            }
        }

        void facultyComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str = string.Empty;
            professorComboBox.Items.Clear();
            if (facultyComboBox.SelectedIndex == 0)
            {
                str = "Амелина О.В.!Артёмов А.В.!Валухов В.А.!Волков В.Н.!Гордиенко А.П.!Демидов А.В.!Захарова О.В.!Константинов И.С.!Конюхова О.В.!Кравцова Э.А.!Лукьянов П.В.!Преснецова В.Ю.!Раков В.И.!Рыженков Д.В.!Санников Д.П.!Селихов А.В.!Стычук А.А.!Ужаринский А.Ю.!Фролов А.И.!Чижов А.В.!Шатеев Р.В.";
            }
            else if (facultyComboBox.SelectedIndex == 1)
            {
                str = "Амелина О.В.!Артёмов А.В.!Батищев А.В.!Биктимиров М.Р.!Волков В.Н.!Демидов А.В.!Загородних Н.А.!Закалкина Е.В.!Корнаева Е.П.!Коськин А.В.!Кравцова Э.А.!Красуля О.А.!Ларина Л.Ю.!Лунёв Р.А.!Лыськов О.Э.!Машкова А.Л.!Митин А.А.!Новиков С.В.!Олькина Е.В.!Преснецова В.Ю.!Рыженков Д.В.!Савина О.А.!Сезонов Д.С.!Соков О.А.!Строев С.П.!Стычук А.А.!Терентьев С.В.!Федоров Г.Д.!Фроленкова Л.Ю.!Фролова В.А.";
            }
            else if (facultyComboBox.SelectedIndex == 2)
            {
                str = "Бондарева Л.А.!Дунаев А.В.!Жидков А.В.!Жильцов М.П.!Козлова Л.Д.!Крутикова В.Ю.!Маковик И.Н.!Марков В.В.!Незнанов А.И.!Подмастерьев К.В.!Потапова Е.В.!Секаева Ж.А.!Селихов А.В.!Семёнов В.В.!Сковпень В.Н.!Углова Н.В.!Яковенко М.В.";
            }
            else if (facultyComboBox.SelectedIndex == 3)
            {
                str = "Ерёменко В.Т.!Мишин Д.С.!Пеньков Н.Г.!Савва Ю.Б.!Фисенко В.Е.!Фисун А.П.";
            }
            else if (facultyComboBox.SelectedIndex == 4)
            {
                str = "Батуров Д.П.!Бурлакова Е.А.!Гордон В.А.!Кирсанова О.В.!Матюхин С.И.!Потураева Т.В.!Ромашин С.Н.!Семёнова Г.А.!Фроленкова Л.Ю.!Шоркин В.С.!Якушина С.И.";
            }
            else if (facultyComboBox.SelectedIndex == 5)
            {
                str = "Аксёнов К.В.!Багров В.В.!Батенков А.А.!Варгашкин В.Я.!Воронина О.А.!Донцов В.М.!Игнатов Ю.В.!Косчинский С.Л.!Лобанова В.А.!Лобода О.А.!Майоров М.В.!Мишин В.В.!Моисеев П.П.!Муравьёв А.А.!Плащенков Д.А.!Рязанцев П.Н.!Селихов А.В.!Сковпень В.Н.!Суздальцев А.И.!Тарарака А.В.!Тугарев А.С.!Тютякин А.В.!Филина А.В.!Шишкин А.А.";
            }
            else if (facultyComboBox.SelectedIndex == 6)
            {
                str = "Качанов А.Н.!Комаристый А.С.!Коренков Д.А.!Королева Т.Г.!Петров Г.Н.!Токарев А.М.!Филина А.В.!Харитонова Л.Г.!Чернышов В.А.";
            }
            foreach (string s in str.Split('!'))
            {
                professorComboBox.Items.Add(s);
            }
        }
    }
}

