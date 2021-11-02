using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MakeReportWord
{
    public partial class CustomInterface : Form
    {
        string text = string.Empty;
        string textDragOnDrop;
        int menuLeftIndex;
        string[] menuLabels;
        string fileNames;
        ToolStripMenuItem DownPanelMI;
        DataComboBox dataComboBox;
        WindowSize wndSize;
        public CustomInterface(string[] fileName)
        {
            InitializeComponent();
            dataComboBox = new DataComboBox();
            replaceMenu();
            menuLeftIndex = 1;
            wndSize = new WindowSize();
            DownPanelMI = SubstitutionMenuItem;
            SaveTitlePagePanelCells();
            DEFAULTtitlepagePanelControls = CopyControls(titlepagePanel, 0, titlepagePanel.Controls.Count - 1);
            if (DefaultDocumentMenuItem.Checked)
            {
                TextHeader("документа");
                TitleOffOnMenuItem.Visible = !DefaultDocumentMenuItem.Checked;
                ShowingTitelPanel();
            }
            else if (LabMenuItem.Checked)
            {
                TextHeader("лабораторной работы");
                ShowTitleElems("0.0 1.0 2.1 3.1 0.3 1.3 0.4 1.4 0.6 1.6 0.7 1.7");
            }
            else if (PracticeMenuItem.Checked)
            {
                TextHeader("практической работы");
                ShowTitleElems("0.0 1.0 2.1 3.1 0.3 1.3 0.4 1.4 0.6 1.6 0.7 1.7");
            }
            else if (KursMenuItem.Checked)
            {
                TextHeader("курсовой работы");
                ShowTitleElems("0.0 1.0 0.1 1.1 4.1 5.1 0.3 1.3 0.4 1.4 0.6 1.6 0.7 1.7");
            }
            else if (RefMenuItem.Checked)
            {
                TextHeader("реферата");
            }
            else if (DiplomMenuItem.Checked)
            {
                TextHeader("дипломной работы");
            }
            else if (VKRMenuItem.Checked)
            {
                TextHeader("ВКР");
            }
            else if (RGRMenuItem.Checked)
            {
                TextHeader("РГР");
            }
            if (fileName.Length > 0)
            {
                if (fileName[0].EndsWith(".wordkiller") && System.IO.File.Exists(fileName[0]))
                {
                    OpenWordKiller(fileName[0]);
                }
                else
                {
                    throw new Exception("Ошибка открытия файла:\nФайл не найден или формат не поддерживается");
                }
            }
        }


        void TextHeader(string type)
        {
            this.Text = "Сотворение " + type + " из небытия";
        }


        void buttonText_Click(object sender, EventArgs e)
        {
            if (DownPanelMI == SubstitutionMenuItem)
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
            if (DownPanelMI == SubstitutionMenuItem)
            {
                int cursorSave = richTextBox.SelectionStart;
                if (elementLabel.Text != "нечто" && ComboBoxSelected() && richTextBox.Text != string.Empty)
                {
                    if (h1ComboBox.SelectedIndex != -1)
                    {
                        dataComboBox.ComboBoxH1[h1ComboBox.SelectedIndex][0] = richTextBox.Text.Split('\n')[1];
                        dataComboBox.ComboBoxH1[h1ComboBox.SelectedIndex][1] = SplitMainText();
                        h1ComboBox.Items[h1ComboBox.SelectedIndex] = dataComboBox.ComboBoxH1[h1ComboBox.SelectedIndex][0];
                    }
                    else if (h2ComboBox.SelectedIndex != -1)
                    {
                        dataComboBox.ComboBoxH2[h2ComboBox.SelectedIndex][0] = richTextBox.Text.Split('\n')[1];
                        dataComboBox.ComboBoxH2[h2ComboBox.SelectedIndex][1] = SplitMainText();
                        h2ComboBox.Items[h2ComboBox.SelectedIndex] = dataComboBox.ComboBoxH2[h2ComboBox.SelectedIndex][0];
                    }
                    else if (lComboBox.SelectedIndex != -1)
                    {
                        dataComboBox.ComboBoxL[lComboBox.SelectedIndex][0] = richTextBox.Text.Split('\n')[1];
                        dataComboBox.ComboBoxL[lComboBox.SelectedIndex][1] = SplitMainText();
                        lComboBox.Items[lComboBox.SelectedIndex] = dataComboBox.ComboBoxL[lComboBox.SelectedIndex][0];
                    }
                    else if (pComboBox.SelectedIndex != -1)
                    {
                        dataComboBox.ComboBoxP[pComboBox.SelectedIndex][0] = richTextBox.Text.Split('\n')[1];
                        dataComboBox.ComboBoxP[pComboBox.SelectedIndex][1] = SplitMainText();
                        pComboBox.Items[pComboBox.SelectedIndex] = dataComboBox.ComboBoxP[pComboBox.SelectedIndex][0];
                        if (!System.IO.File.Exists(SplitMainText()))
                        {
                            fileNames = null;
                        }
                        else
                        {
                            fileNames = SplitMainText();
                        }
                    }
                    else if (tComboBox.SelectedIndex != -1)
                    {
                        dataComboBox.ComboBoxT[tComboBox.SelectedIndex][0] = richTextBox.Text.Split('\n')[1];
                        dataComboBox.ComboBoxT[tComboBox.SelectedIndex][1] = SplitMainText();
                        tComboBox.Items[tComboBox.SelectedIndex] = dataComboBox.ComboBoxT[tComboBox.SelectedIndex][0];
                    }
                    else if (cComboBox.SelectedIndex != -1)
                    {
                        dataComboBox.ComboBoxC[cComboBox.SelectedIndex][0] = richTextBox.Text.Split('\n')[1];
                        dataComboBox.ComboBoxC[cComboBox.SelectedIndex][1] = SplitMainText();
                        cComboBox.Items[cComboBox.SelectedIndex] = dataComboBox.ComboBoxC[cComboBox.SelectedIndex][0];
                        if (!System.IO.File.Exists(SplitMainText()))
                        {
                            fileNames = null;
                        }
                        else
                        {
                            fileNames = SplitMainText();
                        }
                    }
                }
                // снизу крыса
                byte[] bytes = Encoding.Unicode.GetBytes(richTextBox.Text);
                //richTextBox.Text = Encoding.Unicode.GetString(bytes);
                // сверху крыса



                richTextBox.SelectionStart = cursorSave;
                pictureBox.Refresh();

            }
            else if (DownPanelMI == TextMenuItem)
            {
                text = richTextBox.Text;
                UpdateTypeButton();
            }
        }


        void UpdateTypeButton()
        {
            if (dataComboBox.Sum()>0)
            {
                ShowSpecials();
                CountTypeText(dataComboBox.ComboBoxH1.Count, "h1");
                CountTypeText(dataComboBox.ComboBoxH2.Count, "h2");
                CountTypeText(dataComboBox.ComboBoxL.Count, "l");
                CountTypeText(dataComboBox.ComboBoxP.Count, "p");
                CountTypeText(dataComboBox.ComboBoxT.Count, "t");
                CountTypeText(dataComboBox.ComboBoxC.Count, "c");
            }
        }


        void CountTypeText(int countCreatedElements, string str)
        {
            if (countCreatedElements <= (richTextBox.Text.Length - richTextBox.Text.Replace(("☺" + str), "").Length) / (str.Length + 1))
            {
                tableTypeInserts.Controls[str.ToUpper()].Visible = false;
            }
            else
            {
                tableTypeInserts.Controls[str.ToUpper()].Visible = true;
            }
        }

        void buttonSpecial_Click(object sender, EventArgs e)
        {
            PictureBox button = (PictureBox)sender;
            int cursorSave = richTextBox.SelectionStart;
            if (richTextBox.Text.Length > 0 && cursorSave > 0 && richTextBox.Text[cursorSave - 1] == '☺')
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
                if (menuLeftIndex != elementPanel.ColumnStyles.Count - 5)
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
                if (element.Name == "Добавить" && ValidAddInput())
                {
                    string str = richTextBox.Text.Split('\n')[0];
                    // соединить 
                    if (str == "☺h1☺")
                    {
                        string[] strData = new string[] { richTextBox.Text.Split('\n')[1], SplitMainText() };
                        AddToComboBox(h1ComboBox, dataComboBox.ComboBoxH1, strData);
                    }
                    else if (str == "☺h2☺")
                    {
                        string[] strData = new string[] { richTextBox.Text.Split('\n')[1], SplitMainText() };
                        AddToComboBox(h2ComboBox, dataComboBox.ComboBoxH2, strData);
                    }
                    else if (str == "☺l☺")
                    {
                        string[] strData = new string[] { richTextBox.Text.Split('\n')[1], SplitMainText() };
                        AddToComboBox(lComboBox, dataComboBox.ComboBoxL, strData);
                    }
                    else if (str == "☺p☺")
                    {
                        string[] strData = new string[] { richTextBox.Text.Split('\n')[1], SplitMainText() };
                        AddToComboBox(pComboBox, dataComboBox.ComboBoxP, strData);
                    }
                    else if (str == "☺t☺")
                    {
                        string[] strData = new string[] { richTextBox.Text.Split('\n')[1], SplitMainText() };
                        AddToComboBox(tComboBox, dataComboBox.ComboBoxT, strData);
                    }
                    else if (str == "☺c☺")
                    {
                        string[] strData = new string[] { richTextBox.Text.Split('\n')[1], SplitMainText() };
                        AddToComboBox(cComboBox, dataComboBox.ComboBoxC, strData);
                    }
                }
                else if (element.Name == "КнопкаТекст")
                {
                    buttonText_Click(sender, e);
                }
                else if (element.Name == "H1" || element.Name == "H2" || element.Name == "L" || element.Name == "P" || element.Name == "T" || element.Name == "C")
                {
                    buttonSpecial_Click(sender, e);
                }
            }
        }

        string SplitMainText()
        {
            string[] str = richTextBox.Text.Split('\n');
            string mainText = string.Empty;
            mainText = str[3];
            if (str.Length > 4)
            {
                for (int i = 4; str.Length > i; i++)
                {

                    mainText += "\n" + str[i];
                }
            }
            return mainText;
        }

        bool ValidAddInput()
        {
            string str = richTextBox.Text.Split('\n')[0];
            if (richTextBox.Text.Split('\n').Length >= 4 && richTextBox.Text.Split('\n')[2] == "☺Содержимое☺")
            {
                if (str == "☺h1☺" || str == "☺h2☺")
                {
                    return true;
                }
                else if (str == "☺l☺")
                {
                    // ???
                }
                else if (str == "☺p☺")
                {
                    if (System.IO.File.Exists(SplitMainText()))
                    {
                        return true;
                    }
                }
                else if (str == "☺t☺")
                {
                    // ???
                }
                else if (str == "☺c☺")
                {
                    if(System.IO.File.Exists(SplitMainText()))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        void AddToComboBox(ComboBox comboBox, System.Collections.Generic.List<string[]> saveComboBox, string[] strData)
        {
            string str = richTextBox.Text.Split('\n')[1];
            if (!comboBox.Items.Contains(str))
            {
                saveComboBox.Add(strData);
                comboBox.Items.Add(str);
                comboBox.SelectedIndex = comboBox.Items.IndexOf(str);
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
                richTextBox.Focus();
                if (comboBox == h1ComboBox)
                {
                    richTextBox.Text = "☺h1☺\n" + dataComboBox.ComboBoxH1[comboBox.SelectedIndex][0] + "\n☺Содержимое☺\n" + dataComboBox.ComboBoxH1[comboBox.SelectedIndex][1];
                    richTextBox.SelectionStart = 5 + dataComboBox.ComboBoxH1[comboBox.SelectedIndex][0].Length;
                }
                else if (comboBox == h2ComboBox)
                {
                    richTextBox.Text = "☺h2☺\n" + dataComboBox.ComboBoxH2[comboBox.SelectedIndex][0] + "\n☺Содержимое☺\n" + dataComboBox.ComboBoxH2[comboBox.SelectedIndex][1];
                    richTextBox.SelectionStart = 5 + dataComboBox.ComboBoxH2[comboBox.SelectedIndex][0].Length;
                }
                else if (comboBox == lComboBox)
                {
                    richTextBox.Text = "☺l☺\n" + dataComboBox.ComboBoxL[comboBox.SelectedIndex][0] + "\n☺Содержимое☺\n" + dataComboBox.ComboBoxL[comboBox.SelectedIndex][1];
                    richTextBox.SelectionStart = 4 + dataComboBox.ComboBoxL[comboBox.SelectedIndex][0].Length;
                }
                else if (comboBox == pComboBox)
                {
                    richTextBox.Text = "☺p☺\n" + dataComboBox.ComboBoxP[comboBox.SelectedIndex][0] + "\n☺Содержимое☺\n" + dataComboBox.ComboBoxP[comboBox.SelectedIndex][1];
                    richTextBox.SelectionStart = 4 + dataComboBox.ComboBoxP[comboBox.SelectedIndex][0].Length;
                }
                else if (comboBox == tComboBox)
                {
                    richTextBox.Text = "☺t☺\n" + dataComboBox.ComboBoxP[comboBox.SelectedIndex][0] + "\n☺Содержимое☺\n" + dataComboBox.ComboBoxT[comboBox.SelectedIndex][1];
                    richTextBox.SelectionStart = 4 + dataComboBox.ComboBoxT[comboBox.SelectedIndex][0].Length;
                }
                else if (comboBox == cComboBox)
                {
                    richTextBox.Text = "☺c☺\n" + dataComboBox.ComboBoxC[comboBox.SelectedIndex][0] + "\n☺Содержимое☺\n" + dataComboBox.ComboBoxC[comboBox.SelectedIndex][1];
                    richTextBox.SelectionStart = 4 + dataComboBox.ComboBoxC[comboBox.SelectedIndex][0].Length;
                }
            }
            else
            {
                elementLabel.Text = "нечто";
                richTextBox.Text = string.Empty;
            }
        }
        
        /*void ComboBox_IndexChanged(System.Collections.Generic.List<string[]> str, string s, ComboBox comboBox)
        {
            richTextBox.Text = "☺"+s+"☺\n" + str[comboBox.SelectedIndex][0] + "\n☺Содержимое☺\n" + str[comboBox.SelectedIndex][1];
            richTextBox.SelectionStart = 3 + s.Length + str[comboBox.SelectedIndex][0].Length;
        }*/

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
                    for (int i = elementPanel.ColumnCount - 1; i < 2 * elementPanel.ColumnCount - 1 - 2; i++)
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
                    if (comboBox == h1ComboBox)
                    {
                        dataComboBox.ComboBoxH1.RemoveAt(comboBox.SelectedIndex);
                    }
                    else if (comboBox == h2ComboBox)
                    {
                        dataComboBox.ComboBoxH2.RemoveAt(comboBox.SelectedIndex);
                    }
                    else if (comboBox == lComboBox)
                    {
                        dataComboBox.ComboBoxL.RemoveAt(comboBox.SelectedIndex);
                    }
                    else if (comboBox == pComboBox)
                    {
                        dataComboBox.ComboBoxP.RemoveAt(comboBox.SelectedIndex);
                    }
                    else if (comboBox == tComboBox)
                    {
                        dataComboBox.ComboBoxT.RemoveAt(comboBox.SelectedIndex);
                    }
                    else if (comboBox == cComboBox)
                    {
                        dataComboBox.ComboBoxC.RemoveAt(comboBox.SelectedIndex);
                    }
                    comboBox.Items.RemoveAt(comboBox.SelectedIndex);
                    ComboBox_SelectedIndexChanged(sender, e);
                }
            }
        }

        async void ReadScroll_Click(object sender, EventArgs e)
        {
            MakeReport report = new MakeReport();
            if (TextMenuItem.Checked)
            {
                dataComboBox.Text = richTextBox.Text;
            }
            else
            {
                dataComboBox.Text = text;
            }
            List<string> titleData = new List<string>();
            AddTitleData(ref titleData);
            try
            {
                await Task.Run(() => report.CreateReport(dataComboBox, NumberingMenuItem.Checked, ContentMenuItem.Checked, TitleOffOnMenuItem.Checked, int.Parse(FromNumberingTextBoxMenuItem.Text), NumberHeadingMenuItem.Checked,  this.Text.ToString(), titleData.ToArray()));
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

        void AddTitleData(ref List<string> titleData)
        {
            List<Control> controls = new List<Control>();
            foreach (Control control in titlepagePanel.Controls)
            {
                if(control.GetType().ToString() != "System.Windows.Forms.Label")
                {
                    controls.Add(control);
                }
            }
            for(int i=1;i<controls.Count;i++)
            {
                for(int f=0; f < controls.Count-i;f++)
                {
                    if(controls[f].TabIndex > controls[f+1].TabIndex)
                    {
                        Control kek = controls[f];
                        controls[f] = controls[f + 1];
                        controls[f + 1] = kek;
                    }
                }
            }
            foreach (Control control in controls)
            {
                titleData.Add(control.Text);
            }
        }

        void CloseWindow_Click(object sender, EventArgs e)
        {
            CloseWindow.Checked = !CloseWindow.Checked;
        }

        void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        void SaveTitlePagePanelCells()
        {
            rows = new int[0]; columns = new int[0];
            for (int i = 0; i < titlepagePanel.Controls.Count; i++)
            {
                rows = ArrayPushBack<int>(rows, titlepagePanel.GetCellPosition(titlepagePanel.Controls[i]).Row);
                columns = ArrayPushBack<int>(columns, titlepagePanel.GetCellPosition(titlepagePanel.Controls[i]).Column);
            }
        }

        Control[] DEFAULTtitlepagePanelControls;
        int[] rows;
        int[] columns;
        void ShowTitleElems(string cells)
        {
            titlepagePanel.SuspendLayout();
            titlepagePanel.Controls.Clear();
            PushbackControls(DEFAULTtitlepagePanelControls, titlepagePanel, 0, DEFAULTtitlepagePanelControls.Length - 1, this.rows, this.columns);

            ShowAllChildControls(titlepagePanel);
            ResetAllChildColumnSpans(titlepagePanel);
            string[] column_row = cells.Split(' ');
            int[] columns = new int[column_row.Length];
            int[] rows = new int[column_row.Length];
            for (int i = 0; i < column_row.Length; i++)
            {
                columns[i] = int.Parse(column_row[i].Split('.')[0]);
                rows[i] = int.Parse(column_row[i].Split('.')[1]);
            }
            Control[] titleSave = CopyControls(titlepagePanel, rows, columns);
            titlepagePanel.Controls.Clear();
            for (int i = 0; i < titleSave.Length; i++)
            {
                if (columns[i] >= 4 && RowElemCounter(rows, rows[i]) <= 4)
                {
                    columns[i] -= 2;
                }
                if (columns[i] >= 2 && RowElemCounter(rows, rows[i]) <= 2)
                {
                    columns[i] -= 2;
                }
            }
            PushbackControls(titleSave, titlepagePanel, 0, titleSave.Length - 1, rows, columns);
            for (int i = 0; i < titlepagePanel.Controls.Count; i++)
            {
                if (columns[i] == 3 && RowElemCounter(rows, rows[i]) <= 4)
                {
                    titlepagePanel.SetColumnSpan(titlepagePanel.Controls[i], 3);
                }
                else if (columns[i] == 1 && RowElemCounter(rows, rows[i]) <= 2)
                {
                    titlepagePanel.SetColumnSpan(titlepagePanel.Controls[i], 5);
                }
            }
            titlepagePanel.ResumeLayout();
        }

        void ShowAllChildControls(Control control)
        {
            foreach (Control ctrl in control.Controls)
            {
                ctrl.Visible = true;
            }
        }

        void ResetAllChildColumnSpans(TableLayoutPanel table)
        {
            foreach (Control ctrl in table.Controls)
            {
                table.SetColumnSpan(ctrl, 1);
            }
        }

        void ResetAllChildRowSpans(TableLayoutPanel table)
        {
            foreach (Control ctrl in table.Controls)
            {
                table.SetRowSpan(ctrl, 1);
            }
        }

        int RowElemCounter(int[] rows, int row)
        {
            int counter = 0;
            for (int i = 0; i < rows.Length; i++)
            {
                if (rows[i] == row)
                {
                    counter++;
                }
            }
            return counter;
        }

        Control[] CopyControls(TableLayoutPanel tableLayoutPanel, int startElem, int endElem)
        {
            Control[] newArray = new Control[tableLayoutPanel.Controls.Count];
            for (int i = startElem; i <= endElem; i++)
            {
                newArray[i] = tableLayoutPanel.Controls[i];
            }
            return newArray;
        }

        public static T[] ArrayPushBack<T>(T[] array, T element)
        {
            T[] newArray = new T[array.Length + 1];
            for (int i = 0; i < array.Length; i++)
            {
                newArray[i] = array[i];
            }
            newArray[newArray.Length - 1] = element;
            return newArray;
        }

        Control[] CopyControls(TableLayoutPanel tableLayoutPanel, int[] rows, int[] columns)
        {
            Control[] newArray = new Control[0];
            for (int i = 0; i < tableLayoutPanel.Controls.Count; i++)
            {
                if (rows.Length == columns.Length)
                {
                    int cellIndex = CheckControlPosition(tableLayoutPanel, i, rows, columns);
                    if (cellIndex != -1)
                    {
                        newArray = ArrayPushBack(newArray, tableLayoutPanel.Controls[i]);

                        int tmpColumn = columns[newArray.Length - 1];
                        int tmpRow = rows[newArray.Length - 1];
                        columns[newArray.Length - 1] = columns[cellIndex];
                        rows[newArray.Length - 1] = rows[cellIndex];
                        columns[cellIndex] = tmpColumn;
                        rows[cellIndex] = tmpRow;
                    }
                }
                else
                {
                    break;
                }
            }
            return newArray;
        }

        int CheckControlPosition(TableLayoutPanel tableLayoutPanel, int controlIndex, int[] rows, int[] columns)
        {
            if (rows.Length == columns.Length)
            {
                for (int i = 0; i < rows.Length; i++)
                {
                    TableLayoutPanelCellPosition ctrlToCheckPosition = tableLayoutPanel.GetCellPosition(tableLayoutPanel.Controls[controlIndex]);
                    TableLayoutPanelCellPosition ctrlInCellPosition = tableLayoutPanel.GetCellPosition(tableLayoutPanel.GetControlFromPosition(columns[i], rows[i]));
                    if (ctrlToCheckPosition == ctrlInCellPosition)
                    {
                        return i;
                    }
                }
            }
            return -1;
        }

        bool PushbackControls(Control[] controls, TableLayoutPanel tableLayoutPanel, int startElem, int endElem, int[] rows, int[] columns)
        {
            if (columns.Length != rows.Length || columns.Length < endElem - startElem + 1)
            {
                return false;
            }
            for (int i = startElem; i <= endElem; i++)
            {
                tableLayoutPanel.Controls.Add(controls[i], columns[i], rows[i]);
            }
            return true;
        }

        void work_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem toolStripMenuItem = (ToolStripMenuItem)sender;
            if (toolStripMenuItem.Checked)
            {
                return;
            }
            if (toolStripMenuItem.Text == "Обычный документ")
            {
                TextHeader("документа");
                TitleOffOnMenuItem.Visible = !DefaultDocumentMenuItem.Checked;
                ShowingTitelPanel();
            }
            else
            {
                TitleOffOnMenuItem.Visible = true;
                TitleOffOnMenuItem.Checked = false;
                ShowingTitelPanel();
                if (toolStripMenuItem.Text == "Лабораторная работа")
                {
                    TextHeader("лабораторной работы");
                    ShowTitleElems("0.0 1.0 2.1 3.1 0.3 1.3 0.4 1.4 0.6 1.6 0.7 1.7");
                }
                else if (toolStripMenuItem.Text == "Практическая работа")
                {
                    TextHeader("практической работы");
                    ShowTitleElems("0.0 1.0 2.1 3.1 0.3 1.3 0.4 1.4 0.6 1.6 0.7 1.7");
                }
                else if (toolStripMenuItem.Text == "Курсовая работа")
                {
                    TextHeader("курсовой работы");
                    ShowTitleElems("0.0 1.0 0.1 1.1 4.1 5.1 0.3 1.3 0.4 1.4 0.6 1.6 0.7 1.7");
                }
                else if (toolStripMenuItem.Text == "Реферат")
                {
                    TextHeader("реферата");
                }
                else if (toolStripMenuItem.Text == "Дипломная работа")
                {
                    TextHeader("дипломной работы");
                }
                else if (toolStripMenuItem.Text == "ВКР")
                {
                    TextHeader("ВКР");
                }
                else if (toolStripMenuItem.Text == "РГР")
                {
                    TextHeader("РГР");
                }
            }
            DefaultDocumentMenuItem.Checked = false;
            LabMenuItem.Checked = false;
            PracticeMenuItem.Checked = false;
            KursMenuItem.Checked = false;
            RefMenuItem.Checked = false;
            DiplomMenuItem.Checked = false;
            VKRMenuItem.Checked = false;
            RGRMenuItem.Checked = false;
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

        void HideSpecials()
        {
            tableTypeInserts.Visible = false;
        }

        void ShowSpecials()
        {
            tableTypeInserts.Visible = true;
        }

        void HideAddButton()
        {
            PanelWithButton.Controls.Find("Добавить", true)[0].Visible = false;
            PanelWithButton.ColumnStyles[1].Width = 0;
        }

        void ShowAddButton()
        {
            PanelWithButton.ColumnStyles[1].Width = 151;
            PanelWithButton.Controls.Find("Добавить", true)[0].Visible = true;
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
                HideAddButton();
                DownPanelMI = SubstitutionMenuItem;
            }
            else if (MenuItem == TextMenuItem)
            {
                wndSize.Text.Current = this.Size;
                this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                richTextBox.EnableAutoDragDrop = false;
                CustomSizeGrip.Visible = false;
                HideSpecials();
                DownPanelMI = TextMenuItem;
                richTextBox.Margin = new Padding(3, 3, 3, 3);
            }
            MenuItem.Checked = false;
        }

        void UpdateSize(ToolStripMenuItem MenuItem)
        {
            if (MenuItem == TitlePageMenuItem)
            {
                this.MinimumSize = wndSize.Title.Min;
                this.MaximumSize = wndSize.Title.Max;
            }
            else if (MenuItem == SubstitutionMenuItem)
            {
                this.MinimumSize = wndSize.Subst.Min;
                this.MaximumSize = wndSize.Subst.Max;
            }
            else if (MenuItem == TextMenuItem)
            {
                this.MinimumSize = wndSize.Text.Min;
                this.MaximumSize = wndSize.Text.Max;
                this.Size = wndSize.Text.Current;
            }
        }

        void ShowElements(ToolStripMenuItem MenuItem)
        {
            UpdateSize(MenuItem);
            if (MenuItem == TitlePageMenuItem)
            {
                buttonDown.Visible = true;
                titlepagePanel.Visible = true;
                DownPanel.Visible = false;
                buttonUp.Visible = false;
                facultyLabel.Focus();
            }
            else if (MenuItem == SubstitutionMenuItem)
            {
                if(TitleOffOnMenuItem.Checked)
                {
                    buttonUp.Visible = true;
                }
                DownPanel.Visible = true;
                pictureBox.Visible = true;
                elementPanel.Visible = true;
                elementLabel.Text = "нечто";
                buttonText.Text = "К тексту";
                textPicturePanel.ColumnStyles[0].Width = 60;
                textPicturePanel.ColumnStyles[1].Width = 40;
                DownPanelMI = SubstitutionMenuItem;
                ShowAddButton();
                richTextBox.Text = string.Empty;
                richTextBox.Focus();
            }
            else if (MenuItem == TextMenuItem)
            {
                richTextBox.EnableAutoDragDrop = true;
                if (TitleOffOnMenuItem.Checked)
                {
                    buttonUp.Visible = true;
                }
                DownPanel.Visible = true;
                CustomSizeGrip.Visible = true;
                buttonText.Text = "К подстановкам";
                elementLabel.Text = "текст";
                textPicturePanel.ColumnStyles[1].Width = 0;
                textPicturePanel.ColumnStyles[0].Width = 100;
                DownPanelMI = TextMenuItem;
                this.AutoSizeMode = AutoSizeMode.GrowOnly;
                richTextBox.Text = text;
                richTextBox.SelectionStart = richTextBox.Text.Length;
                UpdateTypeButton();
                MatchWordPage();
                richTextBox.Focus();
            }
            if(MenuItem!=null)
            {
                MenuItem.Checked = true;
            }
        }
        
        // Когда возвращаешься из текста в меню заголовков, текстбокс пустой, а заголовок остается выбранным. Можно либо не выбирать заголовок, либо заполнять текстбокс.
        void MatchWordPage()
        {
            int left = 3 + (richTextBox.Width - 790) / 2 + 76;
            int right = 3 + (richTextBox.Width - 790) / 2 + 56 - 16; // 16 is scrollbar width
            richTextBox.Margin = new Padding(left,richTextBox.Margin.Top, right, richTextBox.Margin.Bottom);
        }

        void refreshMenu()
        {
            elementPanel.SuspendLayout();
            for (int i = 0; i < elementPanel.ColumnStyles.Count - 1; i++)
            {
                elementPanel.ColumnStyles[i].SizeType = SizeType.Percent;
                elementPanel.ColumnStyles[i].Width = 0;
            }

            elementPanel.ColumnStyles[0].Width = 4;
            elementPanel.ColumnStyles[elementPanel.ColumnStyles.Count - 1].Width = 4;

            for (int i = menuLeftIndex; i < menuLeftIndex + 4; i++)
            {
                elementPanel.ColumnStyles[i].Width = 23;
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
            if (menuLeftIndex == elementPanel.ColumnStyles.Count - 5)
            {
                buttonForward.Visible = false;
            }
            else
            {
                buttonForward.Visible = true;
            }
        }

        string TypeRichBox()
        {
            string str = string.Empty;
            foreach (char ch in richTextBox.Text)
            {
                if (ch == '\n')
                {
                    break;
                }
                str += ch;
            }
            return str;
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

        void replaceMenuRow()
        {
            Control[] menuPanelSave = new Control[PanelWithButton.Controls.Count];
            for (int i = 0; i < PanelWithButton.Controls.Count; i++)
            {
                menuPanelSave[i] = PanelWithButton.Controls[i];
            }
            PanelWithButton.Controls.Clear();
            for (int i = 0; i < menuPanelSave.Length; i++)
            {
                if (menuPanelSave[i].Name == "buttonText")
                {
                    PanelWithButton.Controls.Add(GetMenuTextBtnReplacement(), 2, 0);
                }
                else if (menuPanelSave[i].Name == "ButtonAdd")
                {
                    PanelWithButton.Controls.Add(GetMenuButtonReplacement(1, "Добавить")[0], 1, 0);
                }
                else
                {
                    PanelWithButton.Controls.Add(menuPanelSave[i]);
                }
            }
        }

        void replaceMenu()
        {
            GlobalFont.SetFont(heading1Label.Font, heading1Label.Font.Style);
            PictureBox[] menuPBarray = GetMenuLabelReplacement(elementPanel.ColumnCount - 2);
            replaceMenuRow();
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

        void UnselectComboBoxes()
        {
            for (int i = elementPanel.ColumnCount - 1; i < elementPanel.Controls.Count - 1; i++)
            {
                ComboBox cmbBox = (ComboBox)elementPanel.Controls[i];
                cmbBox.SelectedIndex = -1;
            }
        }

        void createTemplate(object sender)
        {
            UnselectComboBoxes();
            Control control = (Control)sender;
            if (control.Name == "Заголовок 1")
            {
                DefaultTypeRichBox("h1");
            }
            else if(control.Name == "Заголовок 2")
            {
                DefaultTypeRichBox("h2");
            }
            else if (control.Name == "Список")
            {
                DefaultTypeRichBox("l");
            }
            else if (control.Name == "Картинка")
            {
                DefaultTypeRichBox("p");
                fileNames = null;
                pictureBox.Refresh();
            }
            else if (control.Name == "Таблица")
            {
                DefaultTypeRichBox("t");
            }
            else if (control.Name == "Код")
            {
                DefaultTypeRichBox("c");
            }
        }

        void DefaultTypeRichBox(string type)
        {
            string beginning = "☺" + type + "☺";
            richTextBox.Text = beginning + "\n\n☺Содержимое☺\n";
            richTextBox.SelectionStart = beginning.Length + 1;
        }

        void facultyComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str = string.Empty;
            professorComboBox.Items.Clear();
            if (facultyComboBox.SelectedIndex == 0)
            {
                str = "Амелина О.В.!Артёмов А.В.!Валухов В.А.!Волков В.Н.!Гордиенко А.П.!Демидов А.В.!Загородних Н.А.!Захарова О.В.!Конюхова О.В.!Корнаева Е.П.!Королева А.К.!Короткий А.В.!Коськин А.В.!Кравцова Э.А.!Лукьянов П.В.!Лунёв Р.А.!Лыськов О.Э.!Машкова А.Л.!Митин А.А.!Новиков С.В.!Новикова Е.В.!Олькина Е.В.!Преснецова В.Ю.!Раков В.И.!Рыженков Д.В.!Савина О.А.!Санников Д.П.!Сезонов Д.С.!Соков О.А.!Стычук А.А.!Терентьев С.В.!Ужаринский А.Ю.!Фроленкова Л.Ю.!Фролов А.И.!Фролова В.А.!Чижов А.В.!Шатеев Р.В.";
            }
            else if (facultyComboBox.SelectedIndex == 1)
            {
                str = "Бондарева Л.А.!Дрёмин В.В.!Дунаев А.В.!Жидков А.В.!Козлов И.О.!Козлова Л.Д.!Марков В.В.!Матюхин С.И.!Незнанов А.И.!Подмастерьев К.В.!Секаева Ж.А.!Селихов А.В.!Углова Н.В.!Шуплецов В.В.!Яковенко М.В.";
            }
            else if (facultyComboBox.SelectedIndex == 2)
            {
                str = "Белевская Ю.А.!Ерёменко В.Т.!Мишин Д.С.!Пеньков Н.Г!Савва Ю.Б.!Фисун А.П.";
            }
            else if (facultyComboBox.SelectedIndex == 3)
            {
                str = "Батуров Д.П.!Гордон В.А.!Кирсанова О.В.!Матюхин С.И.!Потураева Т.В.!Ромашин С.Н.!Семёнова Г.А.!Фроленкова Л.Ю.!Якушина С.И.";
            }
            else if (facultyComboBox.SelectedIndex == 4)
            {
                str = "Аксёнов К.В.!Багров В.В.!Батенков А.А.!Варгашкин В.Я.!Власова М.А.!Воронина О.А.!Донцов В.М.!Косчинский С.Л.!Лобанова В.А.!Лобода О.А.!Майоров М.В.!Мишин В.В.!Муравьёв А.А.!Плащенков Д.А.!Рязанцев П.Н.!Селихов А.В.!Суздальцев А.И.!Тугарев А.С.!Тютякин А.В.!Филина А.В.";
            }
            else if (facultyComboBox.SelectedIndex == 5)
            {
                str = "Аксёнов К.В.!Гладышев А.В.!Качанов А.Н.!Коренков Д.А.!Королева Т.Г.!Петров Г.Н.!Филина А.В.!Чернышов В.А.";
            }
            foreach (string s in str.Split('!'))
            {
                professorComboBox.Items.Add(s);
            }
        }

        int GetLineOfCursor(RichTextBox richTextBox)
        {
            int line = 0;
            for (int i=0;i<richTextBox.SelectionStart;i++)
            {
                if(richTextBox.Text[i]=='\n')
                {
                    line += 1;
                }
            }
            return line;
        }

        void richTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            int line = GetLineOfCursor(richTextBox);
            if (DownPanelMI == SubstitutionMenuItem)
            {
                if (ComboBoxSelected())
                {
                    bool last = richTextBox.Text.Split('\n')[1].Length - 1 + richTextBox.Text.Split('\n')[0].Length == richTextBox.SelectionStart - 2;
                    bool start2 = richTextBox.Text.Split('\n')[0].Length == richTextBox.SelectionStart - 1;
                    bool start4 = richTextBox.Text.Split('\n')[0].Length + richTextBox.Text.Split('\n')[1].Length + richTextBox.Text.Split('\n')[2].Length == richTextBox.SelectionStart - 3;
                    if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.X ||
                        e.KeyCode == Keys.Enter && (line >= 0 && line <= 2) ||
                        (e.KeyCode == Keys.Back && (line == 0 || line == 2 || (start2 && richTextBox.SelectionLength == 0 && line==1)|| (line==1&& richTextBox.SelectedText.Contains("\n")) || (start4 && richTextBox.SelectionLength == 0 && line==3)))||
                        (e.KeyCode == Keys.Delete && (line == 0 || line == 2 || last || (richTextBox.SelectedText.Contains("\n") && (line<3))))
                       )
                    {
                        e.Handled = true;
                    }
                    else if (e.KeyCode == Keys.Down && line == 1)
                    {
                        richTextBox.SelectionStart += richTextBox.Text.Split('\n')[2].Length + richTextBox.Text.Split('\n')[1].Length +2;
                    }
                    else if (e.KeyCode == Keys.Up && line == 3)
                    {
                        richTextBox.SelectionStart -= richTextBox.Text.Split('\n')[2].Length + 2;
                    }
                    if (richTextBox.Text == richTextBox.SelectedText && (e.KeyCode == Keys.Back || e.KeyCode == Keys.Delete))
                    {
                        string[] lines = richTextBox.Text.Split('\n');
                        lines[1] = "";
                        lines[3] = "";
                        richTextBox.Text = lines[0] + "\n" + lines[1] + "\n" + lines[2] + "\n" + lines[3];
                        richTextBox.SelectionStart = lines[0].Length + 1;
                    }
                }
            }
           
            // ✔
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.V)
            {
                if (!(DownPanelMI == SubstitutionMenuItem && (line == 0 || line == 2)))
                {
                    richTextBox.SuspendLayout();
                    int insPt = richTextBox.SelectionStart;
                    string postContent = richTextBox.Text.Substring(insPt);
                    string preContent = richTextBox.Text.Substring(0, insPt);
                    try
                    {
                        byte[] bytes = Encoding.Unicode.GetBytes(Clipboard.GetData(DataFormats.UnicodeText).ToString());
                        string ctrl_v = Encoding.Unicode.GetString(bytes);
                        if (line == 1)
                        {
                            ctrl_v = ctrl_v.Replace("\n", "");
                            ctrl_v = ctrl_v.Replace("\r", "");
                        }
                        richTextBox.Text = preContent + ctrl_v + postContent;
                    }
                    catch (NullReferenceException)
                    {
                        System.Console.Write(e);
                    }
                    richTextBox.SelectionStart = richTextBox.TextLength - postContent.Length;
                    richTextBox.ResumeLayout();
                    e.Handled = true;
                }
                else
                {
                    e.Handled = true;
                }
            }
        }

        bool ComboBoxSelected()
        {
            if (h1ComboBox.SelectedIndex != -1 || h2ComboBox.SelectedIndex != -1 || lComboBox.SelectedIndex != -1 || pComboBox.SelectedIndex != -1 || tComboBox.SelectedIndex != -1 || cComboBox.SelectedIndex != -1)
                return true;
            return false;
        }

        private void richTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (DownPanelMI == SubstitutionMenuItem)
            {
                if (ComboBoxSelected())
                {
                    int line = richTextBox.GetLineFromCharIndex(richTextBox.SelectionStart);
                    if ((Control.ModifierKeys == Keys.Control && e.KeyChar == (char)Keys.V && line < 3) ||
                        (richTextBox.SelectionLength > 0 && (e.KeyChar != (char)Keys.Back || e.KeyChar != (char)Keys.Delete)))
                    {
                        e.Handled = true;
                    }
                }
            }
        }

        private void CustomInterface_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                UnselectComboBoxes();
            }
            else if (Control.ModifierKeys == Keys.Control)
            {
                if(e.KeyCode == Keys.S)
                {
                    if (Control.ModifierKeys != Keys.Shift)
                    {
                        Save_Click(sender, e);
                    }
                    else
                    {
                        SaveAsMenuItem_Click(sender, e);
                    }
                }
                else if (e.KeyCode == Keys.O)
                {
                    Open_Click(sender, e);
                }
                else if (e.KeyCode == Keys.N)
                {
                    CreateMenuItem_Click(sender, e);
                }
                else if(e.KeyCode == Keys.Q)
                {
                    Application.Exit();
                }
            }
        }

        private void textPicturePanel_Paint(object sender, PaintEventArgs e)
        {
            if (!richTextBox.Visible)
            {
                Point locationOnForm = textPicturePanel.PointToClient(richTextBox.PointToScreen(richTextBox.Location));
                e.Graphics.FillRectangle(new SolidBrush(richTextBox.BackColor), locationOnForm.X - richTextBox.Margin.Left, locationOnForm.Y - richTextBox.Margin.Left, richTextBox.Width, richTextBox.Height);
            }
        }

        private void richTextBox_VisibleChanged(object sender, EventArgs e)
        {
            textPicturePanel.Invalidate();
        }

        private void ContentMenuItem_Click(object sender, EventArgs e)
        {
            ContentMenuItem.Checked = !ContentMenuItem.Checked;
        }

        private void NumberingMenuItem_Click(object sender, EventArgs e)
        {
            NumberingMenuItem.Checked = !NumberingMenuItem.Checked;
            FromNumberingTextBoxMenuItem.Visible = NumberingMenuItem.Checked;
            Document.ShowDropDown();
            NumberingMenuItem.Select();
            FromNumberingTextBoxMenuItem.Visible = true;
        }

        private void TitleOffOnMenuItem_Click(object sender, EventArgs e)
        {
            ShowingTitelPanel();
        }

        void ShowingTitelPanel()
        {
            if (TitleOffOnMenuItem.Checked && TitlePageMenuItem.Checked)
            {
                HiddenElements(TitlePageMenuItem);
                ShowElements(DownPanelMI);
            }
            TitlePageMenuItem.Visible = !TitleOffOnMenuItem.Checked;
            buttonUp.Visible = !TitleOffOnMenuItem.Checked;
            TitleOffOnMenuItem.Checked = !TitleOffOnMenuItem.Checked;
        }

        private void CreateMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Нужно ли сохранить?", "Нужно ли сохранить?", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            if(result == DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(fileNames))
                {
                    SaveWordKiller(fileNames);
                }
                else
                {
                    SaveAsMenuItem_Click(sender, e);
                }
            }
            fileNames = null;
            ClearGlobal();
            text = "";
            textDragOnDrop = "";
            menuLeftIndex = 1;
            dataComboBox = new DataComboBox();
            richTextBox.Text = "";
            if(DownPanelMI==TextMenuItem)
            {
                HideSpecials();
            }
            foreach (Control control in titlepagePanel.Controls)
            {
                if (control.GetType().ToString() == "System.Windows.Forms.TextBox")
                {
                    control.Text = "";
                }
            }
        }

        private void NumberHeadingMenuItem_Click(object sender, EventArgs e)
        {
            NumberHeadingMenuItem.Checked = !NumberHeadingMenuItem.Checked;
        }

        private void SetAsDefaultMenuItem_Click(object sender, EventArgs e)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo("computerdefaults");
            startInfo.UseShellExecute = true;
            Process.Start(startInfo);
        }

        
        // основателями поведенческой школы в психологии являются: павлов, уотсон, скиннер/спиннер
    }
}

