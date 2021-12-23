using System;
using System.Windows.Forms;

namespace WordKiller
{
    public partial class CustomInterface
    {
        void ShowElements(ToolStripMenuItem MenuItem)
        {
            if (MenuItem != null)
            {
                MenuItem.Checked = true;
            }
            UpdateSize(MenuItem);
            if (MenuItem == TitlePageMenuItem)
            {
                MainPanel.RowStyles[1].Height = 30;
                buttonDown.Visible = true;
                titlepagePanel.Visible = true;
                DownPanel.Visible = false;
                buttonUp.Visible = false;
                MainPanel.RowStyles[2].Height = 0;
                MainPanel.RowStyles[MainPanel.RowCount - 1].Height = 5;
                facultyLabel.Focus();
            }
            else if (MenuItem == SubstitutionMenuItem)
            {
                DownPanel.Visible = true;
                pictureBox.Visible = true;
                elementPanel.Visible = true;
                elementLabel.Text = "нечто";
                buttonText.Text = "К тексту";
                SwitchToPrimaryRTB();
                DownPanelMI = SubstitutionMenuItem;
                ShowAddButton();
                UnselectComboBoxes();
                richTextBox.Text = string.Empty;
                richTextBox.Focus();
                elementComboBox.Visible = false;
                elementLabel.Visible = true;
            }
            else if (MenuItem == TextMenuItem)
            {
                richTextBox.EnableAutoDragDrop = true;
                DownPanel.Visible = true;
                CustomSizeGrip.Visible = true;
                buttonText.Text = "К подстановкам";
                elementLabel.Text = "текст";
                DownPanelMI = TextMenuItem;
                SwitchRTB();
                this.AutoSizeMode = AutoSizeMode.GrowOnly;
                richTextBox.Text = data.Text;
                richTextBox.SelectionStart = richTextBox.Text.Length;
                ShowSpecials();
                UpdateTypeButton();
                //MatchWordPage();
                richTextBox.Focus();
                elementComboBox.Visible = true;
                elementLabel.Visible = false;
                MainPanel.RowStyles[MainPanel.RowCount - 2].Height = 25;
                CursorLocationPanel.Visible = true;
            }
        }
        void SwitchRTB()
        {
            string item = elementComboBox.SelectedItem.ToString();
            if (item == "Весь текст")
            {
                SwitchToPrimaryRTB();
            }
            else
            {
                item = item.Substring(item.IndexOf(" ") + 1);
                SwitchToSecondaryRTB();
                richTextBoxSecondary.Text = GetSelectedSection(item);
            }
        }

        void SwitchToPrimaryRTB()
        {
            richTextBoxSecondary.Visible = false;
            textPicturePanel.ColumnStyles[1].SizeType = SizeType.Absolute;
            textPicturePanel.ColumnStyles[1].Width = 0;
            textPicturePanel.ColumnStyles[0].SizeType = SizeType.Percent;
            richTextBox.Visible = true;
            textPicturePanel.ColumnStyles[2].SizeType = SizeType.Percent;
            if (SubstitutionMenuItem.Checked)
            {
                textPicturePanel.ColumnStyles[0].Width = 60;
                textPicturePanel.ColumnStyles[2].Width = 40;
            }
            else if (TextMenuItem.Checked)
            {
                textPicturePanel.ColumnStyles[0].Width = 100;
                textPicturePanel.ColumnStyles[2].Width = 0;
            }
        }

        //нужнл переделать
        void MatchWordPage()
        {
            int left = 3 + (richTextBox.Width - 790) / 2 + 76;
            int right = 3 + (richTextBox.Width - 790) / 2 + 56 - 16; // 16 is scrollbar width
            richTextBox.Margin = new Padding(left, richTextBox.Margin.Top, right, richTextBox.Margin.Bottom);
        }
        void HideElements(ToolStripMenuItem MenuItem)
        {
            if (MenuItem == TitlePageMenuItem)
            {
                buttonDown.Visible = false;
                titlepagePanel.Visible = false;
                buttonDown.Visible = false;
                MainPanel.RowStyles[1].Height = 0;
                if (typeDocument != TypeDocument.DefaultDocument)
                {
                    MainPanel.RowStyles[2].Height = 45;
                    buttonUp.Visible = true;
                }
                MainPanel.RowStyles[MainPanel.RowCount - 1].Height = 21;
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
                MainPanel.RowStyles[MainPanel.RowCount - 2].Height = 0;
                CursorLocationPanel.Visible = false;
                wndSize.Text.Current = this.Size;
                this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                richTextBox.EnableAutoDragDrop = false;
                CustomSizeGrip.Visible = false;
                HideSpecials();
                DownPanelMI = TextMenuItem;
            }
            MenuItem.Checked = false;
        }

        void View_MenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ClickMenuItem = (ToolStripMenuItem)sender;
            if (TitlePageMenuItem.Checked)
            {
                HideElements(TitlePageMenuItem);
                ShowElements(ClickMenuItem);
            }
            else if (SubstitutionMenuItem.Checked)
            {
                HideElements(SubstitutionMenuItem);
                ShowElements(ClickMenuItem);
            }
            else if (TextMenuItem.Checked)
            {
                HideElements(TextMenuItem);
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

        void UpdateSize(ToolStripMenuItem MenuItem)
        {
            if (MenuItem == TitlePageMenuItem)
            {
                this.MinimumSize = wndSize.Title.Min;
                this.MaximumSize = wndSize.Title.Max;
                MainPanel.RowStyles[0].SizeType = SizeType.Percent;
                MainPanel.RowStyles[3].Height = 0;
                MainPanel.RowStyles[0].Height = 100;
            }
            else if (MenuItem == SubstitutionMenuItem)
            {
                this.MinimumSize = wndSize.Subst.Min;
                this.MaximumSize = wndSize.Subst.Max;
                MainPanel.RowStyles[0].Height = 0;
                MainPanel.RowStyles[3].Height = 100;
            }
            else if (MenuItem == TextMenuItem)
            {
                this.MinimumSize = wndSize.Text.Min;
                this.MaximumSize = wndSize.Text.Max;
                this.Size = wndSize.Text.Current;
                MainPanel.RowStyles[0].Height = 0;
                MainPanel.RowStyles[3].Height = 100;
            }
        }
    }
}
