using System;
using System.Drawing;
using System.Windows.Forms;

namespace MakeReportWord
{
    public partial class CustomInterface
    {
        void CustomInterface_Shown(object sender, EventArgs e)
        {
            this.BackColor = Color.FromArgb(255, 50, 39, 62);
            facultyLabel.BackColor = Color.FromArgb(255, 253, 219, 124);
            Students.BackColor = Color.FromArgb(255, 253, 219, 124);
            Shifr.BackColor = Color.FromArgb(255, 253, 219, 124);
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
            elementComboBox.BackColor = Color.FromArgb(255, 50, 39, 62);
            buttonText.BackColor = Color.FromArgb(255, 238, 230, 246);
            h1ComboBox.BackColor = Color.FromArgb(255, 238, 230, 246);
            pComboBox.BackColor = Color.FromArgb(255, 238, 230, 246);
            h2ComboBox.BackColor = Color.FromArgb(255, 238, 230, 246);
            lComboBox.BackColor = Color.FromArgb(255, 238, 230, 246);
            titlepagePanel.BackColor = Color.FromArgb(255, 50, 39, 62);
            MainPanel.BackColor = Color.FromArgb(255, 50, 39, 62);
            DownPanel.BackColor = Color.FromArgb(255, 50, 39, 62);
            displayedLabel.ForeColor = Color.FromArgb(255, 238, 230, 246);
            elementLabel.ForeColor = Color.FromArgb(255, 238, 230, 246);
            elementComboBox.ForeColor = Color.FromArgb(255, 238, 230, 246);
        }

        //buttonDownStart
        void buttonDown_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonDown.BackgroundImage = Properties.Resources.arrowsDownPressed;
            }
        }
        void buttonDown_MouseEnter(object sender, EventArgs e)
        {
            buttonDown.BackgroundImage = Properties.Resources.arrowsDownSelected;
        }
        void buttonDown_MouseLeave(object sender, EventArgs e)
        {
            buttonDown.BackgroundImage = Properties.Resources.arrowsDown;
        }
        //buttonDownEnd

        //buttonUpStart
        void buttonUp_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonUp.BackgroundImage = Properties.Resources.arrowsUpPressed;
            }
        }
        void buttonUp_MouseEnter(object sender, EventArgs e)
        {
            buttonUp.BackgroundImage = Properties.Resources.arrowsUpSelected;
        }
        void buttonUp_MouseLeave(object sender, EventArgs e)
        {
            buttonUp.BackgroundImage = Properties.Resources.arrowsUp;
        }
        //buttonUpEnd

        //buttonForwardStart
        void buttonForward_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonForward.BackgroundImage = Properties.Resources.arrowsRightPressed;
            }
        }
        void buttonForward_MouseEnter(object sender, EventArgs e)
        {
            buttonForward.BackgroundImage = Properties.Resources.arrowsRightSelected;
        }
        void buttonForward_MouseLeave(object sender, EventArgs e)
        {
            buttonForward.BackgroundImage = Properties.Resources.arrowsRight;
        }
        //buttonForwardEnd

        //buttonBackStart
        void buttonBack_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                buttonBack.BackgroundImage = Properties.Resources.arrowsLeftPressed;
            }
        }
        void buttonBack_MouseEnter(object sender, EventArgs e)
        {
            buttonBack.BackgroundImage = Properties.Resources.arrowsLeftSelected;
        }
        void buttonBack_MouseLeave(object sender, EventArgs e)
        {
            buttonBack.BackgroundImage = Properties.Resources.arrowsLeft;
        }
        //buttonBackEnd

        //menuExortStart
        void menuExort_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.exortPressed;
            }
        }
        void menuExort_MouseEnter(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.exortSelected;
        }
        void menuExort_MouseLeave(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.exort;
        }
        //menuExortEnd

        //menuWexStart
        void menuWex_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.wexPressed;
            }
        }
        void menuWex_MouseEnter(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.wexSelected;
        }
        void menuWex_MouseLeave(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.wex;
        }
        //menuWexEnd

        //menuQuasStart
        void menuQuas_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.quasPressed;
            }
        }
        void menuQuas_MouseEnter(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.quasSelected;
        }
        void menuQuas_MouseLeave(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.quas;
        }
        //menuQuasEnd

        //menuButtonPBStart
        void menuButtonPB_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                PictureBox element = (PictureBox)sender;
                element.BackgroundImage = Properties.Resources.BtnPressed;
            }
        }
        void menuButtonPB_MouseEnter(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.BtnSelected;
        }
        void menuButtonPB_MouseLeave(object sender, EventArgs e)
        {
            PictureBox element = (PictureBox)sender;
            element.BackgroundImage = Properties.Resources.Btn;
        }
        //menuButtonPBEnd
    }
}
