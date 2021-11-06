using System.Drawing;
using System.Windows.Forms;

namespace MakeReportWord
{
    public partial class CustomInterface
    {
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

        void buttonText_Paint(object sender, PaintEventArgs e)
        {
            PictureBox pb = (PictureBox)sender;
            int fontSize = 10;
            if (MouseIsOverControl(pb) && Control.MouseButtons != MouseButtons.Left)
            {
                fontSize = 12;
            }
            using (Font fnt = new Font(GlobalFont.GetFont().Name, fontSize))
            {
                string str;
                if (elementLabel.Text == "текст")
                {
                    str = "К подстановкам";
                }
                else
                {
                    str = "К тексту";
                }
                SizeF stringSize = e.Graphics.MeasureString(str, fnt);
                e.Graphics.DrawString(str, fnt, Brushes.Black, new Point((int)(pb.Width / 2 - stringSize.Width / 2), (int)(pb.Height / 2 - stringSize.Height / 2)));
            }
        }

        void menuPB_Paint(object sender, PaintEventArgs e)
        {
            PictureBox pb = (PictureBox)sender;
            Control controlPB = (Control)sender;
            int index = elementPanel.Controls.IndexOf(controlPB);
            Font fnt = GlobalFont.GetFont();

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

        void menuPBbtn_Paint(object sender, PaintEventArgs e)
        {
            PictureBox pb = (PictureBox)sender;
            string str = pb.Name;
            Font fnt = GlobalFont.GetFont();
            int size = 10; int selected = 2;
            if (MouseIsOverControl(pb) && Control.MouseButtons != MouseButtons.Left)
            {
                fnt = new Font(fnt.Name, size + selected);
            }
            else
            {
                fnt = new Font(fnt.Name, size);
            }
            using (fnt)
            {
                SizeF stringSize = e.Graphics.MeasureString(str, fnt);
                e.Graphics.DrawString(str, fnt, Brushes.Black, new Point((int)(pb.Width / 2 - stringSize.Width / 2), (int)(pb.Height / 2 - stringSize.Height / 2)));
            }
        }

        void pictureBox_Paint(object sender, PaintEventArgs e)
        {
            string str = TypeRichBox();
            if (dragging == 0)
            {
                if (str == "☺h1☺")
                {
                    e.Graphics.DrawImage(Properties.Resources.Red, 0, 0, pictureBox.Width, pictureBox.Height);
                    int index = dataComboBox.ComboBox["h1"].Form.SelectedIndex;
                    if (index!=-1)
                    {
                        if(NumberHeadingMenuItem.Checked)
                        {
                            DrawText((index+1).ToString() + " " + dataComboBox.ComboBox["h1"].Data[index][1].ToUpper(), e);
                        }
                        else
                        {
                            DrawText(dataComboBox.ComboBox["h1"].Data[index][1].ToUpper(), e);
                        }
                    }
                    else
                    {
                        DrawText("Заголовок".ToUpper(), e);
                    }
                }
                else if (str == "☺h2☺")
                {
                    e.Graphics.DrawImage(Properties.Resources.Red, 0, 0, pictureBox.Width, pictureBox.Height);
                    int index = dataComboBox.ComboBox["h2"].Form.SelectedIndex;
                    if (index != -1)
                    {
                        if (NumberHeadingMenuItem.Checked)
                        {
                            DrawText("H1." + (index + 1).ToString() + " " + dataComboBox.ComboBox["h2"].Data[index][1], e);
                        }
                        else
                        {
                            DrawText(dataComboBox.ComboBox["h2"].Data[index][1], e);
                        }
                    }
                    else
                    {
                        DrawText("Заголовок", e);
                    }
                }
                else if (str == "☺l☺")
                {
                    e.Graphics.DrawImage(Properties.Resources.Red, 0, 0, pictureBox.Width, pictureBox.Height);
                    DrawText("Список", e);
                }
                else if (str == "☺t☺")
                {
                    e.Graphics.DrawImage(Properties.Resources.Red, 0, 0, pictureBox.Width, pictureBox.Height);
                    DrawText("Таблица", e);
                }
                else if (str == "☺p☺")
                {
                    if (fileNames == null)
                    {
                        e.Graphics.DrawImage(Properties.Resources.Picture, 0, 0, pictureBox.Width, pictureBox.Height);
                        SizeF stringSize = e.Graphics.MeasureString("Не указан", new Font("Microsoft Sans Serif", 14));
                        e.Graphics.DrawString("Не указан", new Font("Microsoft Sans Serif", 14), new SolidBrush(Color.Black), new Point((int)(pictureBox.Width / 2 - stringSize.Width / 2), pictureBox.Height / 2 + 50));
                    }
                    else
                    {
                        try
                        {
                            e.Graphics.DrawImage(Properties.Resources.purpleBackground, 0, 0, pictureBox.Width, pictureBox.Height);
                            e.Graphics.DrawImage(Image.FromFile(fileNames), 0, 0, pictureBox.Width, pictureBox.Height);
                        }
                        catch (System.OutOfMemoryException)
                        {
                            richTextBox.Clear();
                            e.Graphics.DrawImage(Properties.Resources.DragNDrop, 0, 0, pictureBox.Width, pictureBox.Height);
                        }
                    }
                }
                else if (str == "☺c☺")
                {
                    e.Graphics.DrawImage(Properties.Resources.Code, 0, 0, pictureBox.Width, pictureBox.Height);
                    if (fileNames == null)
                    {
                        SizeF stringSize = e.Graphics.MeasureString("Не указан", new Font("Microsoft Sans Serif", 14));
                        e.Graphics.DrawString("Не указан", new Font("Microsoft Sans Serif", 14), new SolidBrush(Color.Black), new Point((int)(pictureBox.Width / 2 - stringSize.Width / 2), pictureBox.Height / 2 + 50));
                    }
                    else
                    {
                        string nameFile = fileNames.Split('\\')[fileNames.Split('\\').Length - 1];
                        SizeF stringSize = e.Graphics.MeasureString(nameFile, new Font("Microsoft Sans Serif", 14));
                        e.Graphics.DrawString(nameFile, new Font("Microsoft Sans Serif", 14), new SolidBrush(Color.Black), new Point((int)(pictureBox.Width / 2 - stringSize.Width / 2), pictureBox.Height / 2 + 50));
                    }
                }
            }
            else if (dragging == 1)
            {
                e.Graphics.DrawImage(Properties.Resources.pictureCode_Picture, 0, 0, pictureBox.Width, pictureBox.Height);
                DragNDrop_PaintText(e, Color.White, Color.Black);
            }
            else if (dragging == 2)
            {
                e.Graphics.DrawImage(Properties.Resources.pictureCode_Code, 0, 0, pictureBox.Width, pictureBox.Height);
                DragNDrop_PaintText(e, Color.Black, Color.White);
            }
            else if (dragging == 3)
            {
                e.Graphics.DrawImage(Properties.Resources.pictureCode, 0, 0, pictureBox.Width, pictureBox.Height);
                DragNDrop_PaintText(e, Color.Black, Color.Black);
            }
        }
        void DrawText(string text, PaintEventArgs e)
        {
            SizeF stringSize = e.Graphics.MeasureString(text, new Font("Microsoft Sans Serif", 20));
            e.Graphics.DrawString(text, new Font("Microsoft Sans Serif", 20), new SolidBrush(Color.Black), new Point((int)(pictureBox.Width / 2 - stringSize.Width / 2), pictureBox.Height / 2 - 20));
        }

        void DragNDrop_PaintText(PaintEventArgs e, Color picture, Color code)
        {
            string str1 = "Для";
            SizeF stringSize = e.Graphics.MeasureString(str1, new Font("Microsoft Sans Serif", 14));
            e.Graphics.DrawString(str1, new Font("Microsoft Sans Serif", 14), new SolidBrush(picture), new Point((int)(pictureBox.Width / 4 - stringSize.Width / 2), pictureBox.Height / 2 + 30));
            str1 = "картинки";
            stringSize = e.Graphics.MeasureString(str1, new Font("Microsoft Sans Serif", 14));
            e.Graphics.DrawString(str1, new Font("Microsoft Sans Serif", 14), new SolidBrush(picture), new Point((int)(pictureBox.Width / 4 - stringSize.Width / 2), pictureBox.Height / 2 + 50));

            str1 = "Для";
            stringSize = e.Graphics.MeasureString(str1, new Font("Microsoft Sans Serif", 14));
            e.Graphics.DrawString(str1, new Font("Microsoft Sans Serif", 14), new SolidBrush(code), new Point((int)(3 * pictureBox.Width / 4 - stringSize.Width / 2), pictureBox.Height / 2 + 30));
            str1 = "кода";
            stringSize = e.Graphics.MeasureString(str1, new Font("Microsoft Sans Serif", 14));
            e.Graphics.DrawString(str1, new Font("Microsoft Sans Serif", 14), new SolidBrush(code), new Point((int)(3 * pictureBox.Width / 4 - stringSize.Width / 2), pictureBox.Height / 2 + 50));
        }
    }
}
