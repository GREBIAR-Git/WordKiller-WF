using System;
using System.Drawing;
using System.Windows.Forms;

namespace MakeReportWord
{
    public partial class CustomInterface
    {
        int dragging = 0; // 1 - мышь на левой половине пикчербокса, 2 - на правой, 3 - мышь на форме, но не на пикчербоксе, 0 - ничего не перетаскивается

        void DragNDropPanel_DragEnter(object sender, DragEventArgs e)
        {
            string str = TypeRichBox();
            if (str != "☺h1☺" && str != "☺h2☺" && str != "☺l☺" && str != "☺t☺")
            {
                textDragOnDrop = richTextBox.Text;
                richTextBox.Text = string.Empty;
                e.Effect = DragDropEffects.All;
                richTextBox.Visible = true;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        void DragNDropPanel_DragLeave(object sender, EventArgs e)
        {
            string str = TypeRichBox();
            if (str != "☺h1☺" && str != "☺h2☺" && str != "☺l☺" && str != "☺t☺")
            {
                richTextBox.Text = textDragOnDrop;
                Point absCoords = pictureBox.PointToScreen(pictureBox.Location);
                if (Cursor.Position.X < absCoords.X || Cursor.Position.X > absCoords.X + pictureBox.Width || Cursor.Position.Y < absCoords.Y || Cursor.Position.Y > absCoords.Y + pictureBox.Height)
                {
                    dragging = 3;
                }
                else
                {
                    dragging = 0;
                }
                richTextBox.Visible = false;
            }
            pictureBox.Refresh();
        }

        void DragNDropPanel_DragDrop(object sender, DragEventArgs e)
        {
            string str = TypeRichBox();
            if (str != "☺h1☺" && str != "☺h2☺" && str != "☺l☺" && str != "☺t☺")
            {
                var data = e.Data.GetData(DataFormats.FileDrop);
                if (data != null)
                {
                    fileNames = (data as string[])[0];
                    if (fileNames.Length > 0)
                    {
                        string nameFile = fileNames.Split('\\')[fileNames.Split('\\').Length - 1];
                        nameFile = nameFile.Substring(0, nameFile.LastIndexOf('.'));
                        Point controlRelatedCoords = this.DragNDropPanel.PointToClient(new Point(e.X, e.Y));
                        if (controlRelatedCoords.X < 148)
                        {
                            richTextBox.Text = "☺p☺\n" + nameFile + "\n☺Содержимое☺\n" + fileNames;
                        }
                        else
                        {
                            richTextBox.Text = "☺c☺\n" + nameFile + "\n☺Содержимое☺\n" + fileNames;
                        }
                    }
                }
            }
            dragging = 0;
            pictureBox.Refresh();
        }
        void DragNDropPanel_DragOver(object sender, DragEventArgs e)
        {
            string str = TypeRichBox();
            if (str != "☺h1☺" && str != "☺h2☺" && str != "☺l☺" && str != "☺t☺")
            {
                Point controlRelatedCoords = this.DragNDropPanel.PointToClient(new Point(e.X, e.Y));
                if (controlRelatedCoords.X <= 154)
                {
                    dragging = 1;
                }
                else
                {
                    dragging = 2;
                }
            }
            pictureBox.Refresh();
        }

        void CustomInterface_DragEnter(object sender, DragEventArgs e)
        {
            string str = TypeRichBox();
            if (str != "☺h1☺" && str != "☺h2☺" && str != "☺l☺" && str != "☺t☺")
            {
                fileNames = null;
                textDragOnDrop = richTextBox.Text;
                richTextBox.Text = string.Empty;
                richTextBox.Visible = false;
            }
        }

        void CustomInterface_DragOver(object sender, DragEventArgs e)
        {
            string str = TypeRichBox();
            if (str != "☺h1☺" && str != "☺h2☺" && str != "☺l☺" && str != "☺t☺")
            {
                dragging = 3;
            }
            pictureBox.Refresh();
        }

        void CustomInterface_DragLeave(object sender, EventArgs e)
        {
            Point absCoords = pictureBox.PointToScreen(pictureBox.Location);
            if (Cursor.Position.X < absCoords.X || Cursor.Position.X > absCoords.X + pictureBox.Width || Cursor.Position.Y < absCoords.Y || Cursor.Position.Y > absCoords.Y + pictureBox.Height)
            {
                string str = TypeRichBox();
                if (str != "☺h1☺" && str != "☺h2☺" && str != "☺l☺" && str != "☺t☺")
                {
                    richTextBox.Text = textDragOnDrop;
                    dragging = 0;
                    richTextBox.Visible = true;
                }
                pictureBox.Refresh();
            }
        }
    }
}
