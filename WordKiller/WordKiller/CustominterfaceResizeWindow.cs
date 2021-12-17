using System.Drawing;
using System.Windows.Forms;

namespace WordKiller
{
    public partial class CustomInterface
    {
        Point mouse = new Point();
        bool resizing = false;
        private void CustomSizeGrip_MouseDown(object sender, MouseEventArgs e)
        {
            resizing = true;
            if (e.Button == MouseButtons.Left)
            {
                mouse.X = e.X; mouse.Y = e.Y;
            }
        }

        private void CustomSizeGrip_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && resizing)
            {
                Size sizeDiff = new Size(0, e.Y - mouse.Y);
                this.Size += sizeDiff;
            }
            resizing = false;
        }

        private void CustomSizeGrip_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && resizing)
            {
                Size sizeDiff = new Size(0, e.Y - mouse.Y);
                this.Size += sizeDiff;
            }
        }
    }
}
