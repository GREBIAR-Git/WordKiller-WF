using System.Drawing;

namespace MakeReportWord
{
    class WindowSize
    {
        MinMax title;
        MinMax subst;
        MinMax text;
        public MinMax Title { get { return title; } set { title = value; } }
        public MinMax Subst { get { return subst; } set { subst = value; } }
        public MinMax Text { get { return text; } set { text = value; } }
        public WindowSize()
        {
            title = new MinMax(new Size(846, 393));
            subst = new MinMax(new Size(846, 577));
            text = new MinMax(new Size(846, 577), new Size(846, 2000));
        }
    }

    class MinMax
    {
        Size min;
        Size max;
        Size current;
        public Size Min { get { return min; } set { min = value; } }
        public Size Max { get { return max; } set { max = value; } }
        public Size Current { get { return current; } set { current = value; } }
        public MinMax(Size min, Size max)
        {
            this.min = min;
            this.max = max;
            current = min;
        }
        public MinMax(Size minmax)
        {
            this.min = minmax;
            this.max = minmax;
            current = min;
        }
    }
}
