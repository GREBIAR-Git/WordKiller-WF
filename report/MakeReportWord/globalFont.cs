using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MakeReportWord
{
    class globalFont
    {
        static Font font;
        static FontStyle fontStyle;
        public static void SetFont(Font newFont, FontStyle newFontStyle)
        {
            font = new Font(newFont, newFontStyle);
            fontStyle = newFontStyle;
        }
        public static Font GetFont()
        {
            return new Font(font, fontStyle);
        }
    }
}
