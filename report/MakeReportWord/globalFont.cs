using System.Drawing;

namespace MakeReportWord
{
    class GlobalFont
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
