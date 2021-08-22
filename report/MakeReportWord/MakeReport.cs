using System;
using Microsoft.Office.Interop.Word;

namespace MakeReportWord
{
    
    class MakeReport
    {
        Document doc;
        Range word;
        bool pgBreak = true;
        char special = '☺';
        public void CreateReportLab(string faculty, string numberLab, string theme, string discipline, string professor, string year, UserInput content)
        {
            Application app = new Application();
            app.Visible = true;
            doc = app.Documents.Add();
            word = null;

            string text = "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ" + SkipLine(1) +
                "РОССИЙСКОЙ ФЕДЕРАЦИИ" + SkipLine(1) +
                "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ" + SkipLine(1) +
                "ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ" + SkipLine(1) +
                "«ОРЛОВСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ" + SkipLine(1) +
                "ИМЕНИ И.С.ТУРГЕНЕВА»" + SkipLine(2) +
                "Кафедра " + faculty + SkipLine(3);
            WriteTextWord(text);
            PageMargin(2,2,3,1.5f);
            word.Font.Size = 14;
            word.Font.Name = "Times New Roman";
            word.Paragraphs.SpaceAfter = 0;
            word.Paragraphs.Space1();
            word.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            text = "ОТЧЁТ";
            WriteTextWord(text);
            word.Font.Size = 16;
            word.Font.Bold = 1;

            text = "По лабораторной работе №" + numberLab;
            WriteTextWord(text);
            word.Paragraphs.SpaceAfter = 10;
            word.Font.Bold = 0;

            text = "на тему: «" + theme + "»" + SkipLine(1) + "по дисциплине: «" + discipline + "»" + SkipLine(8); ;
            WriteTextWord(text);
            word.Font.Size = 14;
            word.Paragraphs.SpaceAfter = 0;

            text = "Выполнили: Музалевский Н.С., Аллянов М.Д." + SkipLine(1) +
                "Институт приборостроения, автоматизации и информационных технологий" + SkipLine(1) +
                "Направление: 09.03.04 «Программная инженерия»" + SkipLine(1) +
                "Группа: 92ПГ";
            WriteTextWord(text);
            word.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

            text = "Проверил: " + professor;
            WriteTextWord(text);
            word.Paragraphs.SpaceAfter = 10;

            text = SkipLine(1) + "Отметка о зачёте: ";
            WriteTextWord(text);
            word.Paragraphs.SpaceAfter = 0;

            text = "Дата: «____» __________ " + year + "г.";
            WriteTextWord(text);
            word.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

            text = SkipLine(8) + "Орел, " + year;
            WriteTextWord(text);
            word.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            PageBreak();

            if(content.Text!=null)
            {
                ProcessContent(content);
            }
        }

        void ProcessContent(UserInput content)
        {
            int heading1 = 1;
            int heading2 = 1;
            int heading2all = 1;
            string def = string.Empty;
            for (int i = 0; i<content.Text.Length;i++)
            {
                if(content.Text[i]==special)
                {
                    if(def!=string.Empty)
                    {
                        DefaultText(def);
                        def = string.Empty;
                    }
                    if (content.Text[i+1] == 'h')
                    {
                        if (content.Text[i + 2] == '1')
                        {
                            i += 2;
                            string text = heading1.ToString() + " " + ProcessSpecial(heading1,"h1", content);
                            Heading1(text);
                            heading1++;
                            heading2 = 1;
                        }
                        else if (content.Text[i + 2] == '2')
                        {
                            i += 2;
                            string text = (heading1-1).ToString() + "." + heading2.ToString() + " " + ProcessSpecial(heading2all,"h2", content);
                            Heading2(text);
                            heading2all++;
                            heading2++;
                        }
                    }
                    else if (content.Text[i + 1] == 'l')
                    {
                        i += 1;
                        string text = ProcessSpecial(i,"l", content);
                        List(text);
                    }
                    else if (content.Text[i + 1] == 'p')
                    {
                        i += 1;
                        string text = ProcessSpecial(i,"p", content);
                        СaptionForPicture(text);
                    }
                    else if (content.Text[i + 1] == 't')
                    {
                        i += 1;
                        string text = ProcessSpecial(i, "t", content);
                        Table(text);
                    }
                    else if (content.Text[i + 1] == 'c')
                    {
                        i += 1;
                        string text = ProcessSpecial(i, "c", content);
                        Code(text);
                    }
                }
                else
                {
                    def += content.Text[i];
                }
            }
            if (def != string.Empty)
            {
                DefaultText(def);
                def = string.Empty;
            }
        }

        string ProcessSpecial(int i, string special, UserInput content)
        {
            string text = string.Empty;
            if (special == "h1")
            {
                text = content.ComboBoxH1[i - 1];
            }
            else if (special == "h2")
            {
                text = content.ComboBoxH2[i - 1];
            }
            else if (special == "l")
            {
                text = content.ComboBoxL[i - 1];
            }
            else if (special == "p")
            {
                text = content.ComboBoxP[i - 1];
            }
            else if (special == "t")
            {
                text = content.ComboBoxT[i - 1];
            }
            else if (special == "c")
            {
                text = content.ComboBoxT[i - 1];
            }
            return text;
        }

        void WriteTextWord(string text)
        {
            Range wordTemp = doc.Range();
            int Length = wordTemp.Text.Length;
            if (word == null)
            {
                word = doc.Range();
                word.Text = text;
            }
            else
            {
                word.Text += text;
            }
            if(pgBreak)
            {
                word = doc.Range(Length, Type.Missing);
            }
            else
            {
                word = doc.Range(Length-1, Type.Missing);
                pgBreak = true;
            }
        }

        void PageBreak()
        {
            if (pgBreak)
            {
                pgBreak = false;
                Range word1 = doc.Range();
                int Length = word1.Text.Length;
                word = doc.Range(Length - 1, Type.Missing);
                word.InsertBreak(0);
                word.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            }
        }

        void DefaultText(string text)
        {
            WriteTextWord(text);
            word.Font.AllCaps = 0;
            word.Font.Size = 14;
            word.Font.Bold = 0;
            word.Font.ColorIndex = 0;
            word.Paragraphs.Space15();
            word.Paragraphs.FirstLineIndent = CentimetersToPoints(1.5f);
            word.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        }

        void Heading1(string text)
        {
            PageBreak();
            WriteTextWord(text);
            word.Paragraphs.FirstLineIndent = CentimetersToPoints(1.5f);
            word.Font.Size = 14;
            word.Font.Bold = 1;
            word.Font.AllCaps = 1;
            word.Font.ColorIndex = 0;
            word.Paragraphs.Space15();
            word.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }

        void Heading2(string text)
        {
            WriteTextWord(SkipLine(1)+text);
            word.Paragraphs.FirstLineIndent = CentimetersToPoints(1.5f);
            word.Font.AllCaps = 0;
            word.Font.Size = 14;
            word.Font.Bold = 1;
            word.Font.ColorIndex = 0;
            word.Paragraphs.Space15();
            word.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }

        void List(string text)
        {

        }

        void Picture()
        {
            word.Paragraphs.Space15();
            word.Paragraphs.FirstLineIndent = 0;
            word.Font.AllCaps = 0;
            word.Font.ColorIndex = 0;
            word.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }

        void СaptionForPicture(string text)
        {
            WriteTextWord(text);
            word.Paragraphs.FirstLineIndent = 0;
            word.Font.AllCaps = 0;
            word.Font.Size = 14;
            word.Font.Bold = 0;
            word.Font.ColorIndex = 0;
            word.Paragraphs.SpaceAfter= 8;
            word.Paragraphs.Space15();
            word.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }

        void Table(string text)
        {

        }

        void Code(string text)
        {

        }

        void PageMargin(float top, float bottom, float left, float right)
        {
            word.PageSetup.TopMargin = CentimetersToPoints(top);
            word.PageSetup.BottomMargin = CentimetersToPoints(bottom);
            word.PageSetup.LeftMargin = CentimetersToPoints(left);
            word.PageSetup.RightMargin = CentimetersToPoints(right);
        }

        float CentimetersToPoints(float cen)
        {
            return cen * 28.3465f;
        }

        string SkipLine(int quantity)
        {
            var str = string.Empty;
            for (var i = 0; i < quantity; i++)
            {
                str += Environment.NewLine;
            }
            return str;
        }
    }
}
