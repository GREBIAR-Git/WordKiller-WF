using System;
using Word = Microsoft.Office.Interop.Word;

namespace MakeReportWord
{
    internal class MakeReport
    {
        public void CreateReport(string faculty, string numberLab, string theme, string discipline, string professor, string year)
        {
            var end = Type.Missing;
            var app = new Word.Application();
            app.Visible = true;
            var doc = app.Documents.Add();
            var word = doc.Range();
            word.PageSetup.TopMargin = CentimetersToPoints(2);
            word.PageSetup.BottomMargin = CentimetersToPoints(2);
            word.PageSetup.LeftMargin = CentimetersToPoints(3);
            word.PageSetup.RightMargin = CentimetersToPoints(1.5f);
            word.Font.Size = 14;
            word.Font.Name = "Times New Roman";
            word.Paragraphs.SpaceAfter = 0;
            word.Paragraphs.Space1();
            word.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            word.Text = "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ";
            word.Text += "РОССИЙСКОЙ ФЕДЕРАЦИИ";
            word.Text += "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ";
            word.Text += "ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ";
            word.Text += "«ОРЛОВСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ";
            word.Text += "ИМЕНИ И.С.ТУРГЕНЕВА»" + SkipLine(1);
            word.Text += "Кафедра " + faculty + SkipLine(3);
            var LengthDoc = word.Text.Length;
            word.Text += "ОТЧЁТ";
            word = doc.Range(LengthDoc, end);
            word.Font.Size = 16;
            word.Font.Bold = 1;
            LengthDoc += word.Text.Length;
            word.Text += "По лабораторной работе №" + numberLab;
            word = doc.Range(LengthDoc, end);
            word.Paragraphs.SpaceAfter = 10;
            word.Font.Bold = 0;
            LengthDoc += word.Text.Length;
            word.Text += "на тему: «" + theme + "»";
            word = doc.Range(LengthDoc, end);
            word.Font.Size = 14;
            word.Paragraphs.SpaceAfter = 0;
            word.Text += "по дисциплине: «" + discipline + "»" + SkipLine(8);
            LengthDoc += word.Text.Length;
            word.Text += "Выполнили: Музалевский Н.С., Аллянов М.Д.";
            word.Text += "Институт приборостроения, автоматизации и информационных технологий";
            word.Text += "Направление: 09.03.04 «Программная инженерия»";
            word.Text += "Группа: 92ПГ";
            word = doc.Range(LengthDoc, end);
            word.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            LengthDoc += word.Text.Length;
            word.Text += "Проверил: " + professor;
            word = doc.Range(LengthDoc, end);
            word.Paragraphs.SpaceAfter = 10;
            LengthDoc += word.Text.Length;
            word.Text += SkipLine(1) + "Отметка о зачете: ";
            word = doc.Range(LengthDoc, end);
            word.Paragraphs.SpaceAfter = 0;
            LengthDoc += word.Text.Length;
            word.Text += "Дата: «____» __________ " + year + "г.";
            word = doc.Range(LengthDoc, end);
            word.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            LengthDoc += word.Text.Length;
            word.Text += SkipLine(8) + "Орел, " + year;
            word = doc.Range(LengthDoc, end);
            word.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            LengthDoc += word.Text.Length - 1;
            word = doc.Range(LengthDoc, end);
            word.InsertBreak(0);
            word.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
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
