﻿using System;
using Microsoft.Office.Interop.Word;

namespace MakeReportWord
{
    internal class MakeReport
    {
        Document doc;
        Range word;

        public void CreateReportLab(string faculty, string numberLab, string theme, string discipline, string professor, string year)
        {
            var LengthDoc = 0;
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
            word.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            text = "ОТЧЁТ";
            WriteTextWord(ref LengthDoc, text);
            word.Font.Size = 16;
            word.Font.Bold = 1;

            text = "По лабораторной работе №" + numberLab;
            WriteTextWord(ref LengthDoc, text);
            word.Paragraphs.SpaceAfter = 10;
            word.Font.Bold = 0;

            text = "на тему: «" + theme + "»" + SkipLine(1) + "по дисциплине: «" + discipline + "»" + SkipLine(8); ;
            WriteTextWord(ref LengthDoc, text);
            word.Font.Size = 14;
            word.Paragraphs.SpaceAfter = 0;

            text = "Выполнили: Музалевский Н.С., Аллянов М.Д." + SkipLine(1) +
                "Институт приборостроения, автоматизации и информационных технологий" + SkipLine(1) +
                "Направление: 09.03.04 «Программная инженерия»" + SkipLine(1) +
                "Группа: 92ПГ";
            WriteTextWord(ref LengthDoc, text);
            word.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

            text = "Проверил: " + professor;
            WriteTextWord(ref LengthDoc, text);
            word.Paragraphs.SpaceAfter = 10;

            text = SkipLine(1) + "Отметка о зачёте: ";
            WriteTextWord(ref LengthDoc, text);
            word.Paragraphs.SpaceAfter = 0;

            text = "Дата: «____» __________ " + year + "г.";
            WriteTextWord(ref LengthDoc, text);
            word.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

            text = SkipLine(8) + "Орел, " + year;
            WriteTextWord(ref LengthDoc, text);
            word.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            PageBreak(ref LengthDoc);
        }

        void WriteTextWord(ref int Length, string text)
        {
            Length += word.Text.Length;
            word.Text += text;
            word = doc.Range(Length, Type.Missing);
        }

        void WriteTextWord(string text)
        {
            word = doc.Range();
            word.Text = text;
        }

        void PageBreak(ref int Length)
        {
            Length += word.Text.Length - 1;
            word = doc.Range(Length, Type.Missing);
            word.InsertBreak(0);
            word.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        }

        void DefaultText()
        {
            word.Font.Size = 14;
            word.Font.Bold = 0;
            word.Font.ColorIndex = 0;
            word.Paragraphs.Space15();
            word.Paragraphs.FirstLineIndent = CentimetersToPoints(1.5f);
            // word.Paragraphs.LeftIndent = CentimetersToPoints(1.5f);
            word.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        }

        void Heading1()
        {
            word.Font.Size = 14;
            word.Font.Bold = 1;
            word.Font.AllCaps = 1;
            word.Font.ColorIndex = 0;
            word.Paragraphs.Space15();
            word.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }

        void Heading2()
        {
            word.Font.Size = 14;
            word.Font.Bold = 1;
            word.Font.ColorIndex = 0;
            word.Paragraphs.Space15();
            word.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }

        void Picture()
        {
            word.Paragraphs.Space15();
            word.Font.ColorIndex = 0;
            word.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }

        void СaptionForPicture()
        {
            word.Font.Size = 14;
            word.Font.Bold = 0;
            word.Font.ColorIndex = 0;
            word.Paragraphs.SpaceAfter= 8;
            word.Paragraphs.Space15();
            word.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
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
