using System.Collections.Generic;

namespace MakeReportWord
{
    class DataComboBox
    {
        List<string[]> comboBoxH1, comboBoxH2, comboBoxL, comboBoxP, comboBoxT, comboBoxC;
        string text;
        public string Text
        {
            get{return text;}
            set{text = value;}
        }
        public List<string[]> ComboBoxH1
        {
            get { return comboBoxH1; }
            set { comboBoxH1 = value; }
        }
        public List<string[]> ComboBoxH2
        {
            get { return comboBoxH2; }
            set { comboBoxH2 = value; }
        }
        public List<string[]> ComboBoxL
        {
            get { return comboBoxL; }
            set { comboBoxL = value; }
        }
        public List<string[]> ComboBoxP
        {
            get { return comboBoxP; }
            set { comboBoxP = value; }
        }
        public List<string[]> ComboBoxT
        {
            get { return comboBoxT; }
            set { comboBoxT = value; }
        }
        public List<string[]> ComboBoxC
        {
            get { return comboBoxC; }
            set { comboBoxC = value; }
        }
        public int Sum()
        {
            return comboBoxH1.Count + comboBoxH2.Count + comboBoxL.Count + comboBoxP.Count + comboBoxT.Count + comboBoxC.Count;
        }
        public DataComboBox()
        {
            comboBoxH1 = new List<string[]>();
            comboBoxH2 = new List<string[]>();
            comboBoxL = new List<string[]>();
            comboBoxP = new List<string[]>();
            comboBoxT = new List<string[]>();
            comboBoxC = new List<string[]>();
        }
    }
}
