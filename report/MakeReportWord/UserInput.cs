namespace MakeReportWord
{
    class UserInput
    {
        string[] comboBoxH1, comboBoxH2, comboBoxL, comboBoxP, comboBoxT, comboBoxC;
        string text;
        public string Text
        {
            get{return text;}
            set{text = value;}
        }
        public string[] ComboBoxH1
        {
            get { return comboBoxH1; }
            set { comboBoxH1 = value; }
        }
        public string[] ComboBoxH2
        {
            get { return comboBoxH2; }
            set { comboBoxH2 = value; }
        }
        public string[] ComboBoxL
        {
            get { return comboBoxL; }
            set { comboBoxL = value; }
        }
        public string[] ComboBoxP
        {
            get { return comboBoxP; }
            set { comboBoxP = value; }
        }
        public string[] ComboBoxT
        {
            get { return comboBoxT; }
            set { comboBoxT = value; }
        }
        public string[] ComboBoxC
        {
            get { return comboBoxC; }
            set { comboBoxC = value; }
        }
    }
}
