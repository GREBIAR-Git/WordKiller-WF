using System.Collections.Generic;
using System.Windows.Forms;

namespace MakeReportWord
{
    class DataComboBox
    {
        Dictionary<string, ElementComboBox> comboBoxes { get;set; }
        public string text { get; set; } 
        
        public ElementComboBox SearchComboBox(ComboBox comboBoxForm)
        {
            foreach(KeyValuePair<string,ElementComboBox> comboBox in this.comboBoxes)
            {
                if (comboBox.Value.Form == comboBoxForm)
                {
                    return comboBox.Value;
                }
            }
            return null;
        }

        public int Sum()
        {
            int sum=0;
            foreach (KeyValuePair<string, ElementComboBox> comboBox in this.comboBoxes)
            {
                sum += comboBox.Value.Data.Count;
            }
            return sum;
        }
        public DataComboBox(ComboBox h1, ComboBox h2, ComboBox l, ComboBox p, ComboBox t, ComboBox c)
        {
            comboBoxes = new Dictionary<string, ElementComboBox>();
            comboBoxes["h1"] = new ElementComboBox(h1);
            comboBoxes["h2"] = new ElementComboBox(h1);
            comboBoxes["l"] = new ElementComboBox(h1);
            comboBoxes["p"] = new ElementComboBox(h1);
            comboBoxes["t"] = new ElementComboBox(h1);
            comboBoxes["c"] = new ElementComboBox(c);
        }
    }
}
