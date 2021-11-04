using System.Collections.Generic;
using System.Windows.Forms;

namespace MakeReportWord
{
    class ElementComboBox
    {
        public string Name { get; set; }
        public List<string[]> Data { get; set; }
        public ComboBox Form { get; set; }

        public ElementComboBox(ComboBox form)
        {
            Form = form;
            Data = new List<string[]>();
        }
    }
}
