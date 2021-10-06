using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MakeReportWord
{
    public partial class CustomInterface
    {
        string saveFileName = string.Empty;
        void Open_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "|*.wordkiller;";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                OpenWordKiller(openFileDialog.FileName);
            }
        }

        void OpenWordKiller(string fileName)
        {
            saveFileName = fileName;
            ClearGlobal();
            FileStream file = new FileStream(fileName, FileMode.Open);
            StreamReader reader = new StreamReader(file);
            //try
            //{
            string data = reader.ReadToEnd();
            for (int i = 1; i < data.Length; i++)
            {
                if (data[i - 1] == '\r')
                {
                    data = data.Remove(i, 1);
                }
            }
            string[] lines = data.Split('\r');


            bool readingText = false;
            List<Control> controls = new List<Control>();
            foreach (string line in lines)
            {
                if (line.StartsWith("☺Menu☺"))
                {
                    string[] menuItem = line.Remove(0, 6).Split('!');
                    work_Click(TypeMenuItem.DropDown.Items.Find(menuItem[0], false)[0], new EventArgs());
                    if (menuItem[0] != "DefaultDocumentMenuItem")
                    {
                        foreach (Control control in titlepagePanel.Controls)
                        {
                            if (control.GetType().ToString() != "System.Windows.Forms.Label")
                            {
                                controls.Add(control);
                            }
                        }
                    }
                    NumberHeadingMenuItem.Checked = bool.Parse(menuItem[1]);
                }

                if (line.StartsWith("☺TextStart☺"))
                {
                    readingText = true;
                }
                else if (readingText)
                {
                    if (line.StartsWith("☺TextEnd☺"))
                    {
                        readingText = false;
                    }
                    else
                    {
                        text += line + "\n";
                    }
                }
                else
                {
                    string[] variable_value = line.Split('☺');
                    if (variable_value.Length == 2)
                    {
                        for (int i = 0; i < controls.Count; i++)
                        {
                            if (LoadingOfTwo(variable_value, controls[i]))
                            {
                                controls.RemoveAt(i);
                                break;
                            }
                        }
                    }
                    else if (variable_value.Length == 3)
                    {
                        if (variable_value[0].StartsWith("h1ComboBox"))
                        {
                            h1ComboBox.Items.Add(variable_value[1]);
                            string[] str = new string[] { variable_value[1], variable_value[2] };
                            dataComboBox.ComboBoxH1.Add(str);
                        }
                        else if (variable_value[0].StartsWith("h2ComboBox"))
                        {
                            h2ComboBox.Items.Add(variable_value[1]);
                            string[] str = new string[] { variable_value[1], variable_value[2] };
                            dataComboBox.ComboBoxH2.Add(str);
                        }
                        else if (variable_value[0].StartsWith("lComboBox"))
                        {
                            lComboBox.Items.Add(variable_value[1]);
                            string[] str = new string[] { variable_value[1], variable_value[2] };
                            dataComboBox.ComboBoxL.Add(str);
                        }
                        else if (variable_value[0].StartsWith("pComboBox"))
                        {
                            pComboBox.Items.Add(variable_value[1]);
                            string[] str = new string[] { variable_value[1], variable_value[2] };
                            dataComboBox.ComboBoxP.Add(str);
                        }
                        else if (variable_value[0].StartsWith("tComboBox"))
                        {
                            tComboBox.Items.Add(variable_value[1]);
                            string[] str = new string[] { variable_value[1], variable_value[2] };
                            dataComboBox.ComboBoxT.Add(str);
                        }
                        else if (variable_value[0].StartsWith("cComboBox"))
                        {
                            cComboBox.Items.Add(variable_value[1]);
                            string[] str = new string[] { variable_value[1], variable_value[2] };
                            dataComboBox.ComboBoxC.Add(str);
                        }
                    }
                }
            }
            //}
            // catch
            // {
            //   MessageBox.Show("Файл повреждён");
            // }
            reader.Close();
        }

        bool LoadingOfTwo(string[] variable_value, Control control)
        {
            if (variable_value[0].StartsWith(control.Name))
            {
                control.Text = variable_value[1];
                return true;
            }
            return false;
        }

        void Save_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(saveFileName))
            {
                SaveWordKiller(saveFileName);
            }
            else
            {
                SaveAsMenuItem_Click(sender, e);
            }
        }

        void SaveAsMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "|*.wordkiller;";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                saveFileName = saveFileDialog.FileName;
                SaveWordKiller(saveFileName);
            }
        }

        void SaveWordKiller(string nameFile)
        {
            FileStream fileStream = System.IO.File.Open(nameFile, FileMode.Create);
            StreamWriter output = new StreamWriter(fileStream);
            foreach (ToolStripMenuItem item in TypeMenuItem.DropDown.Items)
            {
                if (item.Checked)
                {
                    output.Write(MegaConvert("☺Menu☺" + item.Name.ToString() + "!" + NumberHeadingMenuItem.Checked.ToString()));
                }
            }
            output.WriteLine(MegaConvert("facultyComboBox☺" + facultyComboBox.Text));
            output.WriteLine(MegaConvert("numberTextBox☺" + numberTextBox.Text));
            output.WriteLine(MegaConvert("themeTextBox☺" + themeTextBox.Text));
            output.WriteLine(MegaConvert("disciplineTextBox☺" + disciplineTextBox.Text));
            output.WriteLine(MegaConvert("professorComboBox☺" + professorComboBox.Text));
            output.WriteLine(MegaConvert("yearTextBox☺" + yearTextBox.Text));
            output.WriteLine(MegaConvert("shifrTextBox☺" + shifrTextBox.Text));
            output.WriteLine(MegaConvert("studentsTextBox☺" + studentsTextBox.Text));
            SaveCombobox(output, h1ComboBox, dataComboBox.ComboBoxH1);
            SaveCombobox(output, h2ComboBox, dataComboBox.ComboBoxH2);
            SaveCombobox(output, lComboBox, dataComboBox.ComboBoxL);
            SaveCombobox(output, pComboBox, dataComboBox.ComboBoxP);
            SaveCombobox(output, tComboBox, dataComboBox.ComboBoxT);
            SaveCombobox(output, cComboBox, dataComboBox.ComboBoxC);
            output.WriteLine(MegaConvert("☺TextStart☺"));
            output.WriteLine(MegaConvert(text));
            output.WriteLine(MegaConvert("☺TextEnd☺"));

            output.Close();
        }

        string MegaConvert(string str)
        {
            string megaStr = str;

            megaStr = StringToBinaryString(megaStr);
            megaStr = RepeatEncodingBinary(megaStr);
            // to abc

            return megaStr;
        }

        string StringToBinaryString(string str)
        {
            string hexStr = BitConverter.ToString(Encoding.Unicode.GetBytes(str));
            hexStr = hexStr.Replace("-", "");
            string binarystring = String.Join(String.Empty, hexStr.Select(c => Convert.ToString(Convert.ToInt32(c.ToString(), 16), 2).PadLeft(4, '0')));
            binarystring = binarystring.Replace("0", "a");


            return binarystring;
        }

        string RepeatEncodingBinary(string binarystring)
        {
            string encoded = "";


            int digit = 0;
            while (digit < binarystring.Length-1)
            {
                int counter = 0;
                for (counter = 0; counter < 9 && digit+counter<binarystring.Length && binarystring[digit + counter] == binarystring[digit]; counter++)
                {
                }
                encoded += (counter + 1).ToString() + binarystring[digit].ToString();
                digit += counter + 1;
                // /???
            }



            return encoded;
        }

        void SaveCombobox(StreamWriter output, ComboBox comboBox, List<string[]> Lstr)
        {
            for (int i = 0; i < comboBox.Items.Count; i++)
            {
                output.WriteLine(MegaConvert(comboBox.Name + "☺" + comboBox.Items[i].ToString() + "☺" + Lstr[i][1])   );
            }
        }

        void ClearGlobal()
        {
            dataComboBox = new DataComboBox();
            for (int i = elementPanel.ColumnCount - 1; i < elementPanel.Controls.Count - 1; i++)
            {
                ComboBox cmbBox = (ComboBox)elementPanel.Controls[i];
                cmbBox.Items.Clear();
            }
        }
    }
}
