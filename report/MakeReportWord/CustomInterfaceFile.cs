using System;
using System.Collections.Generic;
using System.IO;
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
            try
            {
                string data = reader.ReadToEnd();
                data = EncodingDecoding.MegaConvertD(data);
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
                            foreach (KeyValuePair<string, ElementComboBox> comboBox in dataComboBox.ComboBox)
                            {
                                if (variable_value[0].StartsWith(comboBox.Key + "ComboBox"))
                                {
                                    comboBox.Value.Form.Items.Add(variable_value[1]);
                                    string[] str = new string[] { variable_value[1], variable_value[2] };
                                    comboBox.Value.Data.Add(str);
                                    break;
                                }
                            }
                        }
                    }
                }
                if(text.Length>0)
                {
                    text = text.Remove(text.Length - 1);
                }
            }
            catch
            {
               MessageBox.Show("Файл повреждён");
            }
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
            string save = string.Empty;
            foreach (ToolStripMenuItem item in TypeMenuItem.DropDown.Items)
            {
                if (item.Checked)
                {
                    save += "☺Menu☺" + item.Name.ToString() + "!" + NumberHeadingMenuItem.Checked.ToString()+"\n";
                }
            }
            save += "facultyComboBox☺" + facultyComboBox.Text + "\n";
            save += "numberTextBox☺" + numberTextBox.Text + "\n";
            save += "themeTextBox☺" + themeTextBox.Text + "\n";
            save += "disciplineTextBox☺" + disciplineTextBox.Text + "\n";
            save += "professorComboBox☺" + professorComboBox.Text + "\n";
            save += "yearTextBox☺" + yearTextBox.Text + "\n";
            save += "shifrTextBox☺" + shifrTextBox.Text + "\n";
            save += "studentsTextBox☺" + studentsTextBox.Text + "\n";
            foreach (KeyValuePair<string, ElementComboBox> comboBox in dataComboBox.ComboBox)
            {
                save += SaveCombobox(output, comboBox.Value, comboBox.Key);
            }
            save += "☺TextStart☺" + "\n";
            save += text + "\n";
            save += "☺TextEnd☺" + "\n";
            output.Write(EncodingDecoding.MegaConvert(save));
            output.Close();
        }

        string SaveCombobox(StreamWriter output, ElementComboBox comboBox, string name)
        {
            string comboBoxSave = string.Empty;
            for (int i = 0; i < comboBox.Form.Items.Count; i++)
            {
                comboBoxSave+=name + "ComboBox" + "☺" + comboBox.Form.Items[i].ToString() + "☺" + comboBox.Data[i][1] + "\n";
            }
            return comboBoxSave;
        }

        void ClearGlobal()
        {
            dataComboBox = new DataComboBox(h1ComboBox, h2ComboBox, lComboBox, pComboBox, tComboBox, cComboBox);
            for (int i = elementPanel.ColumnCount - 1; i < elementPanel.Controls.Count - 1; i++)
            {
                ComboBox cmbBox = (ComboBox)elementPanel.Controls[i];
                cmbBox.Items.Clear();
            }
        }
    }
}
