using System;
using System.Linq;
using System.Text;

namespace MakeReportWord
{
    class EncodingDecoding
    {
        public static char DigitToLetter(char digit)
        {
            string abc = "abcdefghij";
            return abc[int.Parse(digit.ToString())];
        }

        public static char LetterToDigit(char letter)
        {
            string abc = "abcdefghij";
            if (letter != '\r'&&letter!='\n')
            {
                return abc.IndexOf(letter).ToString()[0];
            }
            return letter;
        }

        public static string DigitsToAbc(string digits)
        {
            string abc = "";
            for (int i = 0; i < digits.Length; i++)
            {
                abc += DigitToLetter(digits[i]);
            }
            return abc;
        }

        public static string AbcToDigits(string abc)
        {
            string digits = "";
            for (int i = 0; i < abc.Length; i++)
            {
                digits += LetterToDigit(abc[i]);
            }
            return digits;
        }

        public static string StringToBinaryString(string str)
        {
            StringBuilder sb = new StringBuilder();
            foreach (byte b in System.Text.Encoding.UTF8.GetBytes(str))
                sb.Append(Convert.ToString(b, 2).PadLeft(8, '0'));
            string binaryStr = sb.ToString();
            return sb.ToString();
        }

        public static byte[] convertToBytes(string s)
        {
            byte[] result = new byte[(s.Length + 7) / 8];

            int i = 0;
            int j = 0;
            foreach (char c in s)
            {
                result[i] <<= 1;
                if (c == '1')
                    result[i] |= 1;
                j++;
                if (j == 8)
                {
                    i++;
                    j = 0;
                }
            }
            return result;
        }


        public static string BinaryStringToString(string binary)
        {
            string normal = "";

            binary = binary.Replace("\r", "");
            string[] binaryLines = binary.Split('\n');
            
            foreach(string line in binaryLines)
            {
                var stringArray = Enumerable.Range(0, line.Length / 8).Select(i => Convert.ToByte(line.Substring(i * 8, 8), 2)).ToArray();
                normal += Encoding.UTF8.GetString(stringArray);
                normal += "\n";
            }
            normal = normal.Remove(normal.Length-1);
            normal = normal.Replace("\n","\r\n");

            return normal;
        }

        public static string RepeatEncodingBinary(string binarystring)
        {
            string encoded = "";
            for(int i=0;i< binarystring.Length;i++)
            {
                int f=1;
                for (;f<10&& binarystring.Length>i+f; f++)
                {
                    if(binarystring[i] != binarystring[f+i])
                    {
                        break;
                    }
                }
                encoded += (f).ToString() + binarystring[i];
                i += f-1;
            }
            return encoded;
        }

        public static string RepeatDecodingBinary(string repeated_digit)
        {
            string decoded = "";
            for (int index = 0; index + 1 < repeated_digit.Length; index+=2)
            {
                string digit = repeated_digit[index].ToString();
                if (digit != "\r" && digit != "\n")
                {
                    for (int i = 0; i < int.Parse(repeated_digit[index].ToString()); i++)
                    {
                        decoded += repeated_digit[index + 1];
                    }
                }
                else
                {
                    string digit2 = repeated_digit[index + 1].ToString();
                    if (digit2 == "\r" || digit2 == "\n")
                    {
                        decoded += digit+ digit2;
                    }
                    else
                    {
                        index--;
                        decoded += digit;
                    }
                }
            }
            return decoded;
        }
    }
}
//"☺Menu☺LabMenuItem!Trud����������ބ���1u�q�{�\t�}�\u0001�y�a�\r�q�}�{�{�\u0017�\nA�\u0003�q�\u0003�\u0005�k�xA�pA�\r�q�\t�\u0001�}�e�\u0017�\nA�\u0005�k�\v�{�}�w�}�g�q�\nnumberTextBox�\n"
//"☺Menu☺LabMenuItem!Trud����������ބ���1u�q�{�\t�}�\u0001�y�a�\r�q�}�{�{�\u0017�\nA�\u0003�q�\u0003�\u0005�k�xA�pA�\r�q�\t�\u0001�}�e�\u0017�\nA�\u0005�k�\v�{�}�w�}�g�q�\nnumberTextBox�\nthemeTextBox�\ndisciplineTextBox�\nprofessorComboBox�\nyearTextBox☺20\nshifrTextBox�\nstudentsTextBox�\n☺TextStart�\nddd\n☺TextEnd�\n\n"

//111 000 1 0 1 00 11 000 101110100100110101100101011011100111010111100010100110001011101001001100011000010110001001001101011001010110111001110101010010010111010001100101011011010010000101010100011100100111010101100101
//31 30 11 10 11 20 21 30 11 10 31 10 11 20 11 20 21 10 11 10 21 20 11 10 11 10 21 10 31 20 31 10 11 10 41 30 11 10 11 20 21 30 11 10 31 10 11 20 11 20 213021401110213011201120211011102120111011102110312031101110111011201120111031101130212011101110211021101120114011101110111011303120112031101110111021201110
//3130111011202130111031101120112021101110212011101110211031203110111041301110112021301110311011201120213021401110213011201120211011102120111011102110312031101110111011201120111031101130212011101110211021101120114011101110111011303120112031101110111021201110