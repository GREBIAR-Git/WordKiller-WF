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
            return sb.ToString();
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