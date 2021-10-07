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
            string hexStr = BitConverter.ToString(Encoding.Unicode.GetBytes(str));
            hexStr = hexStr.Replace("-", "");
            string binarystring = String.Join(String.Empty, hexStr.Select(c => Convert.ToString(Convert.ToInt32(c.ToString(), 16), 2).PadLeft(4, '0')));
            return binarystring;
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

            binary.Replace("\r", "");
            string[] binaryLines = binary.Split('\n');
            
            foreach(string line in binaryLines)
            {
                byte[] bytes = convertToBytes(line);
                normal += Convert.ToBase64String(bytes);
                normal += "\n";
            }

            return normal;
        }

        public static string RepeatEncodingBinary(string binarystring)
        {
            string encoded = "";
            for(int i=0;i< binarystring.Length-1;i++)
            {
                int f=1;
                for (;f<9&& binarystring.Length>i+f; f++)
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