using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MakeReportWord
{
    public class ArrayCopy
    {
        public static T[] CopyArray<T>(T[] array, int startElem, int endElem)
        {
            T[] newArray = new T[array.Length];
            for (int i = startElem; i <= endElem; i++)
            {
                newArray[i] = array[i];
            }
            return newArray;
        }
    }
}
