namespace WordKiller
{
    class StringArraySearcher
    {
        public static int IndexOf(string str, string[] array)
        {
            int index = -1;
            for (int i = 0; i < array.Length; i++)
            {
                if (array[i].Contains(str))
                {
                    index = i;
                    break;
                }
            }
            return index;
        }
    }
}
