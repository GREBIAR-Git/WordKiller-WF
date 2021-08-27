namespace MakeReportWord
{
    class CreatedElements
    {
        int h1, h2, l, p, t, c;
        public void Add(string str)
        {
            if(str == "heading1ComboBox")
            {
                h1++;
            }
            else if (str == "heading2ComboBox")
            {
                h2++;
            }
            else if (str == "listComboBox")
            {
                l++;
            }
            else if (str == "pictureComboBox")
            {
                p++;
            }
            else if (str == "tableComboBox")
            {
                t++;
            }
            else if (str == "codeComboBox")
            {
                c++;
            }
        }

        public void Del(string str)
        {
            if (str == "heading1ComboBox" && h1 > 0)
            {
                h1--;
            }
            else if (str == "heading2ComboBox" && h2 > 0)
            {
                h2--;
            }
            else if (str == "listComboBox" && l > 0)
            {
                l--;
            }
            else if (str == "pictureComboBox" && p > 0)
            {
                p--;
            }
            else if (str == "tableComboBox" && t > 0)
            {
                t--;
            }
            else if (str == "codeComboBox" && c > 0)
            {
                c--;
            }
        }

        public int H1
        {
            get { return h1; }
            set { h1 = value; }
        }
        public int H2
        {
            get { return h2; }
            set { h2 = value; }
        }
        public int L
        {
            get { return l; }
            set { l = value; }
        }
        public int P
        {
            get { return p; }
            set { p = value; }
        }
        public int T
        {
            get { return t; }
            set { t = value; }
        }
        public int C
        {
            get { return c; }
            set { c = value; }
        }
        public int sum()
        {
            return h1 + h2 + l + p + t + c;
        }
        public CreatedElements()
        {
            h1 = 0; h2 = 0; l = 0; p = 0; t = 0; c=0;
        }
    }
}
