using System.Windows.Forms;

namespace MakeReportWord
{
    public partial class Question : Form
    {
        public Question(string question,string btLeft,string btRight)
        {
            InitializeComponent();
            labelQuestion.Text = question;
            buttonLeft.Text = btLeft;
            buttonRight.Text = btRight;
        }
    }
}
