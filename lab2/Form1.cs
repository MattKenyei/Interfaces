using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Spire.Xls;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace interface_lab2
{
    public partial class FormTryes : Form
    {
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);

        Stopwatch stopwatch = new Stopwatch();
        static int NumButtons = 9, NumTries = 5;
        int NumVisible = 2, CurTry = 1;
        bool IsStarted = false, IsFinished = false;
        string CurWord = "";
        static string WordString = "Word: ", TryBeginString = "Try: ", TryEndString = " / ", TimeString = "time: ";

        SortedList<int, string> Words = new SortedList<int, string>()
        {
            {1, "A"},
            {2, "B"},
            {3, "C"},
            {4, "D"},
            {5, "E"},
            {6, "F"},
            {7, "G"},
            {8, "H"},
            {9, "I"}
        };
        List<Button> buttons = new List<Button>();

        SortedList<int, float> MedTimes = new SortedList<int, float>();
        List<float> times = new List<float>();

        public FormTryes()
        {
            InitializeComponent();

            buttons.Add(button1);
            buttons.Add(button2);
            buttons.Add(button3);
            buttons.Add(button4);
            buttons.Add(button5);
            buttons.Add(button6);
            buttons.Add(button7);
            buttons.Add(button8);
            buttons.Add(button9);
        }

        private void OutData()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            worksheet.Range[1, 1].Value = "n";
            worksheet.Range[1, 2].Value = "t";

            for (int i = 1; i <= MedTimes.Count(); i++)
            {
                worksheet.Range[i + 1, 2].Value = MedTimes.ElementAt(i - 1).Value.ToString();
                worksheet.Range[i + 1, 1].Value = MedTimes.ElementAt(i - 1).Key.ToString();
            }

            workbook.SaveToFile("C:\\Users\\maksp\\OneDrive\\Рабочий стол\\interface lab2\\1.xlsx", ExcelVersion.Version2013);
        }

        private void RightClick()
        {
            stopwatch.Stop();
            times.Add(((float)(stopwatch.Elapsed.Seconds + stopwatch.Elapsed.Milliseconds * 0.001)));

            foreach (Button button in buttons) button.Visible = false;

            if (CurTry < NumTries) CurTry++;
            else {
                MedTimes.Add(NumVisible, times.Sum()/NumTries);
                times.Clear();
                CurTry = 1;
                NumVisible++;
            };

            buttonStart.Visible = true;

            labelTime.Text = TimeString + stopwatch.Elapsed.Seconds.ToString() + "," + stopwatch.Elapsed.Milliseconds.ToString();
            labelTime.Visible = true;

            if (NumVisible > NumButtons) {
                OutData();
                this.Close(); 
            };
            IsStarted = false;
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            Random random = new Random();

            List<int> vectorNum = new List<int>();

            int num = random.Next(1, NumVisible + 1);
            for (int i = 0; i < NumVisible; i++)
            {
                while (vectorNum.Contains(num)) num = random.Next(1, NumVisible + 1);
                vectorNum.Add(num);
            }

            for (int i = 0; i < NumVisible; i++) {
                string str = "";
                Words.TryGetValue(vectorNum.ElementAt(i), out str);
                buttons.ElementAt(i).Text = str;
            }

            buttonStart.Visible = false;

            for (int i = 0; i < NumVisible; i++) buttons[i].Visible = true;

            IsStarted = true;

            Words.TryGetValue(random.Next(1, NumVisible+1), out CurWord);
            labelWord.Text = WordString + CurWord;

            labelTries.Text = TryBeginString + CurTry + TryEndString + NumTries;

            Point buttonLocation = buttons[0].PointToScreen(Point.Empty);
            buttonLocation.X = buttons[0].PointToScreen(Point.Empty).X + buttons[0].Size.Width + 10;
            buttonLocation.Y = buttons[0].PointToScreen(Point.Empty).Y + (buttons[0].Size.Height + 5) * NumVisible / 2;
            SetCursorPos(buttonLocation.X, buttonLocation.Y);

            labelTime.Visible = false;

            stopwatch.Reset();
            stopwatch.Start();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (IsStarted &&  CurWord == button1.Text) RightClick();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (IsStarted && CurWord == button2.Text) RightClick();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (IsStarted && CurWord == button3.Text) RightClick();
        }

        private void FormTryes_Load(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (IsStarted && CurWord == button4.Text) RightClick();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (IsStarted && CurWord == button5.Text) RightClick();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (IsStarted && CurWord == button6.Text) RightClick();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (IsStarted && CurWord == button7.Text) RightClick();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (IsStarted && CurWord == button8.Text) RightClick();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (IsStarted && CurWord == button9.Text) RightClick();
        }
    }
}