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
using System.Linq.Expressions;

namespace interface_lab2
{
    public partial class FormTryes : Form
    {
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);
        enum Changes
        {
            color,
            size
        }

        Changes flag = Changes.color;
        Stopwatch stopwatch = new Stopwatch();
        static int NumButtons = 9, NumTries = 5;
        int NumVisible = 2, CurTry = 1, NumExp = 0;
        bool IsStarted = false, isEntered = false;
        string CurWord = "";
        List<Tuple<int, string, float>> MedTimesSecondDraw = new List<Tuple<int, string, float>>();
        List<Tuple<int, int, float>> MedTimesSecondSize = new List<Tuple<int, int, float>>();

        static string WordString = "Letter: ", TryBeginString = "Try: ", TryEndString = " / ", TimeString = "time: ", colorString = "Цвет объекта", sizeString = "Размер объекта";

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
        static List<Color> colors = new List<Color>()
        {
            Color.Aqua, Color.Blue, Color.Black, Color.Orange
        };
        static List<int> sizes = new List<int>()
        {
            40,
            50,
            60,
            70
        };
        int curColor = 0;
        int curSize = 0;
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
        private void OutDataSecond()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[1];

            int cur = 0;

            foreach (var font in colors)
            {
                worksheet.Range[1, 1 + cur * 4].Value = "color";
                worksheet.Range[2, 1 + cur * 4].Value = font.ToString();
                worksheet.Range[3, 1 + cur * 4].Value = "n";
                worksheet.Range[3, 2 + cur * 4].Value = "t";

                int curi = 0;
                for (int i = 0; i < MedTimesSecondDraw.Count(); i++)
                {
                    if (font.ToString() == MedTimesSecondDraw.ElementAt(i).Item2)
                    {
                        worksheet.Range[curi + 4, 2 + cur * 4].Value = MedTimesSecondDraw.ElementAt(i).Item3.ToString();
                        worksheet.Range[curi + 4, 1 + cur * 4].Value = MedTimesSecondDraw.ElementAt(i).Item1.ToString();
                        curi++;
                    }
                }
                cur++;
            }

            worksheet = workbook.Worksheets[2];

            cur = 0;

            foreach (var size in sizes)
            {
                worksheet.Range[1, 1 + cur * 4].Value = "size";
                worksheet.Range[2, 1 + cur * 4].Value = size.ToString();
                worksheet.Range[3, 1 + cur * 4].Value = "n";
                worksheet.Range[3, 2 + cur * 4].Value = "t";

                int curi = 0;
                for (int i = 0; i < MedTimesSecondSize.Count(); i++)
                {
                    if (size == MedTimesSecondSize.ElementAt(i).Item2)
                    {
                        worksheet.Range[curi + 4, 2 + cur * 4].Value = MedTimesSecondDraw.ElementAt(i).Item3.ToString();
                        worksheet.Range[curi + 4, 1 + cur * 4].Value = MedTimesSecondDraw.ElementAt(i).Item1.ToString();
                        curi++;
                    }
                }
                cur++;
            }

            workbook.SaveToFile("C:\\Users\\maksp\\OneDrive\\Рабочий стол\\interface lab2\\2.xlsx", ExcelVersion.Version2013);
        }
        private void RightClick()
        {
            if (isEntered)
            {
                numericUpDownExperiment.Visible = false;
                labelExperiment.Visible = false;

                stopwatch.Stop();
                times.Add(((float)(stopwatch.Elapsed.Seconds + stopwatch.Elapsed.Milliseconds * 0.001)));

                foreach (Button button in buttons) button.Visible = false;

                if (CurTry < NumTries) CurTry++;
                else
                {
                    if (NumExp == 1)
                    {
                        MedTimes.Add(NumVisible, times.Sum() / NumTries);
                        NumVisible++;
                    }
                    else if (NumExp == 2)
                    {
                        if (flag == Changes.color)
                        {
                            MedTimesSecondDraw.Add(new Tuple<int, string, float>(NumVisible, colors[curColor].ToString(), times.Sum() / NumTries));
                            if (curColor == colors.Count() - 1)
                            {
                                flag = Changes.size;
                                curColor = 0;
                            }
                            else curColor++;
                        }
                        else if (flag == Changes.size)
                        {
                            MedTimesSecondSize.Add(new Tuple<int, int, float>(NumVisible, sizes[curSize], times.Sum() / NumTries));
                            if (curSize == sizes.Count() - 1)
                            {
                                flag = Changes.color;
                                NumVisible++;
                                curSize = 0;
                            }
                            else curSize++;
                        }
                    }
                    times.Clear();
                    CurTry = 1;
                };

                buttonStart.Visible = true;

                labelTime.Text = TimeString + stopwatch.Elapsed.Seconds.ToString() + "," + stopwatch.Elapsed.Milliseconds.ToString();
                labelTime.Visible = true;

                if (NumVisible > NumButtons)
                {
                    if (NumExp == 1)
                        OutData();
                    else if (NumExp == 2)
                        OutDataSecond();
                    this.Close();
                };
            }
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            if (isEntered)
            {
                Random random = new Random();

                List<int> vectorNum = new List<int>();

                int num = random.Next(1, NumVisible + 1);
                for (int i = 0; i < NumVisible; i++)
                {
                    while (vectorNum.Contains(num)) num = random.Next(1, NumVisible + 1);
                    vectorNum.Add(num);
                }

                for (int i = 0; i < NumVisible; i++)
                {
                    string str = "";
                    Words.TryGetValue(vectorNum.ElementAt(i), out str);
                    buttons.ElementAt(i).Text = str;
                }

                buttonStart.Visible = false;

                for (int i = 0; i < NumVisible; i++) buttons[i].Visible = true;

                IsStarted = true;

                Words.TryGetValue(random.Next(1, NumVisible + 1), out CurWord);
                if (NumExp == 1)
                    labelWord.Text = WordString + CurWord;
                else
                    if (flag == Changes.color) labelWord.Text = colorString;
                else if (flag == Changes.size) labelWord.Text = sizeString;

                labelTries.Text = TryBeginString + CurTry + TryEndString + NumTries;

                Point buttonLocation = buttons[0].PointToScreen(Point.Empty);
                buttonLocation.X = buttons[0].PointToScreen(Point.Empty).X + buttons[0].Size.Width + 10;
                buttonLocation.Y = buttons[0].PointToScreen(Point.Empty).Y + (buttons[0].Size.Height + 5) * NumVisible / 2;
                SetCursorPos(buttonLocation.X, buttonLocation.Y);

                labelTime.Visible = false;
                if(NumExp == 2)
                {
                    foreach (var button in buttons)
                    {
                        if (button.Text != CurWord)
                        {
                            button.BackColor = Color.Green;
                            button.Height = 34;
                        }
                        else
                        {
                            if (flag == Changes.color)
                                button.BackColor = colors[curColor];
                            else
                                button.Width = sizes[curSize];
                        }
                    }
                }

                stopwatch.Reset();
                stopwatch.Start();
            }
            else
            {
                NumExp = ((int)numericUpDownExperiment.Value);
                isEntered = true;
                numericUpDownExperiment.Visible = false;
                labelExperiment.Visible = false;
                labelWord.Visible = true;

                labelTries.Visible = true;
            }
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

        private void numericUpDownExperiment_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDownExperiment.Value == 1 || numericUpDownExperiment.Value == 2)
            {
                labelExperiment.Text = "подтвердите номер эксперимента:";
                if (!buttonStart.Enabled) buttonStart.Enabled = true;
            }
            else
            {
                labelExperiment.Text = "некорректный номер, введите другой:";
                if (buttonStart.Enabled) buttonStart.Enabled = false;
            }
        }

        private void labelExperiment_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (IsStarted && CurWord == button4.Text) RightClick();
        }

        private void labelWord_Click(object sender, EventArgs e)
        {

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