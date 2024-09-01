using Microsoft.VisualBasic.ApplicationServices;
using OfficeOpenXml;  //�ϥ�EPPlus
using System.Speech.Synthesis;  //�L�n��AI�X���n��

namespace CallStudents
{
    public partial class Form1 : Form
    {
        //�O�ssynthesizer���
        private SpeechSynthesizer synthesizer;
        private string msg = "�z���a���b���z";

        public Form1()
        {
            InitializeComponent();
            //�Ыع���A�o�ˤ~�i��NULL
            synthesizer = new SpeechSynthesizer();

            // �p�G�z�b�D�ӷ~���Ҥ��ھ� Polyform Noncommercial ���v�ϥ� EPPlus�G
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //�ϥΪ̸��|
            string excelPath = Path.Combine("students.xlsx");

            //�}�o���|
            //string excelPath = Path.Combine(@"..\..\..\Excel\students.xlsx");
            //("..\\..\\..\\Excel\\students.xlsx") �t�@�ؼg�k

            if (File.Exists(excelPath))
            {
                try
                {
                    GenerateButtonsFromExcel(excelPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ū�� Excel ���ɵo�Ϳ��~: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Excel ��󥼧��: " + excelPath);
            }
        }

        private void GenerateButtonsFromExcel(string excelFilePath)
        {
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int classCount = worksheet.Dimension.Columns; // ����Z�żƶq
                int rowCount = worksheet.Dimension.Rows;      // ����ǥͼƶq

                int labelTopOffset = 50; // �Ĥ@�ӯZ�ż��Ҫ����������q
                int labelLeftOffset = 28; // ���Ҫ����������q
                int buttonTopOffset = 80; // ���s���������Z
                int buttonLeftOffset = 15; // ���s�����������q

                for (int col = 1; col <= classCount; col++)
                {
                    string className = worksheet.Cells[1, col].Text; // ����Z�ŦW��
                    CreateLabelForClass(className, col, labelTopOffset, labelLeftOffset);

                    // ��ܨC�ӯZ�Ū��ǥͫ��s
                    for (int row = 2; row <= rowCount; row++)
                    {
                        string studentName = worksheet.Cells[row, col].Text; // ����ǥͩm�W

                        if (!string.IsNullOrEmpty(studentName))
                        {
                            CreateButtonForStudent(studentName, col, row, buttonTopOffset, buttonLeftOffset);
                        }
                    }
                }
            }
        }

        private void CreateLabelForClass(string className, int classIndex, int topOffset, int labelLeftOffset)
        {
            Label classLabel = new Label
            {
                Text = className,
                Top = topOffset,
                AutoSize = true,
                Font = new Font("Microsoft YaHei", 12)
            };
            classLabel.Left = labelLeftOffset + (classIndex - 1) * (classLabel.Width + 10); // �]�m������m

            this.Controls.Add(classLabel);
        }

        private void CreateButtonForStudent(string studentName, int classIndex, int rowIndex, int buttonTopOffset, int buttonLeftOffset)
        {

            Button studentButton = new Button
            {
                Text = studentName,
                Width = 100,
                Height = 50,
                Font = new Font("Microsoft YaHei", 18)
            };
            studentButton.Top = buttonTopOffset + (rowIndex - 2) * (studentButton.Height + 10);   // �������Z
            studentButton.Left = buttonLeftOffset + (classIndex - 1) * (studentButton.Width + 10);    // �������Z

            // ����s�Q�I����
            studentButton.Click += async (sender, _) =>
            {
                // �N sender �ഫ�� Button �����A�Y�� null �h�]�m���q�{���s
                var button = sender as Button ?? new Button();

                // �T�Ϋ��s
                button.Enabled = false;

                try
                {
                    // ���漽��m�W�����B�ާ@
                    await PlayStudentNameAsync(studentName + msg);
                }
                finally
                {
                    // ���񧹦���A���s�ҥΫ��s
                    button.Enabled = true;
                }
            };

            this.Controls.Add(studentButton);
        }

        // �w�q�@��SemaphoreSlim����A�Ψӱ����@�ɸ귽���õo�s���C
        // ���B�]�m��1�A��̦ܳh�u���@�Ӱ�����i�H�P�ɶi�J�{�ɰϰ�C
        private SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);

        //����ǥͦW�٤�k
        private async Task PlayStudentNameAsync(string name)
        {
            // �������SemaphoreSlim����A�p�G����L������b�ϥθӸ귽�A�h��e������|�Q����C
            await _semaphore.WaitAsync();

            try
            {
                // �Ыؤ@��TaskCompletionSource�A�ΨӦbSpeak�����ɳq�����B���ȡC
                var tcs = new TaskCompletionSource<object?>();

                // �w�qSpeakCompleted�ƥ�B�z���A�ΨӦb�y���X�������ɳ]�mTaskCompletionSource�����G�C
                EventHandler<SpeakCompletedEventArgs>? handler = null;
                handler = (s, e) =>
                {
                    // ����SpeakCompleted�ƥ�B�z���A�קK���ƽեΡC
                    synthesizer.SpeakCompleted -= handler;
                    // �]�mTaskCompletionSource�����G�A�Ѱ�Task�����ݪ��A�C
                    tcs.SetResult(null);
                };
                // ���USpeakCompleted�ƥ�B�z���C
                synthesizer.SpeakCompleted += handler;

                // ���B�a�}�l�y���X���A�ǻ��ǥͩm�W�i��y������C
                synthesizer.SpeakAsync(name);

                // ����TaskCompletionSource�����A�Y�y���X�����񵲧��C
                await tcs.Task;
            }
            finally
            {
                // ����SemaphoreSlim����A���\��L���ݪ�������i�J�{�ɰϰ�C
                _semaphore.Release();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //����click²��\��
        //private async void button1_Click(object sender, EventArgs e)
        //{
        //    //���U���������s
        //    button1.Enabled = false;

        //    try
        //    {
        //        await Task.Run(() => PlayStudentName("�i����" + message));
        //    }
        //    finally
        //    {
        //        //�}�ҫ��s
        //        button1.Enabled = true;
        //    }

        //}

    }
}