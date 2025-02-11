using Microsoft.VisualBasic.ApplicationServices;
using OfficeOpenXml;  //使用EPPlus
using System.Speech.Synthesis;  //微軟的AI合成聲音

namespace CallStudents
{
    public partial class Form1 : Form
    {
        //保存synthesizer資料
        private SpeechSynthesizer synthesizer;
        private string msg = "您的家長在等您";

        public Form1()
        {
            //.NET平台自動執行，主要的工作是做一些初始化的工作。
            InitializeComponent();

            //創建實體，這樣才可為NULL
            synthesizer = new SpeechSynthesizer();

            // 如果您在非商業環境中根據 Polyform Noncommercial 授權使用 EPPlus：
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //使用者路徑
            string excelPath = Path.Combine("students.xlsx");

            //開發路徑
            //string excelPath = Path.Combine(@"..\..\..\Excel\students.xlsx");
            //("..\\..\\..\\Excel\\students.xlsx") 另一種寫法

            if (File.Exists(excelPath))
            {
                try
                {
                    GenerateButtonsFromExcel(excelPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("讀取 Excel 文件時發生錯誤: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Excel 文件未找到: " + excelPath);
            }
        }

        private void GenerateButtonsFromExcel(string excelFilePath)
        {
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int classCount = worksheet.Dimension.Columns; // 獲取班級數量
                int rowCount = worksheet.Dimension.Rows;      // 獲取學生數量

                int labelTopOffset = 50; // 第一個班級標籤的垂直偏移量
                int labelLeftOffset = 28; // 標籤的水平偏移量
                int buttonTopOffset = 80; // 按鈕的垂直間距
                int buttonLeftOffset = 15; // 按鈕的水平偏移量

                for (int col = 1; col <= classCount; col++)
                {
                    string className = worksheet.Cells[1, col].Text; // 獲取班級名稱
                    CreateLabelForClass(className, col, labelTopOffset, labelLeftOffset);

                    // 顯示每個班級的學生按鈕
                    for (int row = 2; row <= rowCount; row++)
                    {
                        string studentName = worksheet.Cells[row, col].Text; // 獲取學生姓名

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
            classLabel.Left = labelLeftOffset + (classIndex - 1) * (classLabel.Width + 10); // 設置水平位置

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
            studentButton.Top = buttonTopOffset + (rowIndex - 2) * (studentButton.Height + 10);   // 垂直間距
            studentButton.Left = buttonLeftOffset + (classIndex - 1) * (studentButton.Width + 10);    // 水平間距

            // 當按鈕被點擊時
            studentButton.Click += async (sender, _) =>
            {
                // 將 sender 轉換為 Button 類型，若為 null 則設置為默認按鈕
                var button = sender as Button ?? new Button();

                // 禁用按鈕
                button.Enabled = false;

                try
                {
                    // 執行播放姓名的異步操作
                    await PlayStudentNameAsync(studentName + msg);
                }
                finally
                {
                    // 播放完成後，重新啟用按鈕
                    button.Enabled = true;
                }
            };

            this.Controls.Add(studentButton);
        }

        // SemaphoreSlim 用來控制同一時刻允許多少個線程的輕量級信號量
        // SemaphoreSlim(初始計數,最大計數)，也能直接寫(1)
        private SemaphoreSlim _semaphore = new SemaphoreSlim(1,1);

        //播放學生名稱方法
        private async Task PlayStudentNameAsync(string name)
        {
            // WaitAsync是SemaphoreSlim的非同步方法。確保在同一時間只允許一個線程進行播放操作。
            await _semaphore.WaitAsync();

            try
            {
                // 創建一個TaskCompletionSource，用來在Speak完成時通知異步任務。
                var tcs = new TaskCompletionSource<object?>();

                // 用來處理 synthesizer.SpeakCompleted 事件的委派 handler是變數
                EventHandler<SpeakCompletedEventArgs>? handler = null;
                //(s觸發者，e參數)
                handler = (s, e) =>
                {
                    // 移除SpeakCompleted事件處理器，避免重複調用。
                    synthesizer.SpeakCompleted -= handler;
                    // 設置TaskCompletionSource的結果，設成NULL解除Task的等待狀態。
                    tcs.SetResult(null);
                };
                // 添加SpeakCompleted事件處理器。
                synthesizer.SpeakCompleted += handler;

                // 調用synthesizer的異步方法，參數(名稱)
                synthesizer.SpeakAsync(name);

                // 等待TaskCompletionSource完成，即語音合成播放結束。
                await tcs.Task;
            }
            finally
            {
                // 釋放SemaphoreSlim信號
                _semaphore.Release();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //按紐click簡單功能
        //private async void button1_Click(object sender, EventArgs e)
        //{
        //    //按下後關閉按鈕
        //    button1.Enabled = false;

        //    try
        //    {
        //        await Task.Run(() => PlayStudentName("張詠萱" + message));
        //    }
        //    finally
        //    {
        //        //開啟按鈕
        //        button1.Enabled = true;
        //    }

        //}

    }
}