using System.Speech.Synthesis;  //微軟的AI合成聲音

namespace CallStudents
{
    public partial class Form1 : Form
    {
        //保存synthesizer資料
        private SpeechSynthesizer synthesizer;
        public string message = "您的家長在等您";

        public Form1()
        {
            InitializeComponent();
            //創建實體，這樣才可為NULL
            synthesizer = new SpeechSynthesizer();
        }

        //播放學生名稱方法
        private void PlayStudentName(string name)
        {
            synthesizer.Speak(name);
        }


        private async void button1_Click(object sender, EventArgs e)
        {
            //按下後關閉按鈕
            button1.Enabled = false;

            try
            {
                await Task.Run(() => PlayStudentName("張詠萱" + message));
            }
            finally
            {
                //開啟按鈕
                button1.Enabled = true;
            }
            
        }

    }
}