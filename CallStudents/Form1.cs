using System.Speech.Synthesis;  //�L�n��AI�X���n��

namespace CallStudents
{
    public partial class Form1 : Form
    {
        //�O�ssynthesizer���
        private SpeechSynthesizer synthesizer;
        public string message = "�z���a���b���z";

        public Form1()
        {
            InitializeComponent();
            //�Ыع���A�o�ˤ~�i��NULL
            synthesizer = new SpeechSynthesizer();
        }

        //����ǥͦW�٤�k
        private void PlayStudentName(string name)
        {
            synthesizer.Speak(name);
        }


        private async void button1_Click(object sender, EventArgs e)
        {
            //���U���������s
            button1.Enabled = false;

            try
            {
                await Task.Run(() => PlayStudentName("�i����" + message));
            }
            finally
            {
                //�}�ҫ��s
                button1.Enabled = true;
            }
            
        }

    }
}