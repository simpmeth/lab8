using System;
using System.Threading;
using System.Windows.Forms;

namespace lab8_2
{
    public partial class SplashScreen : Form
    {
        public SplashScreen()
        {
            InitializeComponent();
           
        }

        private void SplashScreen_Load(object sender, EventArgs e)
        {
            CenterToScreen();
            var thread = new Thread(new ThreadStart(show));
            thread.IsBackground = true;
            this.Hide();
            thread.Start();
            
        }
     

        private const int WS_SYSMENU = 0x80000;
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.Style &= ~WS_SYSMENU;
                return cp;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        void show()
        {
            Thread.Sleep(1000);
            var form1 = new Form1();
            form1.ShowDialog();
        }
    }
}
