using System;
using System.Windows.Forms;

namespace lab8_2
{
    public partial class MySplashScreen : Form
    {
        Form formCallback ;

        public MySplashScreen()
        {
            InitializeComponent();
            this.CenterToScreen();
            var t = new Timer();

            t.Interval = 20000;

            t.Start();

            t.Tick += new EventHandler(t_Tick);

            t.Start();
            Opacity = 0;

            var timer = new Timer();

            timer.Tick += new EventHandler((sender, e) =>

            {

                if ((Opacity += 0.005d) == 1)
                    timer.Stop();

            });

            timer.Interval = 1;

            timer.Start();

        }

        private void t_Tick(object sender, EventArgs e)
        {


            ((Timer)sender).Enabled = false;
            this.Hide();
            new Form1().ShowDialog();
            Application.Exit();
        }
    }

    


}
