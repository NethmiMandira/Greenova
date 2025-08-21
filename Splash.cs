using System;
using System.Windows.Forms;

namespace Geenova
{
    public partial class Splash : Form
    {
        private Timer splashTimer;

        public Splash()
        {
            InitializeComponent();
        }

        private void Splash_Load(object sender, EventArgs e)
        {
            splashTimer = new Timer();
            splashTimer.Interval = 5000; // 5 seconds = 5000 milliseconds
            splashTimer.Tick += SplashTimer_Tick;
            splashTimer.Start();
        }

        private void SplashTimer_Tick(object sender, EventArgs e)
        {
            splashTimer.Stop();
            splashTimer.Dispose();

            Login l = new Login();
            l.Show();// Show the main form or the next form in your application

            this.Hide(); // or this.Close(); depending on app behavior
        }
    }
}
