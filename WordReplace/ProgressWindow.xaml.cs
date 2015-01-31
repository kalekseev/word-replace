using System.ComponentModel;
using System.Windows;


namespace WordReplace
{
    public partial class ProgressWindow : Window
    {
        private int Counter = 0;
        private static object syncCounter = new object();

        public ProgressWindow()
        {
            InitializeComponent();
        }

        public void Update(object sender, ProgressChangedEventArgs e) {
            lock(syncCounter)
                Counter++;
            pbStatus.Value = 100 * Counter / (double)e.ProgressPercentage;
        }

        public void Reset()
        {
            Counter = 0;
            pbStatus.Value = 0;
        }
    }
}
