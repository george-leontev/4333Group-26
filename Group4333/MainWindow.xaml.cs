using System.Windows;

namespace Group4333
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();

        private void ShowAuthorInfo_Click(object sender, RoutedEventArgs e)
        {
            var authorWindow = new Group4333_Leontev();
            authorWindow.ShowDialog();
        }
    }
}