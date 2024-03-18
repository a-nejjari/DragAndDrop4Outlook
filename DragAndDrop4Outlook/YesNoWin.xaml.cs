using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace DragAndDrop4Outlook
{
    /// <summary>
    /// Interaction logic for YesNoWin.xaml
    /// </summary>
    public partial class YesNoWin : Window
    {
        public string fileName { get; set; }
        public YesNoWin(string email, string KLICMelding, string FileName)
        {

            InitializeComponent();
            this.TExtBlockEmailSubject.Text = email;
            this.TextBlockKlickMelding.Text = KLICMelding;
            this.TextBoxFileName.Text = FileName;
        }

        private void yes_button_Click(object sender, RoutedEventArgs e)
        {
            this.fileName = this.TextBoxFileName.Text;
            DialogResult = true;
            this.Close();
        }

        private void no_button1_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            this.Close();
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            this.fileName = TextBoxFileName.Text;
            if (TextBoxFileName.Text.Length < 10) this.Yes_Button.IsEnabled = false;
            else this.Yes_Button.IsEnabled = true;
        }

    }
}
