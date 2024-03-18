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
    /// Interaction logic for Login.xaml
    /// </summary>
    public partial class Login : Window
    {
        public string token = "";
        public bool ok = false;
        private int failedNum = 0;
        public Login()
        {
            InitializeComponent();
        }
        private async void button1_Click(object sender, RoutedEventArgs e)
        {

            if (this.textBoxEmail.Text.Length != 0 && this.passwordBox1.Password.Length != 0)
            {
                KLICApi api = new KLICApi();
                string portalUrl = "https://warmtestad.maps.arcgis.com/";
                var result = await api.getToken(username: this.textBoxEmail.Text.ToLower(), psw: this.passwordBox1.Password, portalUrl: portalUrl);
                if (result != null & result != "")
                {
                    this.token = result;
                    ok = true;
                    this.Close();
                }
                else
                {
                    ok = false;
                    errormessage.Text = "De inloggegevens zijn niet correct";
                    this.textBoxEmail.Text = "";
                    this.passwordBox1.Password = "";
                    failedNum++;
                }


            }
            else
            {
                errormessage.Text = "Vul de inloggegevens in";
            }

            if (failedNum > 3)
            {
                DialogResult = false;
                this.Close();
            }
        }

    }
}
