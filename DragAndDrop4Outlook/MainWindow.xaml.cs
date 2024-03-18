using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Outlook = Microsoft.Office.Interop.Outlook;
using Path = System.IO.Path;

namespace DragAndDrop4Outlook
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string environmentPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string dataDirName = "DragDropMails";
        string dataDirPath = "";
        string token = "";
        //string portalUrl = "https://services-eu1.arcgis.com/fpQ6Cf1ZDEH1n6PO/arcgis/rest/services/Klic_Meldingen_T_V2/FeatureServer/0/";
        string portalUrl = "https://services-eu1.arcgis.com/fpQ6Cf1ZDEH1n6PO/arcgis/rest/services/Klic_Meldingen_P/FeatureServer/20/";
        public MainWindow()
        {
            dataDirPath = Path.Combine(environmentPath, dataDirName);
            if (!Directory.Exists(dataDirPath))
            {
                Directory.CreateDirectory(dataDirPath);
            }
            else
            {
                System.IO.DirectoryInfo di = new DirectoryInfo(dataDirPath);
                foreach (FileInfo file in di.GetFiles())
                {
                    try
                    {
                        file.Delete();
                    }
                    catch (System.Exception e1)
                    {
                        MessageBox.Show(e1.Message);
                    }
                }
                di = null;
            }
            InitializeComponent();
            Login login = new Login();
            login.ShowDialog();
            if (login.ok)
            {
                this.token = login.token;
            }
            if (!login.ok)
            {
                this.Close();
            }
        }
        private async void Email_Drop(object sender, DragEventArgs e)
        {
            StackPanel stackPanel = sender as StackPanel;
            if (stackPanel != null)
            {
                if (e.Data.GetDataPresent("FileGroupDescriptor"))
                {
                    //this.StackPanelEmails.AllowDrop = false;
                    Outlook.Application outlookApp = new Outlook.Application();
                    Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
                    Outlook.Selection selection = outlookApp.ActiveExplorer().Selection;
                    Outlook.MailItem OutlookMailItem = findOutlookEmailInSelection(selection);
                    if (OutlookMailItem != null)
                    {
                        string subject = OutlookMailItem.Subject;
                        string body = OutlookMailItem.Body.Replace('\r', ' ').Replace('\n', ' ');
                        string klicNummer = getKlicMeldingNumerFromSubject(subject);
                        if (klicNummer.Length > 0)
                        {
                            string fileName = klicNummer;
                            //string compressedFile = ConvertOutlookItemToZip(OutlookMailItem, dataDirPath, klicNummer);
                            YesNoWin msgbox = new YesNoWin(KLICMelding: klicNummer, email: subject, FileName: fileName);
                            if ((bool)msgbox.ShowDialog())
                            {
                                fileName = msgbox.fileName;
                                KLICApi api = new KLICApi();
                                //string url = "https://services-eu1.arcgis.com/fpQ6Cf1ZDEH1n6PO/arcgis/rest/services/Klic_Meldingen_T_V2/FeatureServer/0/query";
                                string url = this.portalUrl + "query";
                                string objectId = await api.getObjectID(token: token, url: url, KlicNummer: klicNummer);
                                if (objectId == null) this.Close();
                                if (objectId.Length != 0)
                                {
                                    string compressedFile = ConvertOutlookItemToZip(OutlookMailItem, dataDirPath, fileName);
                                    //url = "https://services-eu1.arcgis.com/fpQ6Cf1ZDEH1n6PO/arcgis/rest/services/Klic_Meldingen_T_V2/FeatureServer/0/" + objectId + "/addAttachment";
                                    url = this.portalUrl + objectId + "/addAttachment";
                                    if (compressedFile.Length != 0)
                                    {
                                        string filePath = @"" + compressedFile;
                                        var result = await api.addAttachment(token: token, url: url, filePath: filePath);
                                        MessageBox.Show("Bijlage is toegevoegd");
                                        api = null;
                                        try
                                        {
                                            File.Delete(compressedFile);
                                            //this.StackPanelEmails.AllowDrop = true;
                                        }
                                        catch (System.Exception exception)
                                        {
                                            showError("Zip bestand wordt niet verwijderd");
                                        }
                                        //this.StackPanelEmails.AllowDrop = true;
                                    }
                                    else
                                    {
                                        showError("Error tijdens het genereren van het Zip bestand");
                                    }
                                }
                                else
                                {
                                    showError("De Klic meldnummer is niet gevonden!");
                                }
                            }
                            else
                            {
                                showError("Controleer of email onderwerp bevat een Klic meldnummer");
                            }
                        }
                        else
                        {
                            showError("Het programma kan het Klic meldnummer niet ophalen uit het email onderwerp!");
                        }
                    }
                    else
                    {
                        this.showError("Selecteer een email");
                    }
                }
                else
                {
                    this.showError("Onbekend bestand");
                }
            }
        }
        private void Email_DragEnter(object sender, DragEventArgs e)
        {
            StackPanel stackPanel = sender as StackPanel;
            initGridMessage();
            if (stackPanel != null)
            {
                if (e.Data.GetDataPresent("FileGroupDescriptor"))
                {
                    this.StackPanelEmails.AllowDrop = true;
                    this.StackPanelZip.AllowDrop = true;
                    e.Effects = DragDropEffects.Copy;
                }
                else
                {
                    this.StackPanelEmails.AllowDrop = false;
                    this.StackPanelZip.AllowDrop = true;
                }
            }
        }
        private void Email_DragOver(object sender, DragEventArgs e)
        {
        }
        private void StackPanelEmails_DragLeave(object sender, DragEventArgs e)
        {
            StackPanel stackPanel = sender as StackPanel;
            this.StackPanelEmails.Background = new SolidColorBrush(Color.FromArgb(255, 0, 0, 255));
        }
        public string ConvertOutlookItemToZip(Outlook.MailItem mailItem, string dirPath, string fileName)
        {
            string fileToCompressName = "";
            string outputFile = "";
            try
            {
                //string recDatum = mailItem.ReceivedTime.ToString("yyyy-MM-dd hhmmss");
                fileToCompressName = Path.Combine(dirPath, fileName + @".msg");
                string zipFilePath = Path.ChangeExtension(fileToCompressName, ".zip");
                if (File.Exists(zipFilePath))
                {
                    try
                    {
                        File.Delete(zipFilePath);
                    }
                    catch (System.Exception e)
                    {
                        MessageBox.Show("Eerst het onderste bestand verwijderen en probeer het opnieuw:\n" + zipFilePath);
                        return outputFile;
                    }
                }
                FileInfo fileToCompress = new FileInfo(fileToCompressName);
                mailItem.SaveAs(fileToCompressName, Outlook.OlSaveAsType.olMSG);
                outputFile = Path.ChangeExtension(fileToCompressName, ".zip");
                using (ZipArchive archive = ZipFile.Open(outputFile, ZipArchiveMode.Create))
                {
                    archive.CreateEntryFromFile(fileToCompressName, Path.GetFileName(fileToCompressName));
                }
            }
            finally
            {
                if (File.Exists(fileToCompressName))
                {
                    File.Delete(fileToCompressName);
                }
            }
            return outputFile;
        }
        private string getKlicMeldingNumerFromSubject(string subject)
        {
            try
            {
                List<string> result = new List<string>();
                MatchCollection matches = Regex.Matches(subject, @"\b[\w']*\b");
                var words = from m in matches.Cast<Match>()
                            where !string.IsNullOrEmpty(m.Value)
                            select TrimSuffix(m.Value);
                Regex rg = new Regex(@"\d{2}[a-zA-Z]{1}\d{7}");
                foreach (string word in words.ToArray())
                {
                    if (word.Length == 10 && rg.IsMatch(word))
                    {
                        result.Add(word);
                    }
                }
                if (result.Count == 1) return result[0];
                else return "";
            }
            catch (System.Exception e)
            {
                return "";
            }
        }
        private string TrimSuffix(string word)
        {
            int apostropheLocation = word.IndexOf('\'');
            if (apostropheLocation != -1)
            {
                word = word.Substring(0, apostropheLocation);
            }
            return word;
        }
        private void showError(string msg)
        {
            this.emailGridMessage.Text = msg + "\nProbeer het opnieuw";
            this.emailGridMessage.Background = new SolidColorBrush(Color.FromArgb(255, 255, 211, 211));
            //this.StackPanelEmails.AllowDrop = true;
        }
        private void showErrorUnzip(string msg)
        {
            this.UnzipGridMessage.Text = msg + "\nProbeer het opnieuw";
            this.UnzipGridMessage.Background = new SolidColorBrush(Color.FromArgb(255, 255, 211, 211));
            //this.StackPanelZip.AllowDrop = true;
        }
        private Outlook.MailItem findOutlookEmailInSelection(Outlook.Selection selection)
        {
            bool oneOutlookItemIsSelected = false;
            Outlook.MailItem OutlookMailItem = null;
            foreach (object selectedItem in selection)
            {
                if (selectedItem is Outlook.MailItem mailItem)
                {
                    if (oneOutlookItemIsSelected)
                    {
                        OutlookMailItem = null;
                        oneOutlookItemIsSelected = false;
                        break;
                    }
                    oneOutlookItemIsSelected = true;
                    OutlookMailItem = mailItem;
                }
            }
            if (oneOutlookItemIsSelected) return OutlookMailItem;
            else return null;
        }
        private bool IsFileInUse(FileInfo file)
        {
            try
            {
                file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException e) when ((e.HResult & 0x0000FFFF) == 32)
            {
                return true;
            }
            return false;
        }
        private void Zip_DragEnter(object sender, DragEventArgs e)
        {
            StackPanel stackPanel = sender as StackPanel;
            initGridMessage();
            if (stackPanel != null)
            {
                if (e.Data.GetDataPresent(DataFormats.Html))
                {
                    this.StackPanelEmails.AllowDrop = true;
                    this.StackPanelZip.AllowDrop = true;
                    e.Effects = DragDropEffects.Copy;
                }
                else
                {
                    this.StackPanelEmails.AllowDrop = true;
                    this.StackPanelZip.AllowDrop = false;
                    e.Effects = DragDropEffects.None;
                }
            }
        }
        private void initGridMessage()
        {
            this.emailGridMessage.Text = "";
            this.UnzipGridMessage.Text = "";
            this.emailGridMessage.Background = new SolidColorBrush(Colors.LightCyan);
            this.UnzipGridMessage.Background = new SolidColorBrush(Colors.LightCyan);
        }
        private void Zip_DragOver(object sender, DragEventArgs e)
        {
        }
        private async void Zip_DropAsync(object sender, DragEventArgs e)
        {
            StackPanel stackPanel = sender as StackPanel;
            if (stackPanel != null)
            {
                if (e.Data.GetDataPresent(DataFormats.Html))
                {
                    var input = e.Data.GetData(DataFormats.Html);
                    string url = getUrlAttachemtentFromHtmlElement((string)input);
                    KLICApi api = new KLICApi();
                    string fileFullName = Path.Combine(dataDirPath, "downlodedZipFile.zip");
                    string zipFilePath = await api.DownloadFileFromUrl(fileFullName, url, token);
                    if (zipFilePath == null)
                    {
                        showErrorUnzip("De token is niet langer geldig. Probeer opnieuw in te loggen.");
                        this.Close();
                    }
                    if (zipFilePath.Length > 4)
                    {
                        string outlookemailPath = unZIpFile(zipFilePath, dataDirPath);
                        if (outlookemailPath.Length > 4)
                        {
                            try
                            {
                                using (Process fileopener = new Process())
                                {
                                    fileopener.StartInfo.FileName = "explorer";
                                    fileopener.StartInfo.Arguments = "\"" + outlookemailPath + "\"";
                                    fileopener.Start();
                                }
                            }
                            catch (System.Exception exc)
                            {
                                showErrorUnzip(exc.Message);
                            }
                        }
                        else
                        {
                            showErrorUnzip("Programma kan het Zip bestand niet uitpakken");
                        }
                    }
                    else
                    {
                        showErrorUnzip("Programma kan het bestand niet downloaden");
                    }
                }
                else
                {
                    e.Effects = DragDropEffects.None;
                }
            }
        }
        private void tryDeleteFile(string path)
        {
            if (File.Exists(path))
            {
                try
                {
                    File.Delete(path);
                }
                catch (Exception e)
                {

                }
            }
        }
        private string unZIpFile(string filePath, string outputDir)
        {
            try
            {
                string outlookItemFullName = "";
                using (ZipArchive archive = ZipFile.OpenRead(filePath))
                {
                    if (archive.Entries.Count == 1)
                    {
                        outlookItemFullName = archive.Entries.FirstOrDefault().FullName;
                        string outlookItemPath = Path.Combine(outputDir, outlookItemFullName);
                        tryDeleteFile(outlookItemPath);
                        if (!File.Exists(outlookItemPath))
                        {
                            ZipFile.ExtractToDirectory(filePath, outputDir);
                            return outlookItemPath;
                        }
                        else
                        {
                            showErrorUnzip("Programma kan het bestand niet aanmaken. Het bestand is al geopend");
                        }
                    }
                    else
                    {
                        showErrorUnzip("Ongeldig input!");
                        return "";
                    }
                }
            }
            catch (System.Exception e)
            {
                showErrorUnzip(e.Message);
            }
            finally
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            showErrorUnzip("Ongeldig input");
            return "";
        }
        private string getUrlAttachemtentFromHtmlElement(string input)
        {
            var regex = new Regex("<a [^>]*href=(?:'(?<href>.*?)')|(?:\"(?<href>.*?)\")", RegexOptions.IgnoreCase);
            var urls = regex.Matches((string)input).OfType<Match>().Select(m => m.Groups["href"].Value);
            foreach (string url in urls)
            {
                if (url.Contains(portalUrl) && url.Contains("/attachments/") && url.Contains("token"))
                {
                    if (url.Split('?').Length == 2)
                    {
                        return url.Split('?')[0];
                    }
                }
            }
            showErrorUnzip("Het programma kan geen url ophalen voor dit bestand");
            return "";
        }
        private void T_DragLeave(object sender, DragEventArgs e)
        {
        }

    }
}
