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
using System.Windows.Navigation;
using System.Windows.Shapes;
using OfficeMasterKey;

namespace omkwpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_DragEnter(object sender, DragEventArgs e)
        {

        }

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.None;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy | DragDropEffects.Move;
            }
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                this.Activate();

                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                UnprotectFiles(files);
            }
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {

            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog
            {
                Multiselect = true,
                DefaultExt = ".xlsx",
                Filter = "Excel and Word Files (*.xslx,*.docx)|*.xlsx;*.docx|Excel Files (*.xslx)|*.xlsx|Word Files (*.docx)|*.docx|*.*|*.*",
                CheckFileExists = true
            };

            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                try
                {
                    UnprotectFiles(dlg.FileNames);
                }
                catch (Exception x)
                {
                    MessageBox.Show(x.Message, "Error Processing Files", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void UnprotectFiles(string[] files)
        {
            foreach (string file in files)
            {
                UnprotectFile(file);
            }
        }

        private void UnprotectFile(string file)
        {
            MasterKey masterKey = new MasterKey();

            try
            {
                StatusText.Text += "Removing protection from: " + file;

                if (masterKey.FileIsXlsx(file))
                {
                    if (masterKey.XlsxIsProtected(file))
                    {
                        string protectedParts = "";

                        if (masterKey.XlsxIsWorkbookProtected(file) && masterKey.XlsxIsWorksheetProtected(file))
                        {
                            protectedParts = "[Workbook,Worksheet]";
                        }
                        else if (masterKey.XlsxIsWorkbookProtected(file))
                        {
                            protectedParts = "[Workbook]";
                        }
                        else if (masterKey.XlsxIsWorksheetProtected(file))
                        {
                            protectedParts = "[Worksheet]";
                        }

                        masterKey.UnprotectXlsx(file);
                        StatusText.Text += "  OK. " + protectedParts + Environment.NewLine;
                    }
                    else
                    {
                        StatusText.Text += "  Not protected." + Environment.NewLine;
                    }
                }
                else if (masterKey.FileIsDocx(file))
                {
                    if (masterKey.DocxIsProtected(file))
                    {
                        masterKey.UnprotectDocx(file);
                        StatusText.Text += "  OK. " + Environment.NewLine;
                    }
                    else
                    {
                        StatusText.Text += "  Not protected." + Environment.NewLine;
                    }
                }
                else
                {
                    StatusText.Text += "  Not recognized as valid DOCX or XLSX file type." + Environment.NewLine;
                }

            }
            catch (Exception x)
            {
                StatusText.Text += x.Message + Environment.NewLine;
            }
        }

    }
}
