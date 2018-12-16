using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using Word = Microsoft.Office.Interop.Word;
using Spire.Doc;
using iTextSharp.text.pdf;
using iTextSharp.text;

namespace XmlConverter
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

        private async void ConvertXml(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(FilePath.Text))
            {
                MessageBox.Show("Unable to locate file");
            }

            Status.Text = "Reading Input File";
            var filePath = FilePath.Text;

            if (Word.IsChecked == true)
            {
                try
                {
                    ConvertButton.IsEnabled = false;
                    await Task.Run(() => CreateWord(filePath));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Something went wrong during conversion");
                }
            }
            else if (PDF.IsChecked == true)
            {
                try
                {
                    ConvertButton.IsEnabled = false;
                    await Task.Run(() => CreatePDF(filePath));
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Something went wrong during conversion");
                }
            }

            Status.Text = "Finished Conversion";
            ConvertButton.IsEnabled = true;
        }

        private async Task CreateWord(string filePath)
        {
            var xmlString = System.IO.File.ReadAllText(filePath);

            Word.Application app = new Word.Application();
            Word.Document doc = app.Documents.Add();
            try
            {
                doc.Content.Font.Size = 12;
                try
                {
                    doc.Content.Text = xmlString;

                    doc.Save();
                    doc.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to convert file");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(app);
            }
        }
        
        private async Task CreatePDF(string filePath)
        {
            var doc = new iTextSharp.text.Document();

            try
            {
                PdfWriter.GetInstance(doc, new FileStream("Output.pdf", FileMode.Create));
                doc.Open();
                //var xmlString = await Task.Run(() => System.IO.File.ReadAllText(filePath));

                using (StreamReader rdr = new StreamReader(filePath))
                {
                    var currentLine = rdr.ReadLine();
                    //Add the content of Text File to PDF File
                    while (currentLine != null)
                    {
                        doc.Add(new iTextSharp.text.Paragraph(currentLine));
                        currentLine = rdr.ReadLine();
                    }
                }
            }
            catch(Exception e)
            {

            }
            finally
            {
                //Close the Document
                doc.Close();
            }
        }

        private void SelectFile(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xml";
            dlg.Filter = "XML Files (*.xml)|*.xml|TXT Files (*.txt)|*.txt";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                FilePath.Text = filename;
            }
        }
    }
}
