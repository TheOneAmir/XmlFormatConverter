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

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void ConvertXml(object sender, RoutedEventArgs e)
        {
            var xmlString = System.IO.File.ReadAllText("this.xml");

            if (Word.IsChecked == true)
            {
                //var xmlString = System.IO.File.ReadAllText("this.xml");
                //doc.LoadFromFile("this.xml", Spire.Doc.FileFormat.Txt);
                //doc.LoadText("this.xml", Encoding.Default);
                //Word.Application app = new Word.Application();
                //Word.Document wDoc = app.Documents.Add();


                //Word.Range rng = app.ActiveDocument.Range(start, end);
                //rng.Text = xmlString;

                //wDoc.SaveAs("wOutput.doc");
                Word.Application app = new Word.Application();
                Word.Document doc = app.Documents.Add();
                try
                {
                    doc.Content.Font.Size = 12;
                    doc.Content.Text = xmlString;

                    doc.Save();
                    doc.Close();

                    //app.Visible = true;    //Optional
                    // MessageBox.Show("File Saved successfully");
                    this.Close();

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
            else if (PDF.IsChecked == true)
            {

            }
        }
    }
}
