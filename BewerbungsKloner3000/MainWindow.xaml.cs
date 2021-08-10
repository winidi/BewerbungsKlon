using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
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

namespace BewerbungsKloner3000
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string originalPath;
        string anschrift;
        string stellenbezeichnung;
        string anrede;
        string name;

        public MainWindow()
        {
            InitializeComponent();

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            originalPath = txt_oriPath.Text;
            anschrift = txt_Anschrift.Text;
            stellenbezeichnung = txt_Stellenbezeichnung.Text;

            string[] anschriftarray = anschrift.Split(
                new[] { "\r\n", "\r", "\n" },
                StringSplitOptions.None
                );

            name = txt_anredeName.Text;

            string[] namearray = name.Split(
                new[] { "\r\n", "\r", "\n" },
                StringSplitOptions.None
                );

            if (rdb_AllgemeineAnrede.IsChecked == true)
            {
                anrede = "Sehr geehrte Damen und Herren,";
            }
            else if (rdb_AnredeFrau.IsChecked == true)
            {
                anrede = "Sehr geehrte Frau " + name + ",";
            }
            else if (rdb_AnredeMann.IsChecked == true)
            {
                anrede = "Sehr geehrter Herr " + name + ",";
            }

            DateTime dt = DateTime.Today;
            string dateString = dt.ToString("d");

            string[] date = dateString.Split('.');

            string day = date[0].Trim();
            if (day[0].ToString().Equals("0")) day = day[1].ToString();
            string month = date[1].Trim();
            if (month.Equals("01")) month = "Januar";
            else if (month.Equals("02")) month = "Februar";
            else if (month.Equals("03")) month = "März";
            else if (month.Equals("04")) month = "April";
            else if (month.Equals("05")) month = "Mai";
            else if (month.Equals("06")) month = "Juni";
            else if (month.Equals("07")) month = "Juli";
            else if (month.Equals("08")) month = "August";
            else if (month.Equals("09")) month = "September";
            else if (month.Equals("10")) month = "Oktober";
            else if (month.Equals("11")) month = "November";
            else if (month.Equals("12")) month = "Dezember";
            string year = date[2];

            string datecomplete = txt_Ort.Text + ", " + day + ". " + month + " " + year;

            //WoRD
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();

            object missing = System.Type.Missing;

            try
            {
                object fileName = originalPath;
                doc = word.Documents.Open(ref fileName,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing);

                doc.Activate();

                foreach (Microsoft.Office.Interop.Word.Range tmpRange in doc.StoryRanges)
                {
                    tmpRange.Find.Text = "firmenName";
                    tmpRange.Find.Replacement.Text = anschriftarray[0];
                    Console.WriteLine(anschriftarray[0]);
                    tmpRange.Find.Replacement.ParagraphFormat.Alignment =
                        Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;


                    tmpRange.Find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
                    object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;

                    tmpRange.Find.Execute(ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref replaceAll,
                        ref missing, ref missing, ref missing, ref missing);
                }
                foreach (Microsoft.Office.Interop.Word.Range tmpRange in doc.StoryRanges)
                {
                    tmpRange.Find.Text = "firmenStraße";
                    tmpRange.Find.Replacement.Text = anschriftarray[1];
                    Console.WriteLine(anschriftarray[1]);
                    tmpRange.Find.Replacement.ParagraphFormat.Alignment =
                        Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;


                    tmpRange.Find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
                    object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;

                    tmpRange.Find.Execute(ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref replaceAll,
                        ref missing, ref missing, ref missing, ref missing);
                }
                foreach (Microsoft.Office.Interop.Word.Range tmpRange in doc.StoryRanges)
                {
                    tmpRange.Find.Text = "firmenPLZ";
                    tmpRange.Find.Replacement.Text = anschriftarray[2];
                    Console.WriteLine(anschriftarray[2]);
                    tmpRange.Find.Replacement.ParagraphFormat.Alignment =
                        Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;


                    tmpRange.Find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
                    object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;

                    tmpRange.Find.Execute(ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref replaceAll,
                        ref missing, ref missing, ref missing, ref missing);
                }

                foreach (Microsoft.Office.Interop.Word.Range tmpRange in doc.StoryRanges)
                {
                    tmpRange.Find.Text = "STELLENBEZEICHNUNG";
                    tmpRange.Find.Replacement.Text = stellenbezeichnung;

                    tmpRange.Find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
                    object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;

                    tmpRange.Find.Execute(ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref replaceAll,
                        ref missing, ref missing, ref missing, ref missing);
                }

                foreach (Microsoft.Office.Interop.Word.Range tmpRange in doc.StoryRanges)
                {
                    tmpRange.Find.Text = "Sehr geehrte Damen und Herren,";
                    tmpRange.Find.Replacement.Text = anrede;

                    tmpRange.Find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
                    object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;

                    tmpRange.Find.Execute(ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref replaceAll,
                        ref missing, ref missing, ref missing, ref missing);
                }
                foreach (Microsoft.Office.Interop.Word.Range tmpRange in doc.StoryRanges)
                {
                    tmpRange.Find.Text = "datum";
                    tmpRange.Find.Replacement.Text = datecomplete;

                    tmpRange.Find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
                    object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;

                    tmpRange.Find.Execute(ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref replaceAll,
                        ref missing, ref missing, ref missing, ref missing);
                }

                // Create Directory
                string firmennameString = anschriftarray[0].ToString();
                string directoryName = new DirectoryInfo(System.IO.Path.GetDirectoryName(originalPath)).FullName;
                string directoryNameNew = $@"{directoryName}\{firmennameString}";

                try
                {
                    if (Directory.Exists(directoryNameNew))
                    {
                        Console.WriteLine("Path exists already");
                        return;
                    }
                    DirectoryInfo di = Directory.CreateDirectory(directoryNameNew);
                    Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(directoryNameNew));

                }
                catch (Exception)
                {
                    Console.WriteLine("The process failed: {0}", e.ToString());
                }
                finally { }

                var pathNamePdf = $@"{directoryNameNew}\Anschreiben_{firmennameString}.pdf";
                var pathNameWord = $@"{directoryNameNew}\Anschreiben_{firmennameString}.docx";


                string fileDirectionPdf = pathNamePdf;
                string fileDirectionWord = pathNameWord;

                doc.SaveAs2(fileDirectionWord);
                doc.ExportAsFixedFormat(fileDirectionPdf, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);

                doc.Close(ref missing, ref missing, ref missing);
                word.Application.Quit(ref missing, ref missing, ref missing);

                // Copy Lebenslauf und Anhänge

                    string lebenslaufOri = txt_lebenslauf.Text;
                    string anhang1Ori = txt_CopyFile1.Text;
                    string anhang2Ori = txt_CopyFile2.Text;
                    string anhang3Ori = txt_CopyFile3.Text;
                    string anhang4Ori = txt_CopyFile4.Text;

                Console.WriteLine(lebenslaufOri);

                try
                {
                    if (lebenslaufOri.Length > 1)
                    {
                        string lebensLaufFN = System.IO.Path.GetFileName(lebenslaufOri);
                        Console.WriteLine("LebenlaufFilename: ");
                        Console.WriteLine(lebensLaufFN);
                        string lebenslaufDes = $@"{directoryNameNew}\{lebensLaufFN}";
                        Console.WriteLine("LebenlaufDes: ");
                        Console.WriteLine(lebenslaufDes);
                        File.Copy(lebenslaufOri, lebenslaufDes);
                    }
                    if (anhang1Ori.Length > 1)
                    {
                        string anhang1FN = System.IO.Path.GetFileName(anhang1Ori);
                        string anhang1Des = $@"{directoryNameNew}\{anhang1FN}";
                        File.Copy(anhang1Ori, anhang1Des);
                    }
                    if (anhang2Ori.Length > 1)
                    {
                        string anhang2FN = System.IO.Path.GetFileName(anhang2Ori);
                        string anhang2Des = $@"{directoryNameNew}\{anhang2FN}";
                        File.Copy(anhang1Ori, anhang2Des);
                    }
                    if (anhang3Ori.Length > 1)
                    {
                        string anhang3FN = System.IO.Path.GetFileName(anhang3Ori);
                        string anhang3Des = $@"{directoryNameNew}\{anhang3FN}";
                        File.Copy(anhang3Ori, anhang3Des);
                    }
                    if (anhang4Ori.Length > 1)
                    {
                        string anhang4FN = System.IO.Path.GetFileName(anhang4Ori);
                        string anhang4Des = $@"{directoryNameNew}\{anhang4FN}";
                        File.Copy(anhang4Ori, anhang4Des);
                    }
                }
                catch
                {
                    Console.WriteLine("failure in copying File");
                }
                // clear input fields
                txt_Anschrift.Text = "";
                txt_Stellenbezeichnung.Text = "";
                txt_anredeName.Text = "";
            }
            catch (Exception)
            {
                doc.Close(ref missing, ref missing, ref missing);
                word.Application.Quit(ref missing, ref missing, ref missing);

                MessageBox.Show("Problem");
            }
        }
    }
}