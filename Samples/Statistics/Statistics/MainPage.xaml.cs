using Syncfusion.UI.Xaml.RichTextBoxAdv;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Media.Imaging;
using Windows.UI.Xaml.Navigation;
using SelectionChangedEventArgs = Syncfusion.UI.Xaml.RichTextBoxAdv.SelectionChangedEventArgs;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Statistics
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
            richTextBoxAdv.SelectionChanged += RichTextBoxAdv_SelectionChanged;
            Stream inputStream = GetManifestResourceStream("Statistics.docx");
            richTextBoxAdv.Load(inputStream, Syncfusion.UI.Xaml.RichTextBoxAdv.FormatType.Docx);
        }
        private Stream GetManifestResourceStream(string fileName)
        {
            Assembly execAssm = this.GetType().GetTypeInfo().Assembly;
            string[] resourceNames = execAssm.GetManifestResourceNames();
            foreach (string resourceName in resourceNames)
            {
                if (resourceName.EndsWith("." + fileName))
                {
                    fileName = resourceName;
                    break;
                }
            }
            return execAssm.GetManifestResourceStream(fileName);
        }
        private void RichTextBoxAdv_SelectionChanged(object obj, SelectionChangedEventArgs args)
        {

            int wordCount = richTextBoxAdv.WordCount;
            WordCount.Text = wordCount.ToString();
            int paragraphCount = richTextBoxAdv.ParagraphCount;
            ParagraphCount.Text = paragraphCount.ToString();
            int pageCount = richTextBoxAdv.PageCount;
            PageCount.Text = pageCount.ToString();
            int currentPageNumber = richTextBoxAdv.CurrentPageNumber;
            CurrentPageNumber.Text = currentPageNumber.ToString();
        }

    }
}
