using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using iwantedue;

namespace IdleEmailsProcessor
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

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            string msgFolder = @"C:\workspace\MPRS_reports\KY_ilde_cases\Mar 2015";
            
            DirectoryInfo msgDir = new DirectoryInfo(msgFolder);
            FileInfo[] msgFiles = msgDir.GetFiles("*.msg", SearchOption.TopDirectoryOnly);

            List<string> rows2write = new List<string>();
            foreach (FileInfo fi in msgFiles)
            {
                readMsg(fi.FullName, rows2write);
            }

            //dump to .csv file
            string logfile = msgFolder + @"\Mar.csv";
            foreach (string line in rows2write)
            {
                using (StreamWriter sw = new StreamWriter(logfile, true))
                {
                    sw.WriteLine(line);
                }
            }
        }

        bool readMsg(string msgfile, List<string> rows2write) 
        {
            bool success = true;

            Stream messageStream = File.Open(msgfile, FileMode.Open, FileAccess.Read);
            OutlookStorage.Message message = new OutlookStorage.Message(messageStream);
            string text = message.BodyText;
            //original OutlookStorage needs to be hacked to include email sent date
            //http://www.codeproject.com/Articles/32899/Reading-an-Outlook-MSG-File-in-C?msg=2932319#xx2932319xx
            //look for 'schao' signature in OutlookStorage.cs
            string dateSent = message.ReceivedOrSentTime.ToString();    
            
            int startInd = text.IndexOf("Tester");
            int endInd = text.IndexOf("Please feed");
            text = text.Substring(startInd, (endInd - startInd));

            char[] charsToTrim = { '\t', '\n', '\r'};
            text = text.Trim(charsToTrim);

            //remove html table header
            int tableHeader = text.IndexOf('\n');
            text = text.Substring(tableHeader);

            //remove some weird special between h and mins (\r\t)
            text = text.Trim(charsToTrim);
            text = text.Replace("\r\n", "");

            //there may be multiple idle testers msg inside 1 single email
            //look for *mins table row ending
            while (text.IndexOf("mins") != -1) {
                string row = text.Substring(0, text.IndexOf("mins")+4);
                char[] sep = { ',','\t' };
                row = row.Trim(sep);

                //break $row into '\t' separated words and form .csv row
                string line = "";
                string[] words = row.Split('\t');
                foreach (string word in words)
                {
                    line += word;
                    line += ',';
                }
                line += dateSent;
                rows2write.Add(line);

                //shift to next line
                text = text.Substring(text.IndexOf("mins") + 4);
            }

            messageStream.Close();
            return success;
        }
    }
}