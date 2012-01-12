using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Win32;

namespace converter
{

    public delegate void ConversionCompletedCallback();

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Thread convertThread;
        private Converter converter;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void convertClick(object sender, RoutedEventArgs e)
        {
            //Kick off a file conversion
            
            //First prompt the user for all of the files to convert
            var fileDialog = new Microsoft.Win32.OpenFileDialog();
            fileDialog.FileName = "";
            fileDialog.Multiselect = true;
            fileDialog.DefaultExt = ".rtf";
            fileDialog.Filter = "Exported QG Files (.rtf)|*.rtf";
            Nullable<bool> result = fileDialog.ShowDialog();

            if (result == true)
            {
                //Now prompt the user for the directory to output the files to
                System.Windows.Forms.FolderBrowserDialog folderDialog = new System.Windows.Forms.FolderBrowserDialog();
                folderDialog.Description = "Choose the output directory.";
                System.Windows.Forms.DialogResult folderResult = folderDialog.ShowDialog(this.GetIWin32Window());
                if (folderResult == System.Windows.Forms.DialogResult.OK)
                {
                    //We are now done with user input - lets get to work converting the files
                    convertButton.Visibility = System.Windows.Visibility.Collapsed;
                    progressGrid.Visibility = System.Windows.Visibility.Visible;

                    converter = new Converter(fileDialog.FileNames,
                        folderDialog.SelectedPath,
                        new ConversionCompletedCallback(conversionCompleted),
                        Dispatcher);
                    convertThread = new Thread(new ThreadStart(converter.perform));
                    convertThread.Start();
                }
            }
        }

        private void conversionCompleted()
        {
            convertButton.Visibility = System.Windows.Visibility.Visible;
            progressGrid.Visibility = System.Windows.Visibility.Collapsed;
        }

    }
}
