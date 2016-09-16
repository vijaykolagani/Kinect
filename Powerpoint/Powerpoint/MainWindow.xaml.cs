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
using Microsoft.Office.Core;
using Graph = Microsoft.Office.Interop.Graph;
using PowerPoint1 = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
					

namespace PowerPoint
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

       
        PowerPoint1.Application objApp;
        PowerPoint1.Presentations objPresSet;
        PowerPoint1._Presentation objPres = null;
        PowerPoint1.Slides objSlides;
        PowerPoint1._Slide objSlide;
        PowerPoint1.TextRange objTextRng;
        PowerPoint1.Shapes objShapes;
        PowerPoint1.Shape objShape;
        PowerPoint1.SlideShowWindows objSSWs;
        PowerPoint1.SlideShowTransition objSST;
        PowerPoint1.SlideShowSettings objSSS;
        PowerPoint1.SlideRange objSldRng;
        Graph.Chart objChart;
        private void ShowPresentation()
        {
            String strTemplate, strPic;
            strTemplate =
              "C:\\Program Files\\Microsoft Office\\Templates\\1033\\WidescreenPresentation.potx";
            strPic = "C:\\Windows\\minterm.png";
            bool bAssistantOn;

            

         
            //Create a new presentation based on a template.
            objApp = new PowerPoint1.Application();
            objApp.Visible = MsoTriState.msoTrue;
            objPresSet = objApp.Presentations;
            objPres = objPresSet.Open(strTemplate,
                MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
            objSlides = objPres.Slides;

            //Build Slide #1:
            //Add text to the slide, change the font and insert/position a 
            //picture on the first slide.
            objSlide = objSlides.Add(1, PowerPoint1.PpSlideLayout.ppLayoutTitleOnly);
            objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
            objTextRng.Text = "My Sample Presentation";
            objTextRng.Font.Name = "Comic Sans MS";
            objTextRng.Font.Size = 48;
            objSlide.Shapes.AddPicture(strPic, MsoTriState.msoFalse, MsoTriState.msoTrue,
                150, 150, 500, 350);

            //Build Slide #2:
            //Add text to the slide title, format the text. Also add a chart to the
            //slide and change the chart type to a 3D pie chart.
            objSlide = objSlides.Add(2, PowerPoint1.PpSlideLayout.ppLayoutTitleOnly);
            objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
            objTextRng.Text = "My Chart";
            objTextRng.Font.Name = "Comic Sans MS";
            objTextRng.Font.Size = 48;
            objChart = (Graph.Chart)objSlide.Shapes.AddOLEObject(150, 150, 480, 320,
                "MSGraph.Chart.8", "", MsoTriState.msoFalse, "", 0, "",
                MsoTriState.msoFalse).OLEFormat.Object;
            objChart.ChartType = Graph.XlChartType.xl3DPie;
            objChart.Legend.Position = Graph.XlLegendPosition.xlLegendPositionBottom;
            objChart.HasTitle = true;
            objChart.ChartTitle.Text = "Here it is...";

            //Build Slide #3:
            //Change the background color of this slide only. Add a text effect to the slide
            //and apply various color schemes and shadows to the text effect.
            objSlide = objSlides.Add(3, PowerPoint1.PpSlideLayout.ppLayoutBlank);
            objSlide.FollowMasterBackground = MsoTriState.msoFalse;
            objShapes = objSlide.Shapes;
            objShape = objShapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect27,
              "The End", "Impact", 96, MsoTriState.msoFalse, MsoTriState.msoFalse, 230, 200);

            //Modify the slide show transition settings for all 3 slides in
            //the presentation.
            int[] SlideIdx = new int[3];
            for (int i = 0; i < 3; i++) SlideIdx[i] = i + 1;
            objSldRng = objSlides.Range(SlideIdx);
            objSST = objSldRng.SlideShowTransition;
            objSST.AdvanceOnTime = MsoTriState.msoTrue;
            objSST.AdvanceTime = 3;
            objSST.EntryEffect = PowerPoint1.PpEntryEffect.ppEffectBoxOut;

            //Prevent Office Assistant from displaying alert messages:
           // bAssistantOn = objApp.Assistant.On;
            //objApp.Assistant.On = false;

            //Run the Slide show from slides 1 thru 3.
            objSSS = objPres.SlideShowSettings;
            objSSS.StartingSlide = 1;
            objSSS.EndingSlide = 3;
            objSSS.Run();

            //Wait for the slide show to end.
            objSSWs = objApp.SlideShowWindows;
          //  while (objSSWs.Count >= 1) System.Threading.Thread.Sleep(100);

            //Reenable Office Assisant, if it was on:
            //if (bAssistantOn)
            //{
              //  objApp.Assistant.On = true;
                //objApp.Assistant.Visible = false;
            //}

            //Close the presentation without saving changes and quit PowerPoint.
            //objPres.Close();
           // objApp.Quit();
        }

        private void Button_Next(object sender, RoutedEventArgs e)
        {
            if(objSSWs != null && objSSWs.Count>=1)
            objPres.SlideShowWindow.View.Next();

        }
        private void Button_prev(object sender, RoutedEventArgs e)
        {
            if(objSSWs != null && objSSWs.Count>=1)
            objPres.SlideShowWindow.View.Previous();
        }
        private void FileOpen_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();



            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".ppt";
            dlg.Filter = "Prsentation (*.ppt;*.pptx;*.pptm;*.ppsx;*.pps;*.ppsm)|*.ppt;*.pptx;*.pptm;*.ppsx;*.pps;*.ppsm";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                



                //Create a new presentation based on a template.
                objApp = new PowerPoint1.Application();
                objApp.Visible = MsoTriState.msoTrue;
                objPresSet = objApp.Presentations;
                objPres = objPresSet.Open(filename,
                    MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
                objSlides = objPres.Slides;
               
                //Prevent Office Assistant from displaying alert messages:
                // bAssistantOn = objApp.Assistant.On;
                //objApp.Assistant.On = false;

                //Run the Slide show 
                objSSS = objPres.SlideShowSettings;
                objSSS.Run();
                objSSWs = objApp.SlideShowWindows;
            }
        }
    }
}
