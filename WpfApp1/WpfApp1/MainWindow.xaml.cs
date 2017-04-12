using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Windows;
using DocumentFormat.OpenXml;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string PowerPointFileName { get; set; }
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog()
            {
                FileName = "AnyPowerPointFile",
                DefaultExt = ".pptx",
                Filter = "PowerPoint Document (.pptx)|*.pptx| (.ppsx) | *.ppsx"
            };

            // Show save file dialog box
            var result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                //Display durations
                PowerPointFileName = dlg.FileName;
                GetSlideDurations(PowerPointFileName);
            }
        }

        private void GetSlideDurations(string powerPointFileName)
        {
            try
            {
                textBox.Text = "";
                TimeSpan presentationDuration = TimeSpan.FromSeconds(0);
                using (PresentationDocument pptDocument = PresentationDocument.Open(powerPointFileName, false))
                {
                    PresentationPart presentationPart = pptDocument.PresentationPart;
                    Presentation presentation = presentationPart.Presentation;
                    int slideNumber = 1;
                    foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;

                        var advanceAfterTimeDuration = ConvertStringToInt(GetSlideAdvanceAfterTimeDuration(slidePart));
                        var anitationsDuration = GetSlideAnimationsDuration(slidePart);
                        var transitionDuration = ConvertStringToInt(GetSlideTransitionsDuration(slidePart));

                        var totalSlideDuration = advanceAfterTimeDuration + anitationsDuration + transitionDuration;
                        TimeSpan slideTime = TimeSpan.FromMilliseconds(totalSlideDuration);
                        presentationDuration = presentationDuration.Add(slideTime);

                        textBox.Text += $"Slide {slideNumber} Total Duration: {totalSlideDuration} ms. (aat: {advanceAfterTimeDuration} ms, ani: {anitationsDuration} ms trn: {transitionDuration} ms)" + Environment.NewLine;
                        slideNumber++;
                    }

                    textBox.Text += $"Total Presentation Duration: {presentationDuration.TotalMilliseconds} msecs." + Environment.NewLine;

                }
            }
            catch (Exception ex)
            {
                textBox.Text = $"Problem occurred parsing file {powerPointFileName}.  Exception: {ex}";
            }

        }

        private string GetSlideTransitionsDuration(SlidePart slidePart)
        {
            string returnDuration = "0";
            try
            {
                Slide slide1 = slidePart.Slide;

                var transitions = slide1.Descendants<Transition>();
                foreach (var transition in transitions)
                {
                    if (transition.Duration.HasValue)
                        return transition.Duration;
                    break;
                }
            }
            catch (Exception ex)
            {
                //Do nothing
            }

            return returnDuration;
        }

        private int GetSlideAnimationsDuration(SlidePart slidePart)
        {
            int returnDuration = 0;
            try
            {
                Slide slide1 = slidePart.Slide;

                var timeNotes = slide1.Descendants<CommonTimeNode>();
                foreach (var timeNode in timeNotes)
                {
                    if (timeNode.Duration.HasValue)
                    {
                        returnDuration += ConvertStringToInt(timeNode.Duration);
                    }
                    
                }
            }
            catch (Exception ex)
            {
                //Do nothing
            }

            return returnDuration;
        }



        private string GetSlideAdvanceAfterTimeDuration(SlidePart slidePart)
        {
            string returnDuration = "0";
            try
            {
                Slide slide1 = slidePart.Slide;

                var transitions = slide1.Descendants<Transition>();
                foreach (var transition in transitions)
                {
                    if (transition.AdvanceAfterTime.HasValue)
                        return transition.AdvanceAfterTime;
                    break;
                }
            }
            catch (Exception ex)
            {
                //Do nothing
            }

            return returnDuration;

        }

        private int ConvertStringToInt(StringValue stringValue)
        {
            int convertedInt = 0;
            try
            {
                Int32.TryParse(stringValue, out convertedInt);

            }
            catch (Exception)
            {

                throw;
            }
            return convertedInt;
        }
    }
}
