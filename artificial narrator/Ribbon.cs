using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Speech;
using Microsoft.Office.Interop.PowerPoint;
using System.Speech.Synthesis;


namespace artificial_narrator
{
    public partial class Ribbon
    {
        public String SelectedVoice = "";
        public Int32 RateNumber = 0;
        SpeechSynthesizer SpeechSynth = new SpeechSynthesizer();
        Prompt LastTalkPrompt = null;
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            InsertNarration.Label = "このスライドに\r\nナレーションを挿入";
            InsertNarrationAll.Label = "すべてのスライドに\r\nナレーションを挿入";
            var VoiceList = SpeechSynth.GetInstalledVoices().Select(v => v.VoiceInfo.Name);
            var FirstSelectVoice = SpeechSynth.GetInstalledVoices().FirstOrDefault(v => v.VoiceInfo.Culture.Name == System.Globalization.CultureInfo.CurrentCulture.Name).VoiceInfo.Name;
            foreach(var Voice in VoiceList)
            {
                var rddi = Factory.CreateRibbonDropDownItem();
                rddi.Label = Voice;
                VoiceListBox.Items.Add(rddi);
                if (Voice == FirstSelectVoice)
                {
                    VoiceListBox.SelectedItem = rddi;
                    SelectedVoice = Voice;
                }
             }
            foreach(var Rate in Enumerable.Range(-10, 21))
            {
                var rddi = Factory.CreateRibbonDropDownItem();
                rddi.Label = $"{Rate}";
                
                RateBox.Items.Add(rddi);
                if (Rate == 0) RateBox.SelectedItem = rddi;
            }
        }

        private void InsertNarration_Click(object sender, RibbonControlEventArgs e)
        {
            InsertNarration.Enabled =InsertNarration.Enabled= false;
            var Slides = GetSlides();
            if (((RibbonButton) sender).Name == "InsertNarrationAll")
            {
                Slides.Clear();
                var SlideTemp = Globals.ThisAddIn.Application.ActivePresentation.Slides;
                foreach (Slide s in SlideTemp)
                {
                    Slides.Add(s);
                }
            }
            foreach (var Slide in Slides)
            {
                var Shapes = Slide.Shapes;
                var oldText = "";
                Shape ShapeToDelete = null; ;
                foreach (Shape TheShape in Shapes)
                {
                    try
                    {
                        oldText = TheShape.Tags["Narrator"];
                        if (oldText != "") ShapeToDelete = TheShape;
                    }
                    catch (Exception ex)
                    {
                        new StreamWriter(Console.OpenStandardError()).WriteLine(ex.Message);
                    }
                }
                if (ShapeToDelete != null)
                {
                    ShapeToDelete.Delete();
                 }
                string text2;
                text2 = GetText(Slide);
                Console.WriteLine(text2);
                var FileName = Path.GetTempFileName() + ".WAV";
                PlayBack(text2, FileName, SelectedVoice);
                var addition = Shapes.AddMediaObject2(FileName, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 20, 20);
                addition.AnimationSettings.PlaySettings.PlayOnEntry = Microsoft.Office.Core.MsoTriState.msoTrue;
                addition.Tags.Add("Narrator", text2);
                addition.Left = -20;
                addition.AnimationSettings.PlaySettings.PlayOnEntry = Microsoft.Office.Core.MsoTriState.msoTrue;
                // Globals.ThisAddIn.Application.ActivePresentation.EnsureAllMediaUpgraded();
            }
            InsertNarration.Enabled = InsertNarration.Enabled = false;
        }

        private List<Microsoft.Office.Interop.PowerPoint.Slide> GetSlides()
        {
            var App = Globals.ThisAddIn.Application;
            var Pres = App.ActivePresentation;
            var SlidIndexes = App.ActiveWindow.Selection.SlideRange;
            var Slides = new List<Slide>();
            foreach (Slide Sl in SlidIndexes)
                Slides.Add(Sl);
            return Slides;
        }


        private String GetText(Microsoft.Office.Interop.PowerPoint. Slide Slide)
        {
            var ShapeList = Enumerable.Range(0, Slide.Shapes.Count);
            return Slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text;
        }

        private void PlayBack(string text, string FileName = "", String Voice = "")
        {
            //var SpeachSynth = new System.Speech.Synthesis.SpeechSynthesizer();
            GetVoiceParameters();
            SpeechSynth.SelectVoice(Voice);
            SpeechSynth.Rate = RateNumber;
            if (FileName == "")
            {
                SpeechSynth.SpeakAsyncCancelAll();
                while(SpeechSynth.State != SynthesizerState.Ready)
                {
                    System.Threading.Thread.Sleep(10);
                }
                SpeechSynth.SetOutputToDefaultAudioDevice();
                LastTalkPrompt = SpeechSynth.SpeakAsync(text);

            }
            else
            { 
                SpeechSynth.SetOutputToWaveFile(FileName);
                SpeechSynth.Speak(text);
            }

        }

        private void TestSpeech_Click(object sender, RibbonControlEventArgs e)
        {
            var ListOfSlides = GetSlides();
            foreach (var s in ListOfSlides)
            {
                var text = GetText(s);
                PlayBack(text, "", SelectedVoice);
            }
        }

        private void GetVoiceParameters(object sender = null, RibbonControlEventArgs e = null)
        {
            SelectedVoice = VoiceListBox.SelectedItem.Label;
            RateNumber = Convert.ToInt32(RateBox.SelectedItem.Label);
        }
    }
}
