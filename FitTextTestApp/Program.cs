namespace FitTextTestApp
{
    using DocumentFormat.OpenXml;
    #region Using

    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using GeneratedCode;

    #endregion

    class Program
    {
        static void Main(string[] args)
        {
            //new GeneratedClass().CreatePackage("e:\\lihailewodege.docx");

            using (WordprocessingDocument word = WordprocessingDocument.Open(@"C:\Users\Administrator\Desktop\1.docx", true))
            {
                /* paragraph list */
                OpenXmlElementList list = word.MainDocumentPart.Document.Body.ChildElements;

                /* paragraph */
                Paragraph paragraph = (Paragraph)list[23];

                /* run */
                OpenXmlElementList plist = paragraph.ChildElements;
                string localname = plist[6].LocalName;
                Run run = (Run)plist[6];

                /* font size */
                StringValue fontsize_str = run.RunProperties.FontSize.Val;
                uint fontsize_temp = 0;
                uint.TryParse(fontsize_str, out fontsize_temp);

                /* 固定长度 */
                uint FIT_WIDTH_VAL = 311;

                /* fittext width */
                UInt32Value fitval_5 = UInt32Value.FromUInt32(FIT_WIDTH_VAL * 5);
                UInt32Value fitval_3 = UInt32Value.FromUInt32(FIT_WIDTH_VAL * 3);

                /* inner text */
                string innertext = run.InnerText;
                string[] innertexts = innertext.Split(new[] { '　' }, System.StringSplitOptions.RemoveEmptyEntries);
                string first_child = innertexts[0];
                string second_child = innertexts[1];

                /* remove run */
                paragraph.RemoveChild<Run>(run);

                /* create new run 1 */
                Run run1 = new Run();
                RunProperties runProperties1 = new RunProperties();
                FitText fitText1 = new FitText() { Val = fitval_5 };
                RunFonts runfont1 = new RunFonts() { Ascii = "仿宋", HighAnsi = "仿宋", EastAsia = "仿宋", ComplexScript = "仿宋" };
                FontSize fontsize1 = new FontSize() { Val = "32" };

                Text text1 = new Text();
                text1.Text = first_child;

                runProperties1.Append(fitText1);
                runProperties1.Append(runfont1);
                runProperties1.Append(fontsize1);

                run1.Append(runProperties1);
                run1.Append(text1);

                /* create new run 2 */
                Run run2 = new Run();
                RunProperties runProperties2 = new RunProperties();
                RunFonts runfont2 = new RunFonts() { Ascii = "仿宋", HighAnsi = "仿宋", EastAsia = "仿宋", ComplexScript = "仿宋" };
                FontSize fontsize2 = new FontSize() { Val = "32" };

                Text text2 = new Text();
                text2.Text = "　";

                runProperties2.Append(runfont2);
                runProperties2.Append(fontsize2);

                run2.Append(runProperties2);
                run2.Append(text2);

                /* create new run 3 */
                Run run3 = new Run();
                RunProperties runProperties3 = new RunProperties();
                FitText fitText3 = new FitText() { Val = fitval_3 };
                RunFonts runfont3 = new RunFonts() { Ascii = "仿宋", HighAnsi = "仿宋", EastAsia = "仿宋", ComplexScript = "仿宋" };
                FontSize fontsize3 = new FontSize() { Val = "32" };

                Text text3 = new Text();
                text3.Text = second_child;

                runProperties3.Append(fitText3);
                runProperties3.Append(runfont3);
                runProperties3.Append(fontsize3);

                run3.Append(runProperties3);
                run3.Append(text3);

                paragraph.Append(run1);
                paragraph.Append(run2);
                paragraph.Append(run3);
            }

        }
    }
}