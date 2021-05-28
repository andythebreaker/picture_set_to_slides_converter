using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using D = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Windows.Forms;
using Text = DocumentFormat.OpenXml.Drawing.Text;
using System.Drawing.Imaging;
using System.IO;

namespace picture_set_to_slides_converter
{
    public partial class Form1 : Form
    {
        public UInt32Value uuid = 1;
        public double cm()
        {
            return ((9144000.0 / 25.4) + (6858000.0 / 19.5)) / 2.0;
        }

        public Int32 cm32(double cmvalue)
        {
            return (Int32)((cmvalue * ((9144000.0 / 25.4) + (6858000.0 / 19.5))) / 2.0);
        }

        public Int64 cm64(double cmvalue)
        {
            return (Int32)((cmvalue * ((9144000.0 / 25.4) + (6858000.0 / 19.5))) / 2.0);
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IEnumerable<string> alljpgsie = Directory.EnumerateFiles(tb2.Text, "*.jpg");
            label7.Text = alljpgsie.Count().ToString();
            string pathtopptx = tb1.Text + "\\" + DateTime.UtcNow.ToString().Replace(',', '_').Replace(':', '_').Replace('/', '_').Replace(' ', '_').Replace('上', 'a').Replace('下', 'p').Replace('午', 'm') + ".pptx";
            label1.Text = pathtopptx;
            CreatePresentation(pathtopptx, alljpgsie);

        }
        // Insert the specified slide into the presentation at the specified position.
        public void InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle, string alljpgsieitem)
        {

            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            if (slideTitle == null)
            {
                throw new ArgumentNullException("slideTitle");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation is not empty.
            if (presentationPart == null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Declare and instantiate a new slide.
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            uint drawingObjectId = 1;

            // Construct the slide content.            
            // Specify the non-visual properties of the new slide.
            P.NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new P.NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new P.NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new P.NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

            // Specify the group shape properties of the new slide.
            slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

            // Declare and instantiate the title shape of the new slide.
            P.Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());

            drawingObjectId++;

            // Specify the required shape properties for the title shape. 
            titleShape.NonVisualShapeProperties = new P.NonVisualShapeProperties
                (new P.NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
            titleShape.ShapeProperties = new P.ShapeProperties();

            // Specify the text of the title shape.
            titleShape.TextBody = new P.TextBody(new BodyProperties(),
                    new ListStyle(),
                    new Paragraph(new Run(new Text() { Text = slideTitle })));

            // Declare and instantiate the body shape of the new slide.
            P.Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());
            drawingObjectId++;

            //=====================================================dupdupdup
            string jpgpath = alljpgsieitem;
            Image image2 = Image.FromFile(jpgpath);
            double imw = (image2.Width / image2.HorizontalResolution) * 2.54 ;
            double imh = (image2.Height / image2.VerticalResolution) * 2.54 ;
            double imwr  ;
            double imhr  ;
            if ((imw > Decimal.ToDouble(pgx.Value)) && (imh > Decimal.ToDouble(pgy.Value)))
            {
                double aimhr = Decimal.ToDouble(pgy.Value);
                double aX = aimhr / imh;
                double bimwr = Decimal.ToDouble(pgx.Value);
                double bX = bimwr / imw;
                if (aX < bX)
                {
                    imwr = aX * imw;
                    imhr = aX * imh;
                }
                else
                {
                    imwr = bX * imw;
                    imhr = bX * imh;
                }
            }
            else if (imh > Decimal.ToDouble(pgy.Value))
            {
                imhr = Decimal.ToDouble(pgy.Value);
                imwr = (imhr / imh) * imw;
            }
            else if (imw > Decimal.ToDouble(pgx.Value))
            {
                imwr = Decimal.ToDouble(pgx.Value);
                imhr = (imwr / imw) * imh;
            }
            else
            {
                imhr = imh; imwr = imw;
            }
            double imlx = Math.Abs((Decimal.ToDouble(pgx.Value)) / 2 - imwr / 2);
            double imly = Math.Abs((Decimal.ToDouble(pgy.Value)) / 2 - imhr / 2);


            //--------
            P.Picture picShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Picture(new P.NonVisualPictureProperties
                                (
                                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)1026U, Name = "Photo", Description = "" },
                                    new P.NonVisualPictureDrawingProperties
                                    (
                                        new D.PictureLocks() { NoChangeAspect = true }
                                    ),
                                    new ApplicationNonVisualDrawingProperties()
                                ),
                                new P.BlipFill
                                (
                                    new D.Blip
                                        (
                                            new D.NonVisualPicturePropertiesExtensionList()
                                        )
                                    { Embed = "rId2" },
                                    new D.Stretch
                                    (
                                        new D.FillRectangle()
                                    )
                                ),
                                new P.ShapeProperties
                                (
                                    new D.Transform2D
                                    (
                                        new D.Offset() { X = cm64(imlx), Y = cm64(imly) },
                                        new D.Extents()
                                        {
                                            Cx = cm64(imwr),
                                            Cy = cm64(imhr)
                                        }
                                    ),
                                    new D.PresetGeometry
                                    (
                                        new D.AdjustValueList()
                                    )
                                    { Preset = D.ShapeTypeValues.Rectangle }
                                )));
            drawingObjectId++;


            // Specify the required shape properties for the body shape.
            bodyShape.NonVisualShapeProperties = new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
                    new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
            bodyShape.ShapeProperties = new P.ShapeProperties();

            // Specify the text of the body shape.
            bodyShape.TextBody = new P.TextBody(new BodyProperties(),
                    new ListStyle(),
                    new Paragraph());

            // Create the slide part for the new slide.
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            //-*-*-*-*-*-*-*-

            ImagePart ip = slidePart.AddImagePart(ImagePartType.Jpeg, "rId2");
            Stream MyStream = ToStream(image2, ImageFormat.Jpeg);
            MyStream.Position = 0;
            ip.FeedData(MyStream);

            // Save the new slide part.
            slide.Save(slidePart);

            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }

            // Use the same slide layout as that of the previous slide.
            if (null != lastSlidePart.SlideLayoutPart)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();
        }

        public void CreatePresentation(string filepath, IEnumerable<string> alljpgsie)
        {
            // Create a presentation at a specified file path. The presentation document type is pptx, by default.
            PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
            PresentationPart presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            int fei = 0;
            foreach (var alljpgsieitem in alljpgsie)
            {
                if (fei == 0)
                {
                    CreatePresentationParts(presentationPart, alljpgsieitem);
                }
                else
                {
                    InsertNewSlide(presentationDoc, fei, " ", alljpgsieitem);
                }
                fei++;
            }

            // Close the presentation handle
            presentationDoc.Close();
        }
        private void CreatePresentationParts(PresentationPart presentationPart, string impth)
        {
            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(
                new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
            SlideIdList slideIdList1 = new SlideIdList(
                new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
            SlideSize slideSize1 = new SlideSize() { Cx = cm32(Decimal.ToDouble(pgx.Value)), Cy = cm32(Decimal.ToDouble(pgy.Value)), Type = (SlideSizeValues)cb1.SelectedIndex };
            NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

            /*SlidePart slidePart1;
            SlideLayoutPart slideLayoutPart1;
            SlideMasterPart slideMasterPart1;
            ThemePart themePart1;*/

            for (int i = 0; i < 1; i++)
            {
                SlidePart slidePart10;
                SlideLayoutPart slideLayoutPart10;
                SlideMasterPart slideMasterPart10;
                ThemePart themePart10;

                slidePart10 = CreateSlidePart(presentationPart, impth);

                slideLayoutPart10 = CreateSlideLayoutPart(slidePart10);
                slideMasterPart10 = CreateSlideMasterPart(slideLayoutPart10);
                themePart10 = CreateTheme(slideMasterPart10);
                slideMasterPart10.AddPart(slideLayoutPart10, (i == 0) ? "rId1" : "rId4");
                presentationPart.AddPart(slideMasterPart10, (i == 0) ? "rId1" : "rId4");
                if (i == 0)
                {
                    presentationPart.AddPart(themePart10, "rId5");
                }

            }

            /*slidePart1 = CreateSlidePart(presentationPart, "rId2");
            slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
            slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
            themePart1 = CreateTheme(slideMasterPart1);

            slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
            presentationPart.AddPart(slideMasterPart1, "rId1");
            presentationPart.AddPart(themePart1, "rId5");*/
        }
        private SlidePart CreateSlidePart(PresentationPart presentationPart, string impth)
        {
            string jpgpath = impth;
            Image image2 = Image.FromFile(jpgpath);
            double imw = (image2.Width / image2.HorizontalResolution) * 2.54 ;
            double imh = (image2.Height / image2.VerticalResolution) * 2.54 ;
            double imwr  ;
            double imhr  ;
            if ((imw > Decimal.ToDouble(pgx.Value)) && (imh > Decimal.ToDouble(pgy.Value)))
            {
                double aimhr = Decimal.ToDouble(pgy.Value);
                double aX = aimhr / imh;
                double bimwr = Decimal.ToDouble(pgx.Value);
                double bX = bimwr / imw;
                if (aX < bX)
                {
                    imwr = aX * imw;
                    imhr = aX * imh;
                }
                else
                {
                    imwr = bX * imw;
                    imhr = bX * imh;
                }
            }
            else if (imh > Decimal.ToDouble(pgy.Value))
            {
                imhr = Decimal.ToDouble(pgy.Value);
                imwr = (imhr / imh) * imw;
            }
            else if (imw > Decimal.ToDouble(pgx.Value))
            {
                imwr = Decimal.ToDouble(pgx.Value);
                imhr = (imwr / imw) * imh;
            }
            else
            {
                imhr = imh; imwr = imw;
            }
            double imlx = Math.Abs((Decimal.ToDouble(pgx.Value)) / 2 - imwr / 2);
            double imly = Math.Abs((Decimal.ToDouble(pgy.Value)) / 2 - imhr / 2);
            SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");
            slidePart1.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new TransformGroup()),
                            new P.Shape(
                                new P.NonVisualShapeProperties(
                                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                    new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                                new P.ShapeProperties(),
                                new P.TextBody(
                                    new BodyProperties(),
                                    new ListStyle(),
                                    new Paragraph(
                                        new Run(
                                            new Text("")),
                                        new EndParagraphRunProperties() { Language = "en-US" }))),
                            new P.Picture
                            (
                                new P.NonVisualPictureProperties
                                (
                                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)1026U, Name = "Photo", Description = "" },
                                    new P.NonVisualPictureDrawingProperties
                                    (
                                        new D.PictureLocks() { NoChangeAspect = true }
                                    ),
                                    new ApplicationNonVisualDrawingProperties()
                                ),
                                new P.BlipFill
                                (
                                    new D.Blip
                                        (
                                            new D.NonVisualPicturePropertiesExtensionList()
                                        )
                                    { Embed = "rId2" },
                                    new D.Stretch
                                    (
                                        new D.FillRectangle()
                                    )
                                ),
                                new P.ShapeProperties
                                (
                                    new D.Transform2D
                                    (
                                        new D.Offset() { X = cm64(imlx), Y = cm64(imly) },
                                        new D.Extents() { Cx = cm64(imwr), Cy = cm64(imhr) }
                                    ),
                                    new D.PresetGeometry
                                    (
                                        new D.AdjustValueList()
                                    )
                                    { Preset = D.ShapeTypeValues.Rectangle }
                                )
                            )
                                        )),
                    new ColorMapOverride(new MasterColorMapping()));
            ImagePart ip = slidePart1.AddImagePart(ImagePartType.Jpeg, "rId2");

            Stream MyStream = ToStream(image2, ImageFormat.Jpeg);
            MyStream.Position = 0;
            ip.FeedData(MyStream);
            return slidePart1;
        }

        public Stream ToStream(Image image, ImageFormat format)
        {
            var stream = new System.IO.MemoryStream();
            image.Save(stream, format);
            stream.Position = 0;
            return stream;
        }

        private SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
        {
            SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
            SlideLayout slideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
              new P.ShapeProperties(),
              new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph(new EndParagraphRunProperties()))))),
            new ColorMapOverride(new MasterColorMapping()));
            slideLayoutPart1.SlideLayout = slideLayout;
            return slideLayoutPart1;
        }

        private SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
        {
            SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
            SlideMaster slideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
              new P.ShapeProperties(),
              new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph())))),
            new P.ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
            new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
            new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
            slideMasterPart1.SlideMaster = slideMaster;

            return slideMasterPart1;
        }

        private ThemePart CreateTheme(SlideMasterPart slideMasterPart1)
        {
            ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
            D.Theme theme1 = new D.Theme() { Name = "Office Theme" };

            D.ThemeElements themeElements1 = new D.ThemeElements(
            new D.ColorScheme(
              new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
              new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
              new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
              new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
              new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
              new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
              new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
              new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
              new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
              new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
              new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
              new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" }))
            { Name = "Office" },
              new D.FontScheme(
              new D.MajorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" }),
              new D.MinorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" }))
              { Name = "Office" },
              new D.FormatScheme(
              new D.FillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
                  new D.SaturationModulation() { Val = 300000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
                 new D.SaturationModulation() { Val = 300000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 35000 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
                 new D.SaturationModulation() { Val = 350000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 100000 }
                ),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.NoFill(),
              new D.PatternFill(),
              new D.GroupFill()),
              new D.LineStyleList(
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              }),
              new D.EffectStyleList(
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
              new D.BackgroundFillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true })))
              { Name = "Office" });

            theme1.Append(themeElements1);
            theme1.Append(new D.ObjectDefaults());
            theme1.Append(new D.ExtraColorSchemeList());

            themePart1.Theme = theme1;
            return themePart1;

        }

        private void ofb1_Click(object sender, EventArgs e)
        {
            DialogResult result = this.fbd1.ShowDialog();
            if (result == DialogResult.OK)
            {
                tb1.Text = fbd1.SelectedPath;
            }
        }

        private void obf2_Click(object sender, EventArgs e)
        {
            DialogResult result = this.fbd2.ShowDialog();
            if (result == DialogResult.OK)
            {
                tb2.Text = fbd2.SelectedPath;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cb1.Items.Add("Screen4x3");//   0
            cb1.Items.Add("Letter");//  1
            cb1.Items.Add("A4");//2
            cb1.Items.Add("Film35mm");//    3
            cb1.Items.Add("Overhead");//    4
            cb1.Items.Add("Banner");//  5
            cb1.Items.Add("Custom");//  6
            cb1.Items.Add("Ledger");//  7
            cb1.Items.Add("A3");//8
            cb1.Items.Add("B4ISO");//   9
            cb1.Items.Add("B5ISO");//   10
            cb1.Items.Add("B4JIS");//   11
            cb1.Items.Add("B5JIS");//   12
            cb1.Items.Add("HagakiCard");//  13
            cb1.Items.Add("Screen16x9");//  14
            cb1.Items.Add("Screen16x10");// 15
            cb1.SelectedIndex = 14;


        }
    }
}
