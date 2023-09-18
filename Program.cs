using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using TextBody = DocumentFormat.OpenXml.Presentation.TextBody;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;
using System.Threading;
using NonVisualDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties;
using Paragraph = DocumentFormat.OpenXml.Drawing.Paragraph;
using Run = DocumentFormat.OpenXml.Drawing.Run;
using RunProperties = DocumentFormat.OpenXml.Drawing.RunProperties;
using FontSize = DocumentFormat.OpenXml.Spreadsheet.FontSize;
using ParagraphProperties = DocumentFormat.OpenXml.Drawing.ParagraphProperties;
using VerticalAlignmentValues = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues;
using HorizontalAlignmentValues = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues;

namespace OpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please insert the soure file path:");
            string filePath = Console.ReadLine();

            using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
            {
                // Access the main presentation part.
                PresentationPart presentationPart = presentationDocument.PresentationPart;

                // Access the presentation.
                Presentation presentation = presentationPart.Presentation;

                // Get the first slide.
                SlideId firstSlideId = presentationPart.Presentation.SlideIdList.ChildElements[0] as SlideId;
                SlidePart firstSlidePart = (SlidePart)presentationPart.GetPartById(firstSlideId.RelationshipId);
                Slide slide = firstSlidePart.Slide;

                foreach (Shape shape in slide.Descendants<Shape>())
                {
                    NonVisualDrawingProperties nonVisualProps = shape.NonVisualShapeProperties.NonVisualDrawingProperties;

                    if (nonVisualProps != null)
                    {
                        string shapeName = nonVisualProps.Name;
                        string shapeType = nonVisualProps.Id;

                        Console.WriteLine($"Shape Name: {shapeName}");
                        Console.WriteLine($"Shape Type ID: {shapeType}");
                        if(shape.ShapeProperties.Transform2D != null) Console.WriteLine($"Shape Offset X: {shape.ShapeProperties.Transform2D.Offset.X}");
                        if (shape.ShapeProperties.Transform2D != null) Console.WriteLine($"Shape Offset max X: {shape.ShapeProperties.Transform2D.Offset.X + (int)shape.ShapeProperties.Transform2D.Extents.Cx}");
                        if (shape.ShapeProperties.Transform2D != null) Console.WriteLine($"Shape Offset Y: {shape.ShapeProperties.Transform2D.Offset.Y}");
                        if (shape.ShapeProperties.Transform2D != null) Console.WriteLine($"Shape Offset max Y: {shape.ShapeProperties.Transform2D.Offset.Y + (int)shape.ShapeProperties.Transform2D.Extents.Cy}");
                        Console.WriteLine($"Shape Text: {shape.TextBody.InnerText}");
                        Console.WriteLine("-----------------");

                        if (shapeName.Contains("Title"))
                        {
                            if (shape.TextBody != null)
                            {
                                TextBody textBody = shape.TextBody;

                                foreach (Paragraph paragraph in textBody.Descendants<Paragraph>())
                                {
                                    foreach (Run run in paragraph.Elements<Run>())
                                    {
                                        foreach (RunProperties runProperties in run.Elements<RunProperties>())
                                        {
                                            // Change font size
                                            FontSize fontSize = runProperties.Elements<FontSize>().FirstOrDefault();
                                            if (fontSize != null)
                                            {
                                                fontSize.Val = 4400;
                                                runProperties.Bold = false;
                                            }
                                            else
                                            {
                                                runProperties.FontSize = 4400;
                                                runProperties.Bold = false;
                                            }

                                            // Change font family
                                            LatinFont latinFont = runProperties.Elements<LatinFont>().FirstOrDefault();
                                            if (latinFont != null)
                                            {
                                                latinFont.Typeface = "Beirut";
                                            }
                                            else
                                            {
                                                runProperties.Append(new LatinFont() { Typeface = "Beirut" });
                                            }

                                            run.RunProperties = runProperties;

                                            // Change alignment to center
                                            ParagraphProperties paragraphProperties = paragraph.ParagraphProperties ?? new ParagraphProperties();
                                            paragraphProperties.Alignment = TextAlignmentTypeValues.Center;
                                            paragraph.ParagraphProperties = paragraphProperties;

                                        }

                                    }
                                }

                                foreach (DocumentFormat.OpenXml.Drawing.Text textElement in textBody.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                                {
                                    Console.WriteLine(textElement.Text);
                                    textElement.Text = "Output Slide";
                                }
                                
                            }
                        }

                    }

                }
                // Save the changes.
                presentationDocument.Save();
                Console.ReadLine();
            }
        }
    }
}
