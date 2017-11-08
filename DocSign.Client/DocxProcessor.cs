using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OXMLD = DocumentFormat.OpenXml.Drawing;

using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System;

namespace DocSign.Client
{
    public class DocxProcessor : IDisposable
    {
        private string _docxFilePath;
        private WordprocessingDocument _wordDocx;

        public DocxProcessor(string docxFilePath)
        {
            _docxFilePath = docxFilePath;
            _wordDocx = WordprocessingDocument.Open(_docxFilePath, true);
        }

        public void MergeFields(List<AppField> appFields)
        {
            if (!File.Exists(_docxFilePath))
                throw new FileNotFoundException(_docxFilePath);

            foreach (var appField in appFields)
            {
                //ref: http://officeopenxml.com/WPfields.php
                foreach (var documentPart in _wordDocx.MainDocumentPart.Document)
                {
                    foreach (var ff in documentPart.Descendants<FieldChar>()) //This could also be a SimpleField 
                    {
                        if (ff.FormFieldData != null)
                            foreach (var ffd in ff.FormFieldData.Descendants<FormFieldName>())
                                if (appField.FieldID == ffd.Val.Value)
                                {
                                    //found my field!!!

                                    string appFieldName = appField.FieldID;
                                    string appFieldValue = appField.Value; //this is the file path

                                    var element = GenerateImageElement(_wordDocx, appFieldValue , appField.SignatureTag);

                                    //Get all runs for the Field
                                    List<Run> fieldRuns = new List<Run>();
                                    Run parentRun = (Run)ff.Parent;
                                    Run textRun = null;

                                    fieldRuns.Add(parentRun); //FieldCharType == "begin"

                                    //loop through runs between 'begin' and 'end'
                                    foreach (Run runSubling in parentRun.ElementsAfter().OfType<Run>())
                                    {
                                        fieldRuns.Add(runSubling);

                                        var sfch = runSubling.Descendants<FieldChar>().FirstOrDefault();

                                        if (runSubling.Descendants<Text>().Any())
                                        {
                                            //this is the run with my text value
                                            //runSubling.Descendants<Text>().First().Text = "test value";// to change the value
                                            textRun = runSubling;
                                        }

                                        if (runSubling.Descendants<FieldChar>().Any(fc => fc.FieldCharType == "end"))
                                            break; //last run
                                    }

                                    if (textRun != null)
                                    {
                                        //Insert image element after the last run of the field. This element should also be a Run
                                        textRun.Parent.InsertAfter<Run>(new Run(element), textRun);

                                        // We can also append the reference to body.
                                        //wordDocx.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));

                                        //Remove my field run elements. I am replacing them with an image
                                        foreach (var r in fieldRuns)
                                            r.Remove();
                                    }

                                }
                    }

                    //foreach (var sf in documentPart.Descendants<SimpleField>()) //This could also be a SimpleField 
                }
            }
        }

        /* FormField in XML
         *   <w:r w:rsidRPr="00E62C69" w:rsidR="00E62C69">
              <w:rPr>
                <w:u w:val="single"/>
              </w:rPr>
              <w:fldChar w:fldCharType="begin">
                <w:ffData>
                  <w:name w:val="f7651_0002"/>
                  <w:enabled/>
                  <w:calcOnExit w:val="0"/>
                  <w:helpText w:type="text" w:val="7651"/>
                  <w:statusText w:type="text" w:val="Do not change this field manually!"/>
                  <w:textInput>
                    <w:default w:val="&lt;Signature&gt;"/>
                  </w:textInput>
                </w:ffData>
              </w:fldChar>
            </w:r>
            <w:bookmarkStart w:name="f7651_0002" w:id="0"/>
            <w:r w:rsidRPr="00E62C69" w:rsidR="00E62C69">
              <w:rPr>
                <w:u w:val="single"/>
              </w:rPr>
              <w:instrText xml:space="preserve"> FORMTEXT </w:instrText>
            </w:r>
            <w:r w:rsidRPr="00E62C69" w:rsidR="00E62C69">
              <w:rPr>
                <w:u w:val="single"/>
              </w:rPr>
            </w:r>
            <w:r w:rsidRPr="00E62C69" w:rsidR="00E62C69">
              <w:rPr>
                <w:u w:val="single"/>
              </w:rPr>
              <w:fldChar w:fldCharType="separate"/>
            </w:r>
            <w:r w:rsidRPr="00E62C69" w:rsidR="00E62C69">
              <w:rPr>
                <w:u w:val="single"/>
              </w:rPr>
              <w:t>&lt;Signature&gt;</w:t>
            </w:r>
            <w:r w:rsidRPr="00E62C69" w:rsidR="00E62C69">
              <w:rPr>
                <w:u w:val="single"/>
              </w:rPr>
              <w:fldChar w:fldCharType="end"/>
            </w:r>
         */

        /// <summary>
        /// Ref: https://code.msdn.microsoft.com/office/CSManipulateImagesInWordDoc-312da7ef
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <returns></returns>
        private static IEnumerable<OXMLD.Blip> GetDocImages(WordprocessingDocument wordDoc)
        {
            // Get the drawing elements in the document.
            var drawingElements = from run in wordDoc.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>()
                                  where run.Descendants<Drawing>().Count() != 0
                                  select run.Descendants<Drawing>().First();

            // Get the blip elements in the drawing elements.
            var blipElements = from drawing in drawingElements
                               where drawing.Descendants<A.Blip>().Count() > 0
                               select drawing.Descendants<A.Blip>().First();

            //blip = wordDoc.MainDocumentPart.Document.Descendants<OXMLPIC.Picture>().
            //               Where(p => signatureFileName == p.NonVisualPictureProperties.NonVisualDrawingProperties.Name).
            //               Select(p => p.BlipFill.Blip).
            //               SingleOrDefault();

            return blipElements;
        }

        private Drawing GenerateImageElement(WordprocessingDocument wordDoc, string signatureFile, string sigToken)
        {
            string relationshipId = "";
            // Define the reference of the image.
            var templateImg = wordDoc.MainDocumentPart.AddImagePart(ImagePartType.Bmp);
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;

            //Load ImagePart with signature image.
            using (Stream signImageData = new FileStream(signatureFile, FileMode.Open))
            {
                templateImg.FeedData(signImageData);
            }
            relationshipId = wordDoc.MainDocumentPart.GetIdOfPart(templateImg);

            // Get the dimensions of the image in English Metric Units (EMU)
            // for use when adding the markup for the image to the document.
            //var imageWidthEMU = (long)((imageFile.Width / imageFile.HorizontalResolution) * 914400L);
            //var imageHeightEMU = (long)((imageFile.Height / imageFile.VerticalResolution) * 914400L);
            long imageWidthEMU = 1724025;
            long imageHeightEMU = 542925;

            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = imageWidthEMU, Cy = imageHeightEMU },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1",
                             Title = sigToken
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "ShowCase_DocSign.bmp"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(

                                         //new A.BlipExtensionList(
                                         //    new A.BlipExtension()
                                         //    {
                                         //        Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                         //    })
                                         )
                                         {
                                             Embed = relationshipId,
                                             //CompressionState = A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = imageWidthEMU, Cy = imageHeightEMU }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         {
                                             Preset = A.ShapeTypeValues.Rectangle
                                         }))
                             )
                             {
                                 Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                             })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     }
                     );

            return element;
        }

        public void Dispose()
        {
            if (_wordDocx != null)
            {
                _wordDocx.Save();
                _wordDocx.Dispose();
            }
        }
    }
}
