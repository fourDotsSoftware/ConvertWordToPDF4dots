using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Windows;
using System.Windows.Media.Imaging;

namespace ConvertWordToPDF4dots
{
    public class WordImageExtractor
    {
        public List<string> ExtractedFilepaths = new List<string>();

        public List<FromToWordImage> ExtractedFromToWordImages = new List<FromToWordImage>();

        public string err = "";

        private object missing = System.Reflection.Missing.Value;
        private object yes = true;
        private object no = false;
        private object oDocuments = null;
        private object doc = null;
        private object Shapes = null;
        private object ShapesCount = null;
        private object Shape = null;


        private object Sections = null;
        private object Headers = null;
        private object HeaderShapes = null;

        public bool ExtractImages(string filepath)
        {
            err = "";

            Image image = null;
            object WordAppSelection = null;
            object HeaderRangeShape = null;
            int iHeaderRangeShapesCount = -1;
            object HeaderRangeShapesCount = null;
            object HeaderRangeShapes = null;
            object HeaderRange = null;
            object Header = null;

            try
            {
                OfficeHelper.CreateWordApplication();

                oDocuments = OfficeHelper.WordApp.GetType().InvokeMember("Documents", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.WordApp, null);

                doc = oDocuments.GetType().InvokeMember("Open", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oDocuments, new object[] { filepath });

                System.Threading.Thread.Sleep(200);

                Sections = doc.GetType().InvokeMember("Sections", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, doc, null);

                object SectionsCount = Sections.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Sections, null);
                int iSectionsCount = (int)SectionsCount;

                for (int m1 = 1; m1 <= iSectionsCount; m1++)
                {
                    object Section = doc.GetType().InvokeMember("Sections", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, doc, new object[] { m1 });

                    Headers = Section.GetType().InvokeMember("Headers", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Section, null);

                    object HeadersCount = Headers.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Headers, null);
                    int iHeadersCount = (int)HeadersCount;

                    for (int m2 = 1; m2 <= iHeadersCount; m2++)
                    {
                        Header = Section.GetType().InvokeMember("Headers", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Section, new object[] { m2 });

                        HeaderRange = Header.GetType().InvokeMember("Range", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Header, null);

                        HeaderRangeShapes = HeaderRange.GetType().InvokeMember("InlineShapes", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRange, null);

                        HeaderRangeShapesCount = HeaderRangeShapes.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShapes, null);
                        iHeaderRangeShapesCount = (int)HeaderRangeShapesCount;

                        for (int m4 = 1; m4 <= iHeaderRangeShapesCount; m4++)
                        {
                            HeaderRangeShape = HeaderRange.GetType().InvokeMember("InlineShapes", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRange, new object[] { m4 });

                            object oShapeType = HeaderRangeShape.GetType().InvokeMember("Type", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShape, null);

                            int iShapeType = (int)oShapeType;

                            //WdInlineShapeType Enumeration (Word)

                            if (iShapeType == 1 || iShapeType == 2 || iShapeType == 4 || iShapeType == 3 || iShapeType == 7)
                            {

                                HeaderRangeShape.GetType().InvokeMember("Select", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShape, null);

                                WordAppSelection = OfficeHelper.WordApp.GetType().InvokeMember("Selection", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.WordApp, null);

                                WordAppSelection.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, WordAppSelection, null);

                                FromToWordImage wim = new FromToWordImage();
                                wim.ShapeNr = m4;
                                wim.WordFilepath = filepath;
                                wim.FromToWordImageType = FromToWordImage.FromToWordImageTypeEnum.HeaderInlineShape;

                                Thread thread = new Thread(new ParameterizedThreadStart(SaveInlineShape));
                                thread.SetApartmentState(ApartmentState.STA);
                                thread.Start(wim);
                                thread.Join();
                            }

                        }

                        // Shapes
                        /*
                        HeaderRangeShapes = Header.GetType().InvokeMember("Shapes", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Header, null);

                        HeaderRangeShapesCount = HeaderRangeShapes.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShapes, null);
                        iHeaderRangeShapesCount = (int)HeaderRangeShapesCount;

                        for (int m4 = 1; m4 <= iHeaderRangeShapesCount; m4++)
                        {
                            HeaderRangeShape = Header.GetType().InvokeMember("Shapes", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Header, new object[] { m4 });

                            HeaderRangeShape.GetType().InvokeMember("Select", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShape, new object[] { true});

                            WordAppSelection = OfficeHelper.WordApp.GetType().InvokeMember("Selection", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.WordApp, null);

                            WordAppSelection.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, WordAppSelection, null);

                            Thread thread = new Thread(SaveInlineShape);
                            thread.SetApartmentState(ApartmentState.STA);
                            thread.Start();
                            thread.Join();
                        }
                        */
                    }

                    Headers = Section.GetType().InvokeMember("Footers", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Section, null);

                    HeadersCount = Headers.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Headers, null);
                    iHeadersCount = (int)HeadersCount;

                    for (int m2 = 1; m2 <= iHeadersCount; m2++)
                    {
                        Header = Section.GetType().InvokeMember("Footers", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Section, new object[] { m2 });

                        HeaderRange = Header.GetType().InvokeMember("Range", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Header, null);

                        HeaderRangeShapes = HeaderRange.GetType().InvokeMember("InlineShapes", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRange, null);

                        HeaderRangeShapesCount = HeaderRangeShapes.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShapes, null);
                        iHeaderRangeShapesCount = (int)HeaderRangeShapesCount;

                        for (int m4 = 1; m4 <= iHeaderRangeShapesCount; m4++)
                        {
                            HeaderRangeShape = HeaderRange.GetType().InvokeMember("InlineShapes", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRange, new object[] { m4 });

                            object oShapeType = HeaderRangeShape.GetType().InvokeMember("Type", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShape, null);

                            int iShapeType = (int)oShapeType;

                            //WdInlineShapeType Enumeration (Word)

                            if (iShapeType == 1 || iShapeType == 2 || iShapeType == 4 || iShapeType == 3 || iShapeType == 7)
                            {

                                HeaderRangeShape.GetType().InvokeMember("Select", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShape, null);

                                WordAppSelection = OfficeHelper.WordApp.GetType().InvokeMember("Selection", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.WordApp, null);

                                WordAppSelection.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, WordAppSelection, null);


                                FromToWordImage wim = new FromToWordImage();
                                wim.ShapeNr = m4;
                                wim.WordFilepath = filepath;
                                wim.FromToWordImageType = FromToWordImage.FromToWordImageTypeEnum.FooterInlineShape;

                                Thread thread = new Thread(new ParameterizedThreadStart(SaveInlineShape));
                                thread.SetApartmentState(ApartmentState.STA);
                                thread.Start(wim);
                                thread.Join();
                            }
                        }

                        // Shapes                                
                        /*
                        HeaderRangeShapes = Header.GetType().InvokeMember("Shapes", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Header, null);

                        HeaderRangeShapesCount = HeaderRangeShapes.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShapes, null);
                        iHeaderRangeShapesCount = (int)HeaderRangeShapesCount;

                        for (int m4 = 1; m4 <= iHeaderRangeShapesCount; m4++)
                        {
                            HeaderRangeShape = Header.GetType().InvokeMember("Shapes", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Header, new object[] { m4 });

                            HeaderRangeShape.GetType().InvokeMember("Select", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShape, null);

                            WordAppSelection = OfficeHelper.WordApp.GetType().InvokeMember("Selection", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.WordApp, null);

                            WordAppSelection.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, WordAppSelection, null);

                            Thread thread = new Thread(SaveInlineShape);
                            thread.SetApartmentState(ApartmentState.STA);
                            thread.Start();
                            thread.Join();
                        }*/
                    }
                }

                object oContent = doc.GetType().InvokeMember("Content", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, doc, null);

                HeaderRangeShapes = oContent.GetType().InvokeMember("InlineShapes", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oContent, null);

                HeaderRangeShapesCount = HeaderRangeShapes.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShapes, null);
                iHeaderRangeShapesCount = (int)HeaderRangeShapesCount;

                for (int m4 = 1; m4 <= iHeaderRangeShapesCount; m4++)
                {
                    HeaderRangeShape = oContent.GetType().InvokeMember("InlineShapes", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oContent, new object[] { m4 });

                    object oShapeType = HeaderRangeShape.GetType().InvokeMember("Type", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShape, null);

                    int iShapeType = (int)oShapeType;

                    //WdInlineShapeType Enumeration (Word)

                    if (iShapeType == 1 || iShapeType == 2 || iShapeType == 4 || iShapeType == 3 || iShapeType == 7)
                    {
                        HeaderRangeShape.GetType().InvokeMember("Select", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShape, null);

                        WordAppSelection = OfficeHelper.WordApp.GetType().InvokeMember("Selection", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.WordApp, null);

                        WordAppSelection.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, WordAppSelection, null);


                        FromToWordImage wim = new FromToWordImage();
                        wim.ShapeNr = m4;
                        wim.WordFilepath = filepath;
                        wim.FromToWordImageType = FromToWordImage.FromToWordImageTypeEnum.DocumentInlineShape;


                        Thread thread = new Thread(new ParameterizedThreadStart(SaveInlineShape));
                        thread.SetApartmentState(ApartmentState.STA);
                        thread.Start(wim);
                        thread.Join();
                    }
                }

                // Shapes

                HeaderRangeShapes = doc.GetType().InvokeMember("Shapes", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, doc, null);

                HeaderRangeShapesCount = HeaderRangeShapes.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShapes, null);
                iHeaderRangeShapesCount = (int)HeaderRangeShapesCount;

                for (int m4 = 1; m4 <= iHeaderRangeShapesCount; m4++)
                {
                    HeaderRangeShape = doc.GetType().InvokeMember("Shapes", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, doc, new object[] { m4 });

                    object oShapeType = HeaderRangeShape.GetType().InvokeMember("Type", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShape, null);

                    int iShapeType = (int)oShapeType;

                    /*
                 'Type 7
    Case msoEmbeddedOLEObject
                  'Type 10
    Case msoLinkedOLEObject
                  'Type 11
    Case msoLinkedPicture
                   'Type 13
    Case msoPicture
                 */
                    if (iShapeType == 7 || iShapeType == 10 || iShapeType == 11 || iShapeType == 13)
                    {
                        HeaderRangeShape.GetType().InvokeMember("Select", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, HeaderRangeShape, null);

                        WordAppSelection = OfficeHelper.WordApp.GetType().InvokeMember("Selection", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.WordApp, null);

                        WordAppSelection.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, WordAppSelection, null);

                        FromToWordImage wim = new FromToWordImage();
                        wim.ShapeNr = m4;
                        wim.WordFilepath = filepath;
                        wim.FromToWordImageType = FromToWordImage.FromToWordImageTypeEnum.DocumentShape;


                        Thread thread = new Thread(new ParameterizedThreadStart(SaveInlineShape));
                        thread.SetApartmentState(ApartmentState.STA);
                        thread.Start(wim);
                        thread.Join();
                    }
                }


                return true;
            }
            catch (Exception ex)
            {
                err += TranslateHelper.Translate("Error could not Replace Image for Document") + " : " + filepath + "\r\n" + ex.Message;
                return false;
            }

            return true;
        }

        Bitmap GetBitmap(BitmapSource source)
        {
            Bitmap bmp = new Bitmap(
              source.PixelWidth,
              source.PixelHeight,
              PixelFormat.Format32bppPArgb);
            BitmapData data = bmp.LockBits(
              new Rectangle(System.Drawing.Point.Empty, bmp.Size),
              ImageLockMode.WriteOnly,
              PixelFormat.Format32bppPArgb);
            source.CopyPixels(
              Int32Rect.Empty,
              data.Scan0,
              data.Height * data.Stride,
              data.Stride);
            bmp.UnlockBits(data);
            return bmp;
        }

        protected void SaveInlineShape(object owim)
        {
            try
            {
                if (System.Windows.Clipboard.GetDataObject() != null)
                {
                    System.Windows.IDataObject data = System.Windows.Clipboard.GetDataObject();
                    if (data.GetDataPresent(System.Windows.DataFormats.Bitmap))
                    {
                        System.Windows.Interop.InteropBitmap image = (System.Windows.Interop.InteropBitmap)data.GetData(System.Windows.DataFormats.Bitmap, true);

                        Bitmap bmp=GetBitmap(image);

                        //string imgfp = System.IO.Path.Combine(Module.CurrentImagesDirectory, Guid.NewGuid().ToString() + ".bmp");

                        FromToWordImage wim = owim as FromToWordImage;

                        string imgfp = frmOptions.GetSaveFilepath(wim.WordFilepath, Module.CurrentImagesDirectory);
                            
                        ExtractedFilepaths.Add(imgfp);

                        //bmp.Save(imgfp);                                               

                        frmOptions.SaveImage(imgfp, bmp);

                        wim.ImageFilepath = imgfp;

                        ExtractedFromToWordImages.Add(wim);

                        if (frmMain.Instance.FirstOutputDocument == string.Empty)
                        {
                            frmMain.Instance.FirstOutputDocument = imgfp;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
    }

    public class FromToWordImage
    {
        public string WordFilepath = "";
        public string ImageFilepath = "";
        public int ShapeNr = -1;

        public FromToWordImageTypeEnum FromToWordImageType = FromToWordImageTypeEnum.DocumentInlineShape;

        public enum FromToWordImageTypeEnum
        {
            HeaderInlineShape,
            FooterInlineShape,
            DocumentInlineShape,
            DocumentShape
        }
    }
}
