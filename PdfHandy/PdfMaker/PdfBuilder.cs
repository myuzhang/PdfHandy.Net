using iText.Forms;
using iText.Forms.Fields;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using PdfMaker.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

namespace PdfMaker
{
    public class PdfBuilder : IDisposable
    {
        private PdfDocument _pdfDoc;

        private Document _doc;

        private PdfAcroForm _form;

        private string _temporaryFile;

        private IList<U3DInfo> _u3DInfos;

        public PdfBuilder(string targetFile, string sourceFile)
        {
            Init(targetFile, sourceFile);
            RegisterCustomFonts("Fonts");
        }

        public string SourceFile { get; private set; }

        public string TargetFile { get; private set; }

        public string DefaultFont { get; set; } = "helvetica";

        public void FillMultiLinesInAcroForm<T>(T detail, int fontSize)
        {
            IDictionary<String, PdfFormField> fields = _form.GetFormFields();

            Type type = detail.GetType();
            PropertyInfo[] properties = type.GetProperties();

            foreach (PropertyInfo property in properties)
            {
                PdfFormField toSet;
                PdfActionAttribute required = property.GetCustomAttribute<PdfActionAttribute>();

                if (fields.TryGetValue(property.Name, out toSet))
                {
                    var value = property.GetValue(detail, null);
                    if (!string.IsNullOrWhiteSpace(value?.ToString()))
                    {
                        if (required != null && required.Action.Equals(ActionFlag.MultiLines))
                        {
                            PdfFontAttribute pdfFont = property.GetCustomAttribute<PdfFontAttribute>();
                            var font = pdfFont?.Type != null
                                ? PdfFontFactory.CreateRegisteredFont(pdfFont.Type)
                                : PdfFontFactory.CreateRegisteredFont(DefaultFont);
                            toSet.SetValue(value.ToString()).SetFontAndSize(font, fontSize);

                            if (pdfFont != null)
                                toSet.SetColor(pdfFont.Color);
                        }
                    }
                }
            }
        }

        public int GetAllignedFontSizeForMultiLines<T>(T detail, int defaultFontSize = 9)
        {
            IDictionary<String, PdfFormField> fields = _form.GetFormFields();

            Type type = detail.GetType();
            PropertyInfo[] properties = type.GetProperties();
            int targetFontsize = 0;

            foreach (PropertyInfo property in properties)
            {
                PdfFormField toSet;
                PdfActionAttribute required = property.GetCustomAttribute<PdfActionAttribute>();

                if (fields.TryGetValue(property.Name, out toSet))
                {
                    var value = property.GetValue(detail, null);
                    if (!string.IsNullOrWhiteSpace(value?.ToString()))
                    {
                        if (required != null && required.Action.Equals(ActionFlag.MultiLines))
                        {
                            PdfFormField field = fields[property.Name];
                            PdfArray position = field.GetWidgets().First().GetRectangle();
                            float width = (float)(position.GetAsNumber(2).GetValue() - position.GetAsNumber(0).GetValue());
                            float height = (float)(position.GetAsNumber(3).GetValue() - position.GetAsNumber(1).GetValue());

                            PdfFontAttribute pdfFont = property.GetCustomAttribute<PdfFontAttribute>();
                            var font = pdfFont?.Type != null
                                ? PdfFontFactory.CreateRegisteredFont(pdfFont.Type)
                                : PdfFontFactory.CreateRegisteredFont(DefaultFont);
                            int latestTargetFontSize = GetFontSizeFittingMultiLinesInAcroForm(value.ToString(), width, height, font, defaultFontSize);
                            if (targetFontsize == 0) targetFontsize = latestTargetFontSize;
                            else
                                targetFontsize = targetFontsize > latestTargetFontSize
                                    ? latestTargetFontSize
                                    : targetFontsize;
                        }
                    }
                }
            }
            return targetFontsize;
        }

        public int GetFontSizeFittingMultiLinesInAcroForm(string multiLineString, float targetFieldWidth, float targetFieldHeight, PdfFont pdfFont, int defaultFontSize)
        {
            if (string.IsNullOrWhiteSpace(multiLineString)) return defaultFontSize;

            List<string> stringLines = new List<string>(multiLineString.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries));

            int lineNumRequired = 0;
            int targetSize = defaultFontSize;
            while (targetSize > 1)
            {
                foreach (var stringLine in stringLines)
                {
                    float lineWidth = pdfFont.GetWidth(stringLine, targetSize);
                    int lineNumNeeded = (int)Math.Ceiling(lineWidth / targetFieldWidth);
                    lineNumRequired += lineNumNeeded;
                }

                float aboveBaseline = pdfFont.GetAscent(stringLines[0], targetSize);
                float underBaseline = pdfFont.GetDescent(stringLines[0], targetSize);
                float textHeight = aboveBaseline - underBaseline;
                float lineHeight = textHeight * 1.5f;
                int lineNumProvided = (int)Math.Floor(targetFieldHeight / lineHeight);

                if (lineNumProvided > lineNumRequired) break;

                lineNumRequired = 0;
                targetSize = targetSize - 1;
            }

            return targetSize;
        }

        public void FillFieldInAcroForm<T>(T detail)
        {
            IDictionary<String, PdfFormField> fields = _form.GetFormFields();

            Type type = detail.GetType();
            PropertyInfo[] properties = type.GetProperties();
            PdfFont defaultFont = PdfFontFactory.CreateRegisteredFont(DefaultFont);

            foreach (PropertyInfo property in properties)
            {
                PdfFormField toSet;
                PdfActionAttribute required = property.GetCustomAttribute<PdfActionAttribute>();

                if (required != null &&
                    (required.Action.Equals(ActionFlag.Ignore) ||
                     required.Action.Equals(ActionFlag.MultiLines)))
                    continue;

                if (fields.TryGetValue(property.Name, out toSet))
                {
                    var value = property.GetValue(detail, null);
                    if (!string.IsNullOrWhiteSpace(value?.ToString()))
                    {
                        toSet.SetValue(value.ToString()).SetReadOnly(true);

                        PdfFontAttribute pdfFont = property.GetCustomAttribute<PdfFontAttribute>();
                        if (pdfFont != null)
                        {
                            toSet.SetFont(pdfFont.Type != null ? PdfFontFactory.CreateRegisteredFont(pdfFont.Type) : defaultFont);

                            if (pdfFont.Size == 0) toSet.SetFontSizeAutoScale();
                            else toSet.SetFontSize(pdfFont.Size);

                            toSet.SetColor(pdfFont.Color).SetJustification(pdfFont.Justification);

                            if (!string.IsNullOrWhiteSpace(pdfFont.WarningCondition))
                            {
                                if (pdfFont.IsWarning(value.ToString(), pdfFont.WarningCondition))
                                {
                                    toSet.SetColor(Color.RED);
                                }
                            }
                        }
                        else
                        {
                            toSet.SetValue(value.ToString()).SetFont(defaultFont).SetFontSizeAutoScale().SetReadOnly(true);
                        }
                    }
                    else
                    {
                        if (required != null)
                        {
                            if (required.Action.Equals(ActionFlag.Required))
                                throw new ArgumentException($"{property.Name} is missing");
                            if (required.Action.Equals(ActionFlag.Optional))
                                continue;
                            if (required.Action.Equals(ActionFlag.NotAvailable))
                                toSet.SetValue("N/A").SetFontSizeAutoScale().SetReadOnly(true);
                        }
                    }
                }
            }
        }

        public void FillImageInAcroForm<T>(T detail)
        {
            IDictionary<String, PdfFormField> fields = _form.GetFormFields();

            Type type = detail.GetType();
            PropertyInfo[] properties = type.GetProperties();

            foreach (PropertyInfo property in properties)
            {
                PdfFormField toSet;
                PdfActionAttribute required = property.GetCustomAttribute<PdfActionAttribute>();

                if (required != null &&
                    (required.Action.Equals(ActionFlag.Ignore) ||
                     required.Action.Equals(ActionFlag.MultiLines)))
                    continue;

                if (fields.TryGetValue(property.Name, out toSet))
                {
                    var value = property.GetValue(detail, null);
                    PdfFormField field = fields[property.Name];
                    if (!string.IsNullOrWhiteSpace(value?.ToString()))
                    {
                        var widgets = field.GetWidgets();
                        if (widgets == null || widgets.Count == 0)
                            throw new ArgumentNullException($"no widgets to the field");

                        PdfArray position = widgets.First().GetRectangle();

                        PdfPage page = field.GetWidgets().First().GetPage();
                        if (page == null)
                            throw new ArgumentNullException(
                                $"field widget annotation is not associated with any page");
                        int pageNum = _pdfDoc.GetPageNumber(page);

                        float width = (float)(position.GetAsNumber(2).GetValue() - position.GetAsNumber(0).GetValue());
                        float height = (float)(position.GetAsNumber(3).GetValue() - position.GetAsNumber(1).GetValue());

                        Image image = new Image(ImageDataFactory.Create(File.ReadAllBytes(value.ToString())));
                        image.ScaleToFit(width, height);

                        float startX = (float)position.GetAsNumber(0).GetValue();
                        float startY = (float)position.GetAsNumber(1).GetValue();

                        PdfFontAttribute pdfFont = property.GetCustomAttribute<PdfFontAttribute>();
                        if (pdfFont != null)
                            ImagePositionAdjustification(pdfFont.ImageJustification, width, height, image.GetImageScaledWidth(), image.GetImageScaledHeight(),
                            (float)position.GetAsNumber(0).GetValue(), (float)position.GetAsNumber(1).GetValue(), out startX, out startY);

                        image.SetFixedPosition(pageNum, startX, startY);
                        _form.RemoveField(field.GetFieldName().ToString());
                        _doc.Add(image);
                    }
                    else
                    {
                        if (required != null)
                        {
                            if (required.Action.Equals(ActionFlag.Required))
                                throw new ArgumentNullException($"No {property.Name} image found");
                            if (required.Action.Equals(ActionFlag.Optional))
                            {
                                _form.RemoveField(field.GetFieldName().ToString());
                                continue;
                            }
                            if (required.Action.Equals(ActionFlag.NotAvailable))
                                toSet.SetValue("N/A").SetFontSizeAutoScale().SetReadOnly(true);
                        }
                    }
                }
            }
        }

        public void FillU3DImageInAcroForm<T>(T detail)
        {
            IDictionary<String, PdfFormField> fields = _form.GetFormFields();

            Type type = detail.GetType();
            PropertyInfo[] properties = type.GetProperties();

            foreach (PropertyInfo property in properties)
            {
                PdfFormField toSet;
                PdfActionAttribute required = property.GetCustomAttribute<PdfActionAttribute>();
                if (fields.TryGetValue(property.Name, out toSet))
                {
                    var value = property.GetValue(detail, null)?.ToString();
                    PdfFormField field = fields[property.Name];
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        var widgets = field.GetWidgets();
                        if (widgets == null || widgets.Count == 0)
                            throw new ArgumentNullException($"no widgets to the field");

                        PdfArray position = widgets.First().GetRectangle();

                        PdfPage page = field.GetWidgets().First().GetPage();
                        if (page == null)
                            throw new ArgumentNullException(
                                $"field widget annotation is not associated with any page");
                        int pageNum = _pdfDoc.GetPageNumber(page);
                        float width = (float)(position.GetAsNumber(2).GetValue() -
                                              position.GetAsNumber(0).GetValue());
                        float height = (float)(position.GetAsNumber(3).GetValue() -
                                               position.GetAsNumber(1).GetValue());

                        var u3DInfo = new U3DInfo
                        {
                            Height = height,
                            Width = width,
                            X = (float)position.GetAsNumber(0).GetValue(),
                            Y = (float)position.GetAsNumber(1).GetValue(),
                            Page = pageNum,
                            U3DFile = value
                        };
                        _u3DInfos.Add(u3DInfo);

                        _form.RemoveField(field.GetFieldName().ToString());
                    }
                    else
                    {
                        if (required != null)
                        {
                            if (required.Action.Equals(ActionFlag.Required))
                                throw new ArgumentNullException($"No {property.Name} image found");
                            if (required.Action.Equals(ActionFlag.Optional))
                            {
                                _form.RemoveField(field.GetFieldName().ToString());
                                continue;
                            }
                            if (required.Action.Equals(ActionFlag.NotAvailable))
                                toSet.SetValue("N/A").SetFontSizeAutoScale().SetReadOnly(true);
                        }
                    }
                }
            }

            // close the target pdf first:
            _pdfDoc?.Close();

            using (var helper = new U3DImageHelper(TargetFile))
            {
                foreach (var u3DInfo in _u3DInfos)
                    helper.AddU3DImage(u3DInfo);
            }

            // reopen the target pdf:
            _temporaryFile = Path.Combine(Path.GetDirectoryName(TargetFile), $"temp_{DateTime.Now.Ticks}.pdf");
            File.Move(TargetFile, _temporaryFile);
            Init(TargetFile, _temporaryFile);
        }

        public void RegisterCustomFont(string fontFile) => PdfFontFactory.Register(fontFile);

        public void RegisterCustomFonts(string fontFolder)
        {
            if (Directory.Exists(fontFolder))
            {
                DirectoryInfo hdDirectoryInWhichToSearch = new DirectoryInfo(fontFolder);
                FileInfo[] files = hdDirectoryInWhichToSearch.GetFiles();
                foreach (var fileInfo in files)
                {
                    try
                    {
                        PdfFontFactory.Register(fileInfo.FullName);
                    }
                    catch (Exception e)
                    {
                        Debug.WriteLine(e);
                    }
                }
            }
        }

        public void Dispose()
        {
            _pdfDoc?.Close();
            if (!string.IsNullOrWhiteSpace(_temporaryFile))
                if (File.Exists(_temporaryFile))
                    File.Delete(_temporaryFile);
        }

        private void ImagePositionAdjustification(int adjustification, float formWidth, float formHeight, float imageWidth, float imageHeight, float formX, float formY, out float imageX, out float imageY)
        {
            imageX = formX;
            imageY = formY;

            if (adjustification == 0)
            {
                return;
            }

            float offset;

            if (adjustification == 1 || adjustification == 2)
            {
                if (imageWidth >= formWidth)
                {
                    imageX = formX;
                    imageY = formY;
                    return;
                }

                offset = adjustification == 1 ? (formWidth - imageWidth) / 2 : formWidth - imageWidth;
                imageX = formX + offset;
                imageY = formY;
                return;
            }

            if (adjustification == 3 || adjustification == 4 || adjustification == 5)
            {
                if (imageHeight >= formHeight)
                {
                    imageY = formY;
                }
                else
                {
                    offset = formHeight - imageHeight;
                    imageY = formY + offset;
                }

                if (adjustification == 3 || imageWidth >= formWidth)
                {
                    imageX = formX;
                    return;
                }

                offset = adjustification == 4 ? (formWidth - imageWidth) / 2 : formWidth - imageWidth;
                imageX = formX + offset;
                return;
            }

            if (imageHeight >= formHeight)
            {
                imageY = formY;
            }
            else
            {
                offset = (formHeight - imageHeight) / 2;
                imageY = formY + offset;
            }

            if (imageWidth >= formWidth)
            {
                imageX = formX;
            }
            else
            {
                offset = (formWidth - imageWidth) / 2;
                imageY = formY + offset;
            }
        }

        private void Init(string targetFile, string sourceFile)
        {
            if (string.IsNullOrWhiteSpace(targetFile) && string.IsNullOrWhiteSpace(sourceFile))
                throw new ArgumentNullException($"Target file and source file can't be empty at the same time.");

            SourceFile = sourceFile;
            TargetFile = targetFile;

            if (string.IsNullOrWhiteSpace(sourceFile))
            {
                var pdfWriter = new PdfWriter(targetFile);
                _pdfDoc = new PdfDocument(pdfWriter);
            }
            else if (string.IsNullOrWhiteSpace(targetFile))
            {
                var pdfReader = new PdfReader(sourceFile);
                _pdfDoc = new PdfDocument(pdfReader);
            }
            else
            {
                var pdfWriter = new PdfWriter(targetFile);
                var pdfReader = new PdfReader(sourceFile);
                _pdfDoc = new PdfDocument(pdfReader, pdfWriter);
            }

            _doc = new Document(_pdfDoc);
            _form = PdfAcroForm.GetAcroForm(_pdfDoc, true);
            _u3DInfos = new List<U3DInfo>();
        }
    }
}
