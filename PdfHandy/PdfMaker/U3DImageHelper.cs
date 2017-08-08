using System;
using System.Drawing;
using PdfMaker.Models;
using Spire.Pdf;
using Spire.Pdf.Annotations;
using Spire.Pdf.Graphics;

namespace PdfMaker
{
    internal class U3DImageHelper : IDisposable
    {
        private readonly string _targetFile;

        private readonly PdfDocument _doc;

        public U3DImageHelper(string targetFile)
        {
            _targetFile = targetFile;
            _doc = new PdfDocument();
            if (System.IO.File.Exists(targetFile))
                _doc.LoadFromFile(targetFile);
        }

        public void AddU3DImage(U3DInfo u3DInfo)
        {
            try
            {
                // page starts from 0 as per spire.pdf
                int pageNum = u3DInfo.Page - 1;
                if (pageNum < 0) pageNum = 0;
                if (!(_doc.Pages.Count > pageNum))
                    for (int i = 0; i <= pageNum - _doc.Pages.Count; i++)
                        _doc.Pages.Add();

                PdfPageBase page = _doc.Pages[pageNum];
                string name = System.IO.Path.GetFileNameWithoutExtension(u3DInfo.U3DFile);

                int y = (int)(u3DInfo.Y - u3DInfo.Height); // adjust the position to map Y to zero by reducing its height - weird ah...
                Rectangle rt = new Rectangle((int)u3DInfo.X, y, (int)u3DInfo.Width, (int)u3DInfo.Height);
                Pdf3DAnnotation annotation =
                    new Pdf3DAnnotation(rt, u3DInfo.U3DFile)
                    {
                        Activation = new Pdf3DActivation { ActivationMode = Pdf3DActivationMode.PageOpen }
                    };
                Pdf3DView view = new Pdf3DView
                {
                    Background = new Pdf3DBackground(new PdfRGBColor(Color.White)),
                    ViewNodeName = name,
                    RenderMode = new Pdf3DRendermode(Pdf3DRenderStyle.Solid),
                    InternalName = name,
                    LightingScheme = new Pdf3DLighting { Style = Pdf3DLightingStyle.Headlamp }
                };
                annotation.Views.Add(view);

                page.AnnotationsWidget.Add(annotation);
            }
            catch
            {
                // ignored
            }
        }

        public void Dispose()
        {
            _doc?.SaveToFile(_targetFile, FileFormat.PDF);
        }
    }
}
