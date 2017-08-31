using PdfMaker;
using PdfMaker.Models;

namespace Test
{
    public class ExampleModel
    {
        [PdfAction(ActionFlag.Required)]
        [PdfFont("Yellow", "helvetica", 18, Justification = PdfTextAlignment.TopMiddle)]
        public string Example1 { get; set; }

        [PdfAction(ActionFlag.Optional)]
        [PdfFont("White", "helvetica-bold", 16)]
        public string Example2 { get; set; }

        [PdfAction(ActionFlag.Required)]
        public string Example3 { get; set; }
    }
}
