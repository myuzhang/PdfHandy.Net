namespace PdfMaker.Models
{
    public class U3DInfo
    {
        // from left bottom corner
        public float X { get; set; }

        // from left bottom corner
        public float Y { get; set; }
        
        public float Width { get; set; }
        
        public float Height { get; set; }

        // Page starts from 1 as per iText7
        // Page starts from 0 as per Spire
        public int Page { get; set; }

        public string U3DFile { get; set; }
    }
}
