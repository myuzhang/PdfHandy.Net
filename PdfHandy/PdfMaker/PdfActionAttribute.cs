using System;
using PdfMaker.Models;

namespace PdfMaker
{
    [AttributeUsage(AttributeTargets.Property)]
    public class PdfActionAttribute : Attribute
    {
        public PdfActionAttribute(ActionFlag action)
        {
            Action = action;
        }

        public ActionFlag Action { get; }
    }
}
