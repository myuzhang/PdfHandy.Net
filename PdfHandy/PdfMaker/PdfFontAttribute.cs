using iText.Kernel.Colors;
using System;
using System.Text.RegularExpressions;

namespace PdfMaker
{
    [AttributeUsage(AttributeTargets.Property)]
    public class PdfFontAttribute : Attribute
    {
        private readonly string _color;

        private string _type;

        public PdfFontAttribute()
        {
            _color = string.Empty;
            _type = string.Empty;
            Size = 0;
            Justification = 0;
            ImageJustification = 0;
            WarningCondition = string.Empty;
        }

        public PdfFontAttribute(string color)
        {
            _color = color;
            _type = string.Empty;
            Size = 0;
            Justification = 0;
            ImageJustification = 0;
            WarningCondition = string.Empty;
        }

        public PdfFontAttribute(string color, string type, int size)
        {
            _color = color;
            _type = type;
            Size = size;
            Justification = 0;
            ImageJustification = 0;
            WarningCondition = string.Empty;
        }

        public Color Color
        {
            get
            {
                if (string.IsNullOrWhiteSpace(_color))
                    return Color.BLACK;

                switch (_color.ToLower())
                {
                    case "yellow": return Color.YELLOW;
                    case "red": return Color.RED;
                    case "white": return Color.WHITE;
                    default: return Color.BLACK;
                }
            }
        }

        // Available pdf font in-build:
        //[0]: "courier"
        //[1]: "courier-bold"
        //[2]: "courier-oblique"
        //[3]: "courier-boldoblique"
        //[4]: "helvetica"
        //[5]: "helvetica-bold"
        //[6]: "helvetica-oblique"
        //[7]: "helvetica-boldoblique"
        //[8]: "symbol"
        //[9]: "times-roman"
        //[10]: "times-bold"
        //[11]: "times-italic"
        //[12]: "times-bolditalic"
        //[13]: "zapfdingbats"
        // Available pdf font installed:
        //[14]: "helveticaneueltstd-bd"
        //[15]: "helvetica neue lt std 75 bold"
        //[16]: "helveticaneueltstd-lt"
        //[17]: "helvetica neue lt std 45 light"
        public string Type
        {
            get
            {
                if (string.IsNullOrWhiteSpace(_type))
                    return null;

                return _type.ToLower();
            }
            set { _type = value; }
        }

        public int Size { get; set; }

        public int Justification { get; set; }

        //0: original point for bottom left justified - default behavior
        //1: original point for bottom middle justified
        //2: original point for bottom right justified
        //3: original point for top left justified
        //4: original point for top middle justified
        //5: original point for top right justified
        //6: original point for center justified
        public int ImageJustification { get; set; }

        public string WarningCondition { get; set; }

        public bool IsWarning(string value, string warningCondition)
        {
            if (string.IsNullOrWhiteSpace(value) ||
                string.IsNullOrWhiteSpace(warningCondition))
                return false;

            var conditions = warningCondition.Split(',');
            foreach (var condition in conditions)
            {
                var isWarning = true;
                var subConditions = condition.Split('+');
                foreach (var subCondition in subConditions)
                {
                    var con = subCondition.Split(':');
                    if (con.Length != 2) continue;

                    isWarning = isWarning && CheckCondition(con[0], con[1], value);
                }
                if (isWarning) return true;
            }

            return false;
        }

        private bool CheckCondition(string condition, string conditionValue, string actualValue)
        {
            int actual, expected;
            switch (condition.ToLower())
            {
                case "gt":
                    if (TryAnyInt(actualValue, out actual) && TryAnyInt(conditionValue, out expected))
                        if (actual > expected) return true;
                    break;
                case "ge":
                    if (TryAnyInt(actualValue, out actual) && TryAnyInt(conditionValue, out expected))
                        if (actual >= expected) return true;
                    break;
                case "lt":
                    if (TryAnyInt(actualValue, out actual) && TryAnyInt(conditionValue, out expected))
                        if (actual < expected) return true;
                    break;
                case "le":
                    if (TryAnyInt(actualValue, out actual) && TryAnyInt(conditionValue, out expected))
                        if (actual <= expected) return true;
                    break;
                case "eq":
                    if (TryAnyInt(actualValue, out actual) && TryAnyInt(conditionValue, out expected))
                        if (actual == expected) return true;
                    break;
                case "match":
                    if (actualValue.IndexOf(conditionValue, StringComparison.CurrentCultureIgnoreCase) != -1)
                        return true;
                    break;
            }
            return false;
        }

        // return the first avaible int
        private bool TryAnyInt(string s, out int result)
        {
            if (int.TryParse(s, out result))
                return true;

            var match = Regex.Match(s, @"\d+");
            if (!match.Success) return false;

            var found = match.Groups[0].Value;
            var foundIndex = s.IndexOf(found, StringComparison.Ordinal);
            if (foundIndex != 0 && s[foundIndex - 1] == '-')
                result = 0 - int.Parse(match.Groups[0].Value);
            else
                result = int.Parse(match.Groups[0].Value);
            return true;
        }
    }
}
