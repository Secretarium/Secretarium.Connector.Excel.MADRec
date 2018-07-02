using System.Collections.Generic;

namespace Secretarium.Excel
{
    public class MADRecContrib
    {
        public string name { get; set; }
        public object value { get; set; }
    }

    public class MADRecPutValues : MADRecContrib
    {
        public List<MADRecPutValues> values { get; set; }
    }
    public class MADRecPutLEIArgs
    {
        public string LEI { get; set; }
        public List<MADRecPutValues> values { get; set; }
    }

    public class MADRecGetLEIArgs
    {
        public string LEI { get; set; }
    }

    public class MADRecFieldResult
    {
        public string name { get; set; }
        public int total { get; set; }
        public int[] split { get; set; }
        public object contribution { get; set; }
        public int group { get; set; }
        public List<MADRecFieldResult> values { get; set; }
    }

    public class MADRecResult
    {
        public string LEI { get; set; }
        public List<MADRecFieldResult> values { get; set; }
        
        private MADRecFieldResult FindField(List<MADRecFieldResult> values, string name)
        {
            if (values == null || values.Count == 0) return null;

            var n = name.ToUpper();
            foreach (var v in values)
            {
                if (v.name.ToUpper() == n)
                    return v;
            }

            return null;
        }

        public MADRecFieldResult FindReport(params string[] fieldPath)
        {
            if (values == null || values.Count == 0 || fieldPath.Length == 0) return null;

            var v = values;
            for (var i = 0; i < fieldPath.Length - 1; i++)
            {
                var f = FindField(v, fieldPath[i]);
                if (f == null) return null;
                v = f.values;
            }

            return FindField(v, fieldPath[fieldPath.Length - 1]);
        }
    }
}
