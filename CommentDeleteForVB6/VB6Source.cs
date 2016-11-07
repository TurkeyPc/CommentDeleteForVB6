using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;

namespace CommentDeleteForVB6
{
    public class VB6LogicalRow
    {
        private List<string> mPhysicalRows;
        private const string ContinueString = " _";

        public VB6LogicalRow()
        {
            mPhysicalRows = new List<string>();
        }

        public void Add(string row)
        {
            mPhysicalRows.Add(row);
        }

        public IEnumerable<string> PhysicalRows()
        {
            return mPhysicalRows;
        }

        public int Count
        {
            get { return mPhysicalRows.Count(); }
        }

        public bool CanAddRow
        {
            get
            {
                if (Count == 0) return true;
                return IsContinueRow(GetLastItem(mPhysicalRows));
            }
        }

        private bool IsContinueRow(string p)
        {
            if (p.Length < ContinueString.Length) return false;

            var s = new string(p.Reverse().Take(ContinueString.Length).Reverse().ToArray());
            return s == ContinueString;
        }

        public string GetLastItem(IEnumerable<string> p)
        {
            return p.Reverse().FirstOrDefault();
        }

        private string CutLast(string p,int v)
        {
            if (p.Length <= v) return "";

            return p.Substring(0, p.Length - v);
        }

        public IEnumerable<string> AliveCode()
        {
            var vvv = DeleteComment2(mPhysicalRows).ToList();

            if (vvv.Count >= 1)
                if (GetLastItem(vvv).Trim() == "")
                    vvv.RemoveAt(vvv.Count - 1);

            if (vvv.Count >= 1)
            {
                if (IsContinueRow(GetLastItem(vvv)))
                {
                    var lastindex = vvv.Count - 1;
                    vvv[lastindex] = CutLast(vvv[lastindex], 2);
                }
            }

            return vvv;
        }

        private static IEnumerable<string> DeleteComment2(IEnumerable<string> s)
        {
            var v = new List<string>(s);

            foreach (var vvv in s)
            {
                var t = DeleteComment(vvv);
                yield return t;
                if (t != vvv) yield break;
            }
        }

        private static string DeleteComment(string s)
        {
            var c = s.IndexOf("'");

            if (c == -1)
                return s;

            if (s.IndexOf("\"") == -1)
                return s.Substring(0, c);

            bool InString = false;

            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == '"')
                    InString = !InString;

                if (!InString && s[i] == '\'')
                    return s.Substring(0, i);
            }

            return s;
        }
    }

    public class VB6Source
    {
        private readonly string[] source;

        public VB6Source(string path)
        {
            this.source = File.ReadAllLines(path, Encoding.Default);
        }

        public VB6Source(IEnumerable<string> source)
        {
            this.source = source.ToArray();
        }

        public IEnumerable<string> CommentDeleted
        {
            get
            {
                return PhysicalRows(LogicalRows());
            }
        }

        private IEnumerable<VB6LogicalRow> LogicalRows()
        {
            var v = new VB6LogicalRow();

            foreach (var sss in source)
            {
                if (v.CanAddRow)
                    v.Add(sss);

                if (!v.CanAddRow)
                {
                    yield return v;
                    v = new VB6LogicalRow();
                }
            }

            if (v.Count > 0)
                yield return v;
        }

        private IEnumerable<string> PhysicalRows(IEnumerable<VB6LogicalRow> s)
        {
            foreach (var logicalrow in s)
                foreach (var physicalrow in logicalrow.AliveCode())
                    yield return physicalrow;
        }
    }
}
