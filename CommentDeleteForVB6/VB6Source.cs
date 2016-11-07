using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;

namespace CommentDeleteForVB6
{

    public class VB6PhysicalRow
    {
        private readonly string m;
        private const string ContinueString = " _";

        public VB6PhysicalRow(string p)
        {
            m = p;
        }

        public bool IsContinueRow
        {
            get
            {
                if (m.Length < ContinueString.Length) return false;

                var s = new string(m.Reverse().Take(ContinueString.Length).Reverse().ToArray());
                return s == ContinueString;
            }
        }

        public string FactValue
        {
            get
            {
                return m;
            }
        }

        public string EssenceValue
        {
            get
            {
                var c = DeleteComment;

                if (IsContinueRow)
                    return "";
                else
                    return CutLast(m, ContinueString.Length);
            }
        }

        private string CutLast(string p,int v)
        {
            if (p.Length <= v) return "";

            return p.Substring(0, p.Length - v);
        }

        public string DeleteComment
        {
            get
            {
                var c = m.IndexOf("'");

                if (c == -1)
                    return m;

                if (m.IndexOf("\"") == -1)
                    return m.Substring(0, c);

                bool InString = false;

                for (int i = 0; i < m.Length; i++)
                {
                    if (m[i] == '"')
                        InString = !InString;

                    if (!InString && m[i] == '\'')
                        return m.Substring(0, i);
                }

                return m;
            }
        }
    }

    public class VB6LogicalRow
    {
        private List<VB6PhysicalRow> mPhysicals;

        public VB6LogicalRow()
        {
            mPhysicals = new List<VB6PhysicalRow>();
        }

        public void Add(string row)
        {
            mPhysicals.Add(new VB6PhysicalRow(row));
        }

        public int Count
        {
            get
            {
                return mPhysicals.Count();
            }
        }

        public bool CanAddRow
        {
            get
            {
                VB6PhysicalRow r = null;
                foreach (var item in mPhysicals)
                    r = item;

                if (r == null)
                    return true;

                return r.IsContinueRow;
            }
        }

        public string GetLastItem(IEnumerable<string> p)
        {
            return p.Reverse().FirstOrDefault();
        }

        public IEnumerable<string> AliveCode()
        {
            var vvv = DeleteComment2().ToList();

            if (vvv.Count >= 1)
                if (GetLastItem(vvv).Trim() == "")
                    vvv.RemoveAt(vvv.Count - 1);

            if (vvv.Count >= 1)
            {
                var a = new VB6PhysicalRow(GetLastItem(vvv));

                if (a.IsContinueRow)
                {
                    var lastindex = vvv.Count - 1;
                    vvv[lastindex] = a.EssenceValue;
                }
            }

            return vvv;
        }

        private IEnumerable<string> DeleteComment2()
        {
            foreach (var vvv in mPhysicals)
            {
                yield return vvv.DeleteComment;
                if (vvv.DeleteComment != vvv.FactValue) yield break;
            }
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
