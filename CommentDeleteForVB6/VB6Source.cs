using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;

namespace CommentDeleteForVB6
{

    public class VB6PhysicalRow
    {
        public readonly string FactValue;

        private const string ContinueString = " _";

        public VB6PhysicalRow(string p)
        {
            FactValue = p;
        }

        public bool IsContinueRow
        {
            get
            {
                if (FactValue.Length < ContinueString.Length) return false;

                var s = new string(FactValue.Reverse().Take(ContinueString.Length).Reverse().ToArray());
                return s == ContinueString;
            }
        }

        public string EssenceValue
        {
            get
            {
                if (DeleteComment != FactValue)
                    return DeleteComment;

                if (IsContinueRow)
                    return CutLast(FactValue, ContinueString.Length);

                return FactValue;
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
                bool InString = false;

                for (int i = 0; i < FactValue.Length; i++)
                {
                    if (FactValue[i] == '"')
                        InString = !InString;

                    if (!InString && FactValue[i] == '\'')
                        return FactValue.Substring(0, i);
                }

                return FactValue;
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
                var r = mPhysicals.Reverse<VB6PhysicalRow>().FirstOrDefault();

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
            var vvv = DeleteComment().ToList();

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

        private IEnumerable<string> DeleteComment()
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
