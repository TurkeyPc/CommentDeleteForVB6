using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;

namespace CommentDeleteForVB6
{
    public class VB6LogicalRow
    {
        private List<string> mPhysicalRows;

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

        public IEnumerable<string> AliveCode()
        {
            return DeleteComment2(mPhysicalRows);
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
                return PhysicalRows(LogicalRows(source));
            }
        }

        public IEnumerable<VB6LogicalRow> LogicalRows()
        {
            return LogicalRows(source);
        }

        private static IEnumerable<VB6LogicalRow> LogicalRows(IEnumerable<string> s)
        {
            var v = new VB6LogicalRow();

            foreach (var sss in s)
            {
                if (new string(sss.Reverse().Take(2).Reverse().ToArray()) != " _")
                {
                    v.Add(sss);
                    yield return v;
                    v = new VB6LogicalRow();
                    continue;
                }

                v.Add(sss);
            }

            if (v.Count > 0)
                yield return v;
        }

        private static IEnumerable<string> PhysicalRows(IEnumerable<VB6LogicalRow> s)
        {
            foreach (var logicalrow in s)
                foreach (var physicalrow in logicalrow.AliveCode())
                    yield return physicalrow;
        }
    }
}
