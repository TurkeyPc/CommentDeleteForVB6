using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace CommentDeleteForVB6
{
    public class VB6Source
    {
        private readonly string[] source;

        public VB6Source(string path)
        {
            this.source = File.ReadAllLines(path);
        }

        public VB6Source(IEnumerable<string> source)
        {
            this.source = source.ToArray();
        }

        public IEnumerable<string> CommentDeleted
        {
            get
            {
                return PhysicalRows( LogicalRows(source).Select(p => DeleteComment2(p)));
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

        private static IEnumerable<List<string>> LogicalRows(IEnumerable<string> s)
        {
            var v = new List<string>();

            foreach (var sss in s)
            {
                if (sss.Trim() == "")
                {
                    if (v.Count > 0)
                    {
                        yield return v;
                        v = new List<string>();
                    }
                    continue;
                }

                if (sss.Reverse().First() != '_')
                {
                    v.Add(sss);
                    yield return v;
                    v = new List<string>();
                    continue;
                }

                v.Add(sss);
            }

            if (v.Count > 0)
                yield return v;
        }

        private static IEnumerable<string> PhysicalRows(IEnumerable<IEnumerable<string>> s)
        {
            foreach (var ie in s)
                foreach (var item in ie)
                    yield return item;
        }
    }
}
