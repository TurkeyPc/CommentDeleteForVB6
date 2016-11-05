using System.Collections.Generic;
using System.Linq;
using System.IO;


namespace CommentDeleteForVB6
{
    class VB6Source
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
                return null;
            }
        }

    }
}
