using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Data;
using System.IO;
using Novacode;

namespace WordReplace
{
    internal class Binding {
        private string value;
        private int index;
        private int length;


        public string Value
        {
            get { return this.value; }
        }

        public int Index
        {
            get { return index; }
        }

        public int Length
        {
            get { return length; }
        }

        public Binding(string Value, int Index, int Length)
        {
            value = Value;
            index = Index;
            length = Length;
        }
    }

    class DocTemplate
    {
        private MemoryStream stream;
        private List<List<Binding>> bindings;
        private static char[] TRIM_CHARS = { ' ', '{', '}' };

        public DocTemplate(MemoryStream Stream)
        {
            stream = Stream;
            bindings = FindReplaces(DocX.Load(stream));
        }

        public void CreateDocument(string outPath, BindMap bm)
        {
            var outDoc = DocX.Create(outPath);
            lock (stream)
                outDoc.ApplyTemplate(stream);
            List<Paragraph> ps = CollectParagraphs(outDoc);

            var ParagraphBinds = bindings.Zip(ps, (pts, pr) => new { Bindings = pts, Paragraph = pr });
            foreach (var ParagraphBind in ParagraphBinds)
            {
                foreach (var binding in ParagraphBind.Bindings)
                {

                    var newValue = bm.Get(binding.Value, "undefined");
                    ParagraphBind.Paragraph.InsertText(binding.Index + binding.Length, newValue, false, null);
                    ParagraphBind.Paragraph.RemoveText(binding.Index, binding.Length, false);
                }

            }
            outDoc.Save();
        }

        private List<Paragraph> CollectParagraphs(DocX doc)
        {
            List<Paragraph> ps = new List<Paragraph>();
            Headers headers = doc.Headers;
            List<Header> headerList = new List<Header> { headers.first, headers.even, headers.odd };
            foreach (Header h in headerList)
                if (h != null)
                    ps.AddRange(h.Paragraphs);
            ps.AddRange(doc.Paragraphs);
            Footers footers = doc.Footers;
            List<Footer> footerList = new List<Footer> { footers.first, footers.even, footers.odd };
            foreach (Footer f in footerList)
                if (f != null)
                    ps.AddRange(f.Paragraphs);
            return ps;
        }

        private List<List<Binding>> FindReplaces(DocX doc)
        {
            List<Paragraph> ps = CollectParagraphs(doc);
            return ps.Select(p => FindBindings(p)).ToList();
        }

        private List<Binding> FindBindings(Paragraph par)
        {
            MatchCollection mc = Regex.Matches(par.Text, @"{[^}]+}", RegexOptions.None);
            List<Binding> result = new List<Binding>();
            foreach (Match m in mc.Cast<Match>().Reverse())
                result.Add(new Binding(m.Value.Trim(TRIM_CHARS), m.Index, m.Length));
            return result;
        }
    }
}
