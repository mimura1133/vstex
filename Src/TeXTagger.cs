using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Tagging;

namespace VsTeXProject
{
    internal class TeXTag : IGlyphTag { }
    class TeXTagger : ITagger<TeXTag>
    {
        private const string _searchText = "todo";

        /// <summary>
        /// This method creates ToDoTag TagSpans over a set of SnapshotSpans.
        /// </summary>
        /// <param name="spans">A set of spans we want to get tags for.</param>
        /// <returns>The list of ToDoTag TagSpans.</returns>
        IEnumerable<ITagSpan<TeXTag>> ITagger<TeXTag>.GetTags(NormalizedSnapshotSpanCollection spans)
        {
            //todo: implement tagging
            foreach (SnapshotSpan curSpan in spans)
            {
                int loc = curSpan.GetText().ToLower().IndexOf(_searchText);
                if (loc > -1)
                {
                    SnapshotSpan todoSpan = new SnapshotSpan(curSpan.Snapshot, new Span(curSpan.Start + loc, _searchText.Length));
                    yield return new TagSpan<TeXTag>(todoSpan, new TeXTag());
                }
            }

        }

        public event EventHandler<SnapshotSpanEventArgs> TagsChanged
        {
            add { }
            remove { }
        }
    }
}
