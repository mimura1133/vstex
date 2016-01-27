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

    #region Tags

    internal class TeXClassifierCommentOutFormatTag : IGlyphTag
    {
    }

    internal class TeXClassifierBeginEndFormatTag : IGlyphTag
    {
    }

    internal class TeXClassifierFunctionFormatTag : IGlyphTag
    {
    }

    internal class TeXClassifierBraceFormatTag : IGlyphTag
    {
    }

    #endregion

    #region Taggers

    internal class TeXClassifierCommentOutFormatTagTagger : ITagger<TeXClassifierCommentOutFormatTag>
    {
        IEnumerable<ITagSpan<TeXClassifierCommentOutFormatTag>> ITagger<TeXClassifierCommentOutFormatTag>.GetTags(NormalizedSnapshotSpanCollection spans)
        {
            foreach (SnapshotSpan curSpan in spans)
            {
                var text = curSpan.GetText();
                for (var pt = 0; pt < curSpan.Length; pt++)
                {
                    if (text[pt] == '\\')
                    {
                        pt++;
                        continue;
                    }

                    if (text[pt] == '%')
                    {
                        yield return
                            new TagSpan<TeXClassifierCommentOutFormatTag>(
                                new SnapshotSpan(curSpan.Snapshot, new Span(curSpan.Start + pt, text.Length - pt)),
                                new TeXClassifierCommentOutFormatTag());

                        break;
                    }
                }
            }
        }

#pragma warning disable 67
        public event EventHandler<SnapshotSpanEventArgs> TagsChanged;
#pragma warning restore 67
    }

    internal class TeXClassifierBeginEndFormatTagTagger : ITagger<TeXClassifierBeginEndFormatTag>
    {
        IEnumerable<ITagSpan<TeXClassifierBeginEndFormatTag>> ITagger<TeXClassifierBeginEndFormatTag>.GetTags(
            NormalizedSnapshotSpanCollection spans)
        {
            foreach (SnapshotSpan curSpan in spans)
            {
                var text = curSpan.GetText().ToLower();
                for (var pt = 0; pt < curSpan.Length; pt++)
                {
                    if(pt > 0 && text[pt - 1] == '\\')
                    {
                        if (text[pt] == 'e' && curSpan.Length - pt > 3 && text[pt + 1] == 'n' && text[pt + 2] == 'd')
                        {
                            yield return
                                new TagSpan<TeXClassifierBeginEndFormatTag>(
                                    new SnapshotSpan(curSpan.Snapshot, new Span(curSpan.Start + pt - 1, 4)),
                                    new TeXClassifierBeginEndFormatTag());

                            break;
                        }
                        if (text[pt] == 'b' && curSpan.Length - pt > 5 && text[pt + 1] == 'e' && text[pt + 2] == 'g' && text[pt + 3] == 'i' && text[pt + 4] == 'n')
                        {
                            yield return
                                new TagSpan<TeXClassifierBeginEndFormatTag>(
                                    new SnapshotSpan(curSpan.Snapshot, new Span(curSpan.Start + pt - 1, 6)),
                                    new TeXClassifierBeginEndFormatTag());

                            break;
                        }
                    }
                }
            }
        }
#pragma warning disable 67
        public event EventHandler<SnapshotSpanEventArgs> TagsChanged;
#pragma warning restore 67

    }

    internal class TeXClassifierFunctionFormatTagTagger : ITagger<TeXClassifierFunctionFormatTag>
    {
        IEnumerable<ITagSpan<TeXClassifierFunctionFormatTag>> ITagger<TeXClassifierFunctionFormatTag>.GetTags(NormalizedSnapshotSpanCollection spans)
        {
            foreach (SnapshotSpan curSpan in spans)
            {
                var text = curSpan.GetText();
                for (var pt = 0; pt < curSpan.Length; pt++)
                {
                    if (pt > 0 && text[pt - 1] == '\\')
                    {
                        var start = pt;
                        for (; pt < curSpan.Length; pt++)
                        {
                            if (text[pt] == ' ' || text[pt] == '{' || text[pt] == '}' || text[pt] == '[' ||
                                text[pt] == ']')
                                break;
                        }
                        yield return
                            new TagSpan<TeXClassifierFunctionFormatTag>(
                                new SnapshotSpan(curSpan.Snapshot, new Span(curSpan.Start + start - 1, pt - start + 1)),
                                new TeXClassifierFunctionFormatTag());
                    }
                }
            }
        }

#pragma warning disable 67
        public event EventHandler<SnapshotSpanEventArgs> TagsChanged;
#pragma warning restore 67
    }

    internal class TeXClassifierBraceFormatTagTagger : ITagger<TeXClassifierBraceFormatTag>
    {
        private HashSet<int> _tagHashSet = new HashSet<int>();

        IEnumerable<ITagSpan<TeXClassifierBraceFormatTag>> ITagger<TeXClassifierBraceFormatTag>.GetTags(NormalizedSnapshotSpanCollection spans)
        {

            foreach (SnapshotSpan curSpan in spans)
            {
                var line = curSpan.Snapshot.GetLineNumberFromPosition(curSpan.Start);
                var text = curSpan.GetText();
                var inBrace = _tagHashSet.Contains(line);

                _tagHashSet.Remove(line + 1);

                for (var pt = 0; pt < curSpan.Length; pt++)
                {
                    if (text[pt] == '\\')
                    {
                        continue;
                    }

                    if (text[pt] == '%')
                        break;

                    if (inBrace || text[pt] == '{' || text[pt] == '[')
                    {
                        var endbrace = false;
                        var end = pt;

                        for (; end < text.Length; end++)
                        {
                            if (text[end] == '\\')
                                continue;
                            if (text[end] == '}' || text[end] == ']')
                            {
                                end++;
                                endbrace = true;
                                break;
                            }
                        }
                        if (endbrace)
                        {
                            this._tagHashSet.Remove(line + 1);
                            inBrace = false;
                        }
                        else
                        {
                            this._tagHashSet.Add(line + 1);
                        }

                        yield return
                            new TagSpan<TeXClassifierBraceFormatTag>(
                                new SnapshotSpan(curSpan.Snapshot, new Span(curSpan.Start + pt, end - pt)),
                                new TeXClassifierBraceFormatTag());

                        pt = end -1;
                    }
                }
            }
        }

#pragma warning disable 67
        public event EventHandler<SnapshotSpanEventArgs> TagsChanged;
#pragma warning restore 67
    }

    #endregion
}
