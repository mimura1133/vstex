using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Utilities;

namespace VsTeXProject
{

    [Export(typeof(IGlyphFactoryProvider))]
    [Name("TeXClassifierCommentOutFormatGlyph")]
    [Order(Before = "VsTextMarker")]
    [ContentType("vstex")]
    [TagType(typeof(TeXClassifierCommentOutFormatTag))]
    class TeXClassifierCommentOutFormatGlyphFactoryProvider : IGlyphFactoryProvider
    {
        public IGlyphFactory GetGlyphFactory(IWpfTextView view, IWpfTextViewMargin margin)
        {
            return null;
        }
    }

    [Export(typeof(IGlyphFactoryProvider))]
    [Name("TeXClassifierBeginEndFormatGlyph")]
    [Order(Before = "VsTextMarker")]
    [ContentType("vstex")]
    [TagType(typeof(TeXClassifierBeginEndFormatTag))]
    class TeXClassifierBeginEndFormatGlyphFactoryProvider : IGlyphFactoryProvider
    {
        public IGlyphFactory GetGlyphFactory(IWpfTextView view, IWpfTextViewMargin margin)
        {
            return null;
        }
    }

    [Export(typeof(IGlyphFactoryProvider))]
    [Name("TeXClassifierFunctionFormatGlyph")]
    [Order(Before = "VsTextMarker")]
    [ContentType("vstex")]
    [TagType(typeof(TeXClassifierFunctionFormatTag))]
    class TeXClassifierFunctionFormatGlyphFactoryProvider : IGlyphFactoryProvider
    {
        public IGlyphFactory GetGlyphFactory(IWpfTextView view, IWpfTextViewMargin margin)
        {
            return null;
        }
    }

    [Export(typeof(IGlyphFactoryProvider))]
    [Name("TeXClassifierBraceFormatGlyph")]
    [Order(Before = "VsTextMarker")]
    [ContentType("vstex")]
    [TagType(typeof(TeXClassifierBraceFormatTag))]
    class TeXClassifierBraceFormatGlyphFactoryProvider : IGlyphFactoryProvider
    {
        public IGlyphFactory GetGlyphFactory(IWpfTextView view, IWpfTextViewMargin margin)
        {
            return null;
        }
    }
}
