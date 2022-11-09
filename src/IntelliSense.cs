using System.Collections.Generic;
using System.Globalization;
using System.Collections.Immutable;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Core.Imaging;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Adornments;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;

namespace EmojiSense
{
    [Export(typeof(IAsyncCompletionSourceProvider))]
    [ContentType("text")]
    [Name(nameof(IntelliSense))]
    public class IntelliSense : IAsyncCompletionSourceProvider
    {
        public IAsyncCompletionSource GetOrCreate(ITextView textView) =>
            textView.Properties.GetOrCreateSingletonProperty(() => new AsyncCompletionSource());
    }

    public class AsyncCompletionSource : IAsyncCompletionSource
    {
        private static ImmutableArray<CompletionItem> _cache;
        private static readonly ImageElement _icon = new(KnownMonikers.GlyphRight.ToImageId(), "Variable");
        private static readonly Regex _regex = new(@"(?:\s|^):([^:\s]*):?", RegexOptions.Compiled);

        public static CompletionFilter PeopleFilter = new("People", "P", new ImageElement(KnownMonikers.FeedbackSmile.ToImageId()));
        public static CompletionFilter NatureFilter = new("Nature", "N", new ImageElement(KnownMonikers.ASPWebSite.ToImageId()));
        public static CompletionFilter ObjectFilter = new("Objects", "O", new ImageElement(KnownMonikers.ProjectAlerts.ToImageId()));
        public static CompletionFilter PlacesFilter = new("Places", "L", new ImageElement(KnownMonikers.FlagDarkRed.ToImageId()));
        public static CompletionFilter SymbolsFilter = new("Symbols", "S", new ImageElement(KnownMonikers.ConnectArrow.ToImageId()));

        public Task<CompletionContext> GetCompletionContextAsync(IAsyncCompletionSession session, CompletionTrigger trigger, SnapshotPoint triggerLocation, SnapshotSpan applicableToSpan, CancellationToken token)
        {
            if (_cache == null)
            {
                int index = 1;
                var people = Emojis.People.Select(e => CreateCompletionItem(e, index++, PeopleFilter));
                var nature = Emojis.Nature.Select(e => CreateCompletionItem(e, index++, NatureFilter));
                var objects = Emojis.Objects.Select(e => CreateCompletionItem(e, index++, ObjectFilter));
                var places = Emojis.Places.Select(e => CreateCompletionItem(e, index++, PlacesFilter));
                var symbols = Emojis.Symbols.Select(e => CreateCompletionItem(e, index++, SymbolsFilter));
                _cache = people.Union(nature).Union(objects).Union(places).Union(symbols).ToImmutableArray();
            }

            return Task.FromResult(new CompletionContext(_cache));
        }

        private CompletionItem CreateCompletionItem(KeyValuePair<string, string> emojiPair, int index, CompletionFilter compFilter)
        {
            string name = emojiPair.Key;
            string displayName = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(name.Trim(':'));
            string value = emojiPair.Value;
            string order = index.ToString().PadLeft(4, '0');
            ImmutableArray<CompletionFilter> filter = new List<CompletionFilter>() { compFilter }.ToImmutableArray();
            ImmutableArray<ImageElement> icons = ImmutableArray<ImageElement>.Empty;

            return new CompletionItem(displayName, this, _icon, filter, value, value, order, displayName, icons);
        }

        public Task<object> GetDescriptionAsync(IAsyncCompletionSession session, CompletionItem item, CancellationToken token)
        {
            return Task.FromResult<object>(null);
        }

        public CompletionStartData InitializeCompletion(CompletionTrigger trigger, SnapshotPoint triggerLocation, CancellationToken token)
        {
            if (trigger.Character == ':' &&
                (triggerLocation == triggerLocation.Snapshot.Length ||
                (triggerLocation.GetChar() != ':' &&
                !triggerLocation.GetChar().IsWhiteSpaceOrZero())))
            {
                return CompletionStartData.DoesNotParticipateInCompletion;
            }

            if (trigger.Character != ':')
            {
                return CompletionStartData.DoesNotParticipateInCompletion;
            }

            ITextSnapshotLine line = triggerLocation.GetContainingLine();
            string lineText = line.GetText();

            foreach (Match match in _regex.Matches(lineText))
            {
                int offset = match.Value[0].IsWhiteSpaceOrZero() ? 1 : 0;
                int start = match.Index + line.Start + offset;
                int end = start + match.Length - offset;

                if (triggerLocation >= start && triggerLocation <= end)
                {
                    SnapshotSpan span = new(triggerLocation.Snapshot, Span.FromBounds(start, end));
                    return new CompletionStartData(CompletionParticipation.ProvidesItems, span);
                }
            }
            return CompletionStartData.DoesNotParticipateInCompletion;
        }
    }
}
