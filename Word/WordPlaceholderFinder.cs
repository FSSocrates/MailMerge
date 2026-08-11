using Word = Microsoft.Office.Interop.Word;

public static class WordPlaceholderFinder
{
    public static bool Contains(
        Word.Document document)
    {
        foreach (Word.Range story
                 in document.StoryRanges)
        {
            try
            {
                if (ContainsRecursive(story))
                    return true;
            }
            finally
            {
                Com.Release(story);
            }
        }

        return false;
    }

    private static bool ContainsRecursive(
        Word.Range story)
    {
        Word.Find? find = null;

        try
        {
            find = story.Find;

            find.ClearFormatting();
            find.Text = "{{*}}";
            find.Forward = true;
            find.Wrap =
                Word.WdFindWrap.wdFindStop;
            find.MatchWildcards = true;

            if (find.Execute())
                return true;
        }
        finally
        {
            Com.Release(find);
        }

        Word.Range? next = null;

        try
        {
            next = story.NextStoryRange;

            return next != null &&
                   ContainsRecursive(next);
        }
        finally
        {
            Com.Release(next);
        }
    }
}
