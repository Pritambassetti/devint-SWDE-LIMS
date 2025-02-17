using Aspose.Words;
using Aspose.Words.Fields;
using System.Collections;

namespace TemplateSpeAsposeCLI.Utils
{
    public static class UtilsWord
    {
        public static int GetParagraphIndex(Document doc, string paragraphName)
        {
            int index = 0;
            foreach (Section section in doc.Sections)
            {
                foreach (Paragraph paragraph in section.Body.Paragraphs)
                {
                    if (paragraph.ListFormat.IsListItem)
                    {
                        // Get the label such as 4.1.1. (finish by a dot)
                        if (paragraph.ListLabel.LabelString == paragraphName)
                            return index;
                    }
                    index++;
                }
            }
            return -1;
        }
        public static Paragraph GetParagraph(Document doc, string paragraphName)
        {
            foreach (Section section in doc.Sections)
            {
                foreach (Paragraph paragraph in section.Body.Paragraphs)
                {
                    if (paragraph.ListFormat.IsListItem)
                    {
                        if (paragraph.ListLabel.LabelString == paragraphName)
                            return paragraph;
                    }
                }
            }
            return null;
        }
        public static string GetParagraphText(Document doc, string paragraphName)
        {
            Paragraph paragraph = GetParagraph(doc, paragraphName);
            if (paragraph != null)
            {
                return paragraph.GetText();
            }
            else return "";
        }
        public static void WriteAfterBookmark(DocumentBuilder builder, string bookmarkName, string text)
        {
            if (builder.MoveToBookmark(bookmarkName, true, false))
            {
                builder.Write(text);
            }
        }
        public static void WriteBeforeBookmark(DocumentBuilder builder, string bookmarkName, string text)
        {
            if (builder.MoveToBookmark(bookmarkName, false, true))
            {
                builder.Write(text);
            }
        }
        public static void ReplaceBookmarkWithText(Document doc, DocumentBuilder builder, string bookmarkName, string text)
        {
            if (builder.MoveToBookmark(bookmarkName, true, true))
            {
                Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];
                bookmark.Text = "";
                bookmark.Remove();
                builder.Write(text);
            }
        }

        public static void UpdateTableOfContents(Document doc)
        {
            doc.UpdateFields();
        }

        /// <summary>
        /// Removes the specified table of contents field from the document.
        /// </summary>
        /// <param name="doc">The document to remove the field from.</param>
        /// <param name="index">The zero-based index of the TOC to remove.</param>
        public static void RemoveTableOfContents(Document doc, int index = 0)
        {
            // Store the FieldStart nodes of TOC fields in the document for quick access.
            ArrayList fieldStarts = new ArrayList();
            // This is a list to store the nodes found inside the specified TOC. They will be removed
            // At the end of this method.
            ArrayList nodeList = new ArrayList();

            foreach (FieldStart start in doc.GetChildNodes(NodeType.FieldStart, true))
            {
                if (start.FieldType == FieldType.FieldTOC)
                {
                    // Add all FieldStarts which are of type FieldTOC.
                    fieldStarts.Add(start);
                }
            }

            // Ensure the TOC specified by the passed index exists.
            if (index > fieldStarts.Count - 1)
                return;

            bool isRemoving = true;
            // Get the FieldStart of the specified TOC.
            Node currentNode = (Node)fieldStarts[index];

            while (isRemoving)
            {
                // It is safer to store these nodes and delete them all at once later.
                nodeList.Add(currentNode);
                currentNode = currentNode.NextPreOrder(doc);

                // Once we encounter a FieldEnd node of type FieldTOC then we know we are at the end
                // Of the current TOC and we can stop here.
                if (currentNode.NodeType == NodeType.FieldEnd)
                {
                    FieldEnd fieldEnd = (FieldEnd)currentNode;
                    if (fieldEnd.FieldType == FieldType.FieldTOC)
                        isRemoving = false;
                }
            }

            // Remove all nodes found in the specified TOC.
            foreach (Node node in nodeList)
            {
                node.Remove();
            }
        }

        public static void ConcatDoc(Document doc, Document doc2, bool preserveHeader = false, bool linkPreviousHeader = false)
        {
            // Remove the headers and footers from each of the sections in the source document.
            if (!preserveHeader)
            {
                foreach (Section section in doc2.Sections)
                {
                    section.ClearHeadersFooters();
                }
                //Unlink the headers and footers in the source document to stop this from continuing the headers and footers
                doc2.FirstSection.HeadersFooters.LinkToPrevious(linkPreviousHeader);
            }
            doc.AppendDocument(doc2, ImportFormatMode.KeepSourceFormatting);
        }
    }
}
