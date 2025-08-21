# Spire.Doc C# Hello World
## Creating a simple Hello World document using Spire.Doc
```csharp
// Create a new Document object
Document document = new Document();

// Add a new Section to the document
Section section = document.AddSection();

// Add a new Paragraph to the section
Paragraph paragraph = section.AddParagraph();

// Add text content to the paragraph
paragraph.AppendText("Hello World!");
```

---

# spire.doc csharp text highlighting
## find and highlight text in document
```csharp
// Find all occurrences of the string "word" in the document and retrieve their TextSelections
TextSelection[] textSelections = document.FindAllString("word", false, true);

// Iterate through each TextSelection
foreach (TextSelection selection in textSelections)
{
    // Get the entire range of the selection and set its CharacterFormat's HighlightColor property to Yellow
    selection.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
}
```

---

# Spire.Doc C# Content Replacement
## Replace content in a Word document with content from another document
```csharp
// Get the first section of the first document
Section section1 = document1.Sections[0];
// Create a regular expression object to search for a pattern
Regex regex = new Regex(@"\[MY_DOCUMENT\]", RegexOptions.None);
// Find all occurrences of the pattern in the first document
TextSelection[] textSections = document1.FindAllPattern(regex);
// Loop through each occurrence of the pattern
foreach (TextSelection seletion in textSections)
{
    // Get the paragraph that contains the pattern
    Paragraph para = seletion.GetAsOneRange().OwnerParagraph;
    // Get the range of text that contains the pattern
    TextRange textRange = seletion.GetAsOneRange();
    // Get the index of the paragraph in the first document's section
    int index = section1.Body.ChildObjects.IndexOf(para);
    // Loop through each section in the second document
    foreach (Section section2 in document2.Sections)
    {
        // Loop through each paragraph in the second section
        foreach (Paragraph paragraph in section2.Paragraphs)
        {
            // Insert the paragraph from the second document into the first document's section
            section1.Body.ChildObjects.Insert(index, paragraph.Clone() as Paragraph);
        }
    }
    // Remove the range of text that contains the pattern from the paragraph
    para.ChildObjects.Remove(textRange);
}
```

---

# Spire.Doc C# Text Replacement
## Replace text in a document using regular expressions
```csharp
// Create a new instance of the Document class
Document doc = new Document();

// Create a regular expression object to search for a pattern
Regex regex = new Regex(@"\#\w+\b");

// Replace all occurrences of the pattern in the document with the string "Spire.Doc"
doc.Replace(regex, "Spire.Doc");
```

---

# Spire.Doc C# Text Replacement
## Replace text with a field in a Word document
```csharp
// Find the first occurrence of the string "summary" in the document, and return its TextSelection object
TextSelection selection = document.FindString("summary", false, true);

// Convert the TextSelection object to a TextRange object
TextRange textRange = selection.GetAsOneRange();

// Get the Paragraph object that contains the TextRange object
Paragraph ownParagraph = textRange.OwnerParagraph;

// Find the index of the TextRange object in the ChildObjects collection of the Paragraph object
int rangeIndex = ownParagraph.ChildObjects.IndexOf(textRange);

// Remove the TextRange object from the ChildObjects collection of the Paragraph object at its index
ownParagraph.ChildObjects.RemoveAt(rangeIndex);

// Create a new list to store cloned objects
List<DocumentObject> tempList = new List<DocumentObject>();

// Loop through the ChildObjects collection of the Paragraph object, starting from the index after the removed TextRange object
for (int i = rangeIndex; i < ownParagraph.ChildObjects.Count; i++)
{
    // Clone the current object in the ChildObjects collection and add it to the tempList
    tempList.Add(ownParagraph.ChildObjects[rangeIndex].Clone());
    // Remove the current object from the ChildObjects collection at its index
    ownParagraph.ChildObjects.RemoveAt(rangeIndex);
}

// Append a field called "MyFieldName" to the end of the Paragraph, with a field type of MergeField
ownParagraph.AppendField("MyFieldName", FieldType.FieldMergeField);

// Loop through each object in the tempList
foreach (DocumentObject obj in tempList)
{
    // Add each object from the tempList back into the ChildObjects collection of the Paragraph
    ownParagraph.ChildObjects.Add(obj);
}
```

---

# Spire.Doc CSharp Text Replacement
## Replace text with table in Word document
```csharp
// Find the first occurrence of the string "Christmas Day, December 25" in the document, and return its TextSelection object
Section section = document.Sections[0];
TextSelection selection = document.FindString("Christmas Day, December 25", true, true);

// Convert the TextSelection object to a TextRange object
TextRange range = selection.GetAsOneRange();
// Get the Paragraph object that contains the TextRange object
Paragraph paragraph = range.OwnerParagraph;
// Get the text body that contains the paragraph
Body body = paragraph.OwnerTextBody;
// Find the index of the TextRange object in the ChildObjects collection of the Paragraph object
int index = body.ChildObjects.IndexOf(paragraph);

// Add a new table and reset the number of rows and columns to 3
Table table = section.AddTable(true);
table.ResetCells(3, 3);

// Remove the TextRange object from the ChildObjects collection of the Paragraph object
body.ChildObjects.Remove(paragraph);

// Insert the table into the ChildObjects collection at the position of the previous TextRange object
body.ChildObjects.Insert(index, table);
```

---

# spire.doc csharp document replacement
## replace text with document content
```csharp
// Create a new Word document object
Document doc = new Document("sourceDocument.docx");

// Create an object of another Word document to be used for replacement
IDocument replaceDoc = new Document("replacementDocument.docx");

// Search for a string in the doc document and replace it with the content of the replaceDoc document
doc.Replace("Document1", replaceDoc, false, true);

// Save the modified document
doc.SaveToFile("output.docx", FileFormat.Docx);

// Dispose of the document object to release resources
doc.Dispose();
```

---

# Spire.Doc C# Find and Replace
## Replace text with HTML content in Word documents
```csharp
// Create a temporary section and add HTML content
Section tempSection = document.AddSection();
Paragraph par = tempSection.AddParagraph();
par.AppendHTML(HTML);

// Store HTML content as replacement objects
List<DocumentObject> replacement = new List<DocumentObject>();
foreach (DocumentObject obj in tempSection.Body.ChildObjects)
{
    DocumentObject docObj = obj as DocumentObject;
    replacement.Add(docObj);
}

// Find all placeholder text occurrences
TextSelection[] selections = document.FindAllString("[#placeholder]", false, true);

// Create and sort location objects for each selection
List<TextRangeLocation> locations = new List<TextRangeLocation>();
foreach (TextSelection selection in selections)
{
    locations.Add(new TextRangeLocation(selection.GetAsOneRange()));
}
locations.Sort();

// Replace each placeholder with HTML content
foreach (TextRangeLocation location in locations)
{
    ReplaceWithHTML(location, replacement);
}

// Remove temporary section
document.Sections.Remove(tempSection);

private static void ReplaceWithHTML(TextRangeLocation location, List<DocumentObject> replacement)
{
    TextRange textRange = location.Text;
    int index = location.Index;
    Paragraph paragraph = location.Owner;
    Body sectionBody = paragraph.OwnerTextBody;
    int paragraphIndex = sectionBody.ChildObjects.IndexOf(paragraph);

    int replacementIndex = -1;
    if (index == 0)
    {
        paragraph.ChildObjects.RemoveAt(0);
        replacementIndex = sectionBody.ChildObjects.IndexOf(paragraph);
    }
    else if (index == paragraph.ChildObjects.Count - 1)
    {
        paragraph.ChildObjects.RemoveAt(index);
        replacementIndex = paragraphIndex + 1;
    }
    else
    {
        Paragraph paragraph1 = (Paragraph)paragraph.Clone();
        while (paragraph.ChildObjects.Count > index)
        {
            paragraph.ChildObjects.RemoveAt(index);
        }
        int i = 0;
        int count = index + 1;
        while (i < count)
        {
            paragraph1.ChildObjects.RemoveAt(0);
            i += 1;
        }
        sectionBody.ChildObjects.Insert(paragraphIndex + 1, paragraph1);
        replacementIndex = paragraphIndex + 1;
    }

    for (int i = 0; i <= replacement.Count - 1; i++)
    {
        sectionBody.ChildObjects.Insert(replacementIndex + i, replacement[i].Clone());
    }
}

public class TextRangeLocation : IComparable<TextRangeLocation>
{
    public TextRangeLocation(TextRange text)
    {
        this.Text = text;
    }

    public TextRange Text
    {
        get { return m_Text; }
        set { m_Text = value; }
    }

    private TextRange m_Text;
    public Paragraph Owner
    {
        get { return this.Text.OwnerParagraph; }
    }

    public int Index
    {
        get { return this.Owner.ChildObjects.IndexOf(this.Text); }
    }

    public int CompareTo(TextRangeLocation other)
    {
        return -(this.Index - other.Index);
    }
}
```

---

# Spire.Doc C# Text Replacement
## Replace text with images in a Word document
```csharp
// Find all occurrences of a specific string in the document
TextSelection[] selections = doc.FindAllString("E-iceblue", true, true);

// Iterate over all occurrences found
foreach (TextSelection selection in selections)
{
    // Create a new DocPicture object and load the image into it
    DocPicture pic = new DocPicture(doc);
    pic.LoadImage(image);

    // Get the text range to be replaced
    TextRange range = selection.GetAsOneRange();

    // Get the position of the text range in its paragraph
    int index = range.OwnerParagraph.ChildObjects.IndexOf(range);

    // Insert the image at the position of the text
    range.OwnerParagraph.ChildObjects.Insert(index, pic);

    // Remove the original text
    range.OwnerParagraph.ChildObjects.Remove(range);
}
```

---

# Spire.Doc C# Text Replacement
## Replace text in a Word document using Spire.Doc library
```csharp
// Create a new instance of the Document class
Document document = new Document();

// Replace all occurrences of the word "word" with the text "ReplacedText"
document.Replace("word", "ReplacedText", false, true);
```

---

# Spire.Doc C# Document Content Extraction
## Extract content between paragraphs from a Word document
```csharp
// This function clones the text between the start and end paragraphs from the source document and adds it to the destination document.
private static void ExtractBetweenParagraphs(Document sourceDocument, Document destinationDocument, int startPara, int endPara)
{
    for (int i = startPara - 1; i < endPara; i++)
    {
        // Clone the paragraph object from the source document.
        DocumentObject doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone();

        // Add the cloned paragraph object to the destination document.
        destinationDocument.Sections[0].Body.ChildObjects.Add(doobj);
    }
}
```

---

# Spire.Doc extract content between paragraph styles
## This code demonstrates how to extract paragraphs between two specified paragraph styles from a Word document
```csharp
// Method to extract paragraphs between two paragraph styles
private static void ExtractBetweenParagraphStyles(Document sourceDocument, Document destinationDocument, string stylename1, string stylename2)
{
    int startindex = 0;
    int endindex = 0;

    // Iterate through sections in the source document
    foreach (Section section in sourceDocument.Sections)
    {
        // Iterate through paragraphs in the section
        foreach (Paragraph paragraph in section.Paragraphs)
        {
            // Find the starting paragraph style
            if (paragraph.StyleName == stylename1)
            {
                startindex = section.Body.Paragraphs.IndexOf(paragraph);
            }

            // Find the ending paragraph style
            if (paragraph.StyleName == stylename2)
            {
                endindex = section.Body.Paragraphs.IndexOf(paragraph);
            }
        }

        // Copy paragraphs between the starting and ending indexes
        for (int i = startindex + 1; i < endindex; i++)
        {
            // Clone the document object
            DocumentObject doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone();

            // Add the cloned object to the destination document
            destinationDocument.Sections[0].Body.ChildObjects.Add(doobj);
        }
    }
}
```

---

# Spire.Doc Paragraph Style Extraction
## Extract paragraphs from a Word document based on their style
```csharp
// Create a new Document object.
Document document = new Document();

// Define the name of the style.
String styleName1 = "Heading1";

// Create a StringBuilder object to store the text with the specified style.
StringBuilder style1Text = new StringBuilder();

// Load the Word document file.
document.LoadFromFile("document.docx");

// Append a line to the StringBuilder indicating the style name.
style1Text.AppendLine("The following is the content of the paragraph with the style name " + styleName1 + ": ");

// Iterate over each section in the document.
foreach (Section section in document.Sections)
{
    // Iterate over each paragraph in the section.
    foreach (Paragraph paragraph in section.Paragraphs)
    {
        // Check if the paragraph has the specified style.
        if (paragraph.StyleName != null && paragraph.StyleName.Equals(styleName1))
        {
            // Append the text of the paragraph to the StringBuilder.
            style1Text.AppendLine(paragraph.Text);
        }
    }
}
```

---

# Spire.Doc C# Bookmark Content Extraction
## Extract content from a bookmark in a Word document
```csharp
// Create a new Document object to represent the source document.
Document sourcedocument = new Document();

// Create a new Document object to represent the destination document.
Document destinationDoc = new Document();

// Add a section to the destination document.
Section section = destinationDoc.AddSection();

// Add a paragraph to the section.
Paragraph paragraph = section.AddParagraph();

// Create a BookmarksNavigator object using the source document.
BookmarksNavigator navigator = new BookmarksNavigator(sourcedocument);

// Move the navigator to the bookmark with the specified name.
navigator.MoveToBookmark("Test", true, true);

// Get the content of the bookmark as a TextBodyPart.
TextBodyPart textBodyPart = navigator.GetBookmarkContent();

// Create a list to store the TextRanges extracted from the bookmark.
List<TextRange> list = new List<TextRange>();

// Iterate over each body item in the TextBodyPart.
foreach (var item in textBodyPart.BodyItems)
{
    // Check if the body item is a Paragraph.
    if (item is Paragraph)
    {
        // Iterate over each child object in the Paragraph.
        foreach (var childObject in (item as Paragraph).ChildObjects)
        {
            // Check if the child object is a TextRange.
            if (childObject is TextRange)
            {
                // Cast the child object to TextRange and add it to the list.
                TextRange range = childObject as TextRange;
                list.Add(range);
            }
        }
    }
}

// Copy the TextRanges from the list to the destination document's paragraph.
for (int m = 0; m < list.Count; m++)
{
    paragraph.Items.Add(list[m].Clone());
}
```

---

# Spire.Doc C# Comment Range Extraction
## Extract content from a comment range in a Word document
```csharp
// Get the first comment from the source document.
Comment comment = sourceDoc.Comments[0];

// Get the paragraph that owns the comment.
Paragraph para = comment.OwnerParagraph;

// Find the index of the CommentMarkStart and CommentMarkEnd within the paragraph's ChildObjects.
int startIndex = para.ChildObjects.IndexOf(comment.CommentMarkStart);
int endIndex = para.ChildObjects.IndexOf(comment.CommentMarkEnd);

// Iterate over the ChildObjects in the paragraph between the start and end indices.
for (int i = startIndex; i <= endIndex; i++)
{
    // Clone the DocumentObject at the current index.
    DocumentObject doobj = para.ChildObjects[i].Clone();

    // Add the cloned DocumentObject to a new paragraph in the destination section.
    destinationSec.AddParagraph().ChildObjects.Add(doobj);
}
```

---

# spire.doc csharp content extraction
## extract content from paragraphs to table
```csharp
// Extracts content by table from the source document to the destination document
private static void ExtractByTable(Document sourceDocument, Document destinationDocument, int startPara, int tableNo)
{
    // Get the specified table from the source document
    Table table = sourceDocument.Sections[0].Tables[tableNo - 1] as Table;

    // Get the index of the table in the source document
    int index = sourceDocument.Sections[0].Body.ChildObjects.IndexOf(table);

    // Copy each child object from the source document to the destination document
    for (int i = startPara - 1; i <= index; i++)
    {
        DocumentObject doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone();
        destinationDocument.Sections[0].Body.ChildObjects.Add(doobj);
    }
}
```

---

# spire.doc csharp form field extraction
## extract document content starting from form field
```csharp
Document sourceDocument = new Document();
Document destinationDoc = new Document();
Section section = destinationDoc.AddSection();
int index = 0;

// Iterate through each form field in the body of the source document's first section.
foreach (FormField field in sourceDocument.Sections[0].Body.FormFields)
{
    // Check if the form field is of type FieldFormTextInput.
    if (field.Type == FieldType.FieldFormTextInput)
    {
        // Get the paragraph that contains the form field.
        Paragraph paragraph = field.OwnerParagraph;

        // Find the index of the paragraph within the child objects of the source document's body.
        index = sourceDocument.Sections[0].Body.ChildObjects.IndexOf(paragraph);

        // Exit the loop after finding the first form text input field.
        break;
    }
}

// Copy three consecutive child objects starting from the found index from the source document's body to the destination document's section.
for (int i = index; i < index + 3; i++)
{
    // Clone the child object at the current index.
    DocumentObject doobj = sourceDocument.Sections[0].Body.ChildObjects[i].Clone();

    // Add the cloned child object to the body of the destination document's section.
    section.Body.ChildObjects.Add(doobj);
}
```

---

# Spire.Doc C# Section Management
## Add and delete sections in a Word document
```csharp
// Create a new instance of the Document class.
Document doc = new Document();

// Call the AddSection method to add a new section to the document.
AddSection(doc);

// Call the DeleteSection method to remove the last section from the document.
DeleteSection(doc);

// Method to add a section to the document.
private void AddSection(Document doc)
{
    doc.AddSection();
}

// Method to delete the last section from the document.
private void DeleteSection(Document doc)
{
    doc.Sections.RemoveAt(doc.Sections.Count - 1);
}
```

---

# Spire.Doc C# Section Cloning
## Clone sections from one document to another
```csharp
// Initialize a variable to store a cloned section.
Section cloneSection = null;

// Iterate through each section in the source document.
foreach (Section section in srcDoc.Sections)
{
    // Clone the current section
    cloneSection = section.Clone();
    
    // Add the cloned section to the destination document
    desDoc.Sections.Add(cloneSection);
}
```

---

# spire.doc csharp section
## clone section content from one section to another
```csharp
// Get the first section from the document and assign it to the sec1 variable.
Section sec1 = doc.Sections[0];

// Get the second section from the document and assign it to the sec2 variable.
Section sec2 = doc.Sections[1];

// Iterate through each DocumentObject in the child objects collection of sec1's Body.
foreach (DocumentObject obj in sec1.Body.ChildObjects)
{
    // Clone the current DocumentObject and add it to the child objects collection of sec2's Body.
    sec2.Body.ChildObjects.Add(obj.Clone());
}
```

---

# Spire.Doc C# Page Setup Modification
## Modify page margins and size for document sections
```csharp
// Iterate through each section in the document
foreach (Section section in doc.Sections)
{
    // Set the page margins of the current section
    section.PageSetup.Margins = new MarginsF(100, 80, 100, 80);

    // Set the page size of the current section to Letter size
    section.PageSetup.PageSize = PageSize.Letter;
}

// Example to modify only one section
Section section0 = doc.Sections[0];
section0.PageSetup.Margins = new MarginsF(100, 80, 100, 80);
section0.PageSetup.FooterDistance = 35.4f;
section0.PageSetup.HeaderDistance = 34.4f;
```

---

# spire.doc csharp remove section content
## remove headers, body and footer content from document sections
```csharp
// Iterate through each section in the document.
foreach (Section section in doc.Sections)
{
    // Clear the child objects in the header of the current section.
    section.HeadersFooters.Header.ChildObjects.Clear();

    // Clear the child objects in the body of the current section.
    section.Body.ChildObjects.Clear();

    // Clear the child objects in the footer of the current section.
    section.HeadersFooters.Footer.ChildObjects.Clear();
}
```

---

# Spire.Doc C# Tab Stops
## Add tab stops to paragraphs in Word document
```csharp
// Create a new document
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Add a paragraph to the section
Paragraph paragraph1 = section.AddParagraph();

// Add a tab stop at position 28 to the paragraph
Tab tab = paragraph1.Format.Tabs.AddTab(28);
tab.Justification = TabJustification.Left;

// Append the text "Washing Machine" with a tab character
paragraph1.AppendText("\tWashing Machine");

// Add another tab stop at position 280 to the paragraph
tab = paragraph1.Format.Tabs.AddTab(280);
tab.Justification = TabJustification.Left;
tab.TabLeader = TabLeader.Dotted;

// Append the text "$650" with a tab character and dotted leader
paragraph1.AppendText("\t$650");

// Add a new paragraph to the section
Paragraph paragraph2 = section.AddParagraph();

// Add a tab stop at position 28 to the second paragraph
tab = paragraph2.Format.Tabs.AddTab(28);
tab.Justification = TabJustification.Left;

// Append the text "Refrigerator" with a tab character
paragraph2.AppendText("\tRefrigerator");

// Add another tab stop at position 280 to the second paragraph
tab = paragraph2.Format.Tabs.AddTab(280);
tab.Justification = TabJustification.Left;
tab.TabLeader = TabLeader.NoLeader;

// Append the text "$800" with a tab character and no leader
paragraph2.AppendText("\t$800");
```

---

# Spire.Doc Paragraph Word Wrap
## Allow Latin text to wrap in the middle of a word
```csharp
// Get the first paragraph in the first section of the document
Paragraph para = document.Sections[0].Paragraphs[0];

// Allow Latin text to wrap in the middle of a word
para.Format.WordWrap = true;
```

---

# Spire.Doc C# Paragraph Copying
## Copy paragraphs between Word documents using Spire.Doc
```csharp
// Get the first section of document1
Section s = document1.Sections[0];

// Get the first paragraph of section s
Paragraph p1 = s.Paragraphs[0];

// Get the second paragraph of section s
Paragraph p2 = s.Paragraphs[1];

// Add a new section to document2
Section s2 = document2.AddSection();

// Clone and add the cloned paragraph (NewPara1) from document1 to s2
Paragraph NewPara1 = (Paragraph)p1.Clone();
s2.Paragraphs.Add(NewPara1);

// Clone and add the cloned paragraph (NewPara2) from document1 to s2
Paragraph NewPara2 = (Paragraph)p2.Clone();
s2.Paragraphs.Add(NewPara2);
```

---

# spire.doc csharp catalogue
## create a catalogue structure with custom list styles
```csharp
// Create a new Document object.
Document document = new Document();

// Add a Section to the document.
Section section = document.AddSection();

// Get the first Paragraph of the Section, or add a new Paragraph if none exists.
Paragraph paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();

// Add a new Paragraph to the Section.
paragraph = section.AddParagraph();

// Set the text content of the Paragraph and apply the Heading1 style.
paragraph.AppendText(BuiltinStyle.Heading1.ToString());
paragraph.ApplyStyle(BuiltinStyle.Heading1);

// Apply a numbered list format to the Paragraph.
paragraph.ListFormat.ApplyNumberedStyle();

// Add another Paragraph to the Section.
paragraph = section.AddParagraph();

// Set the text content of the Paragraph and apply the Heading2 style.
paragraph.AppendText(BuiltinStyle.Heading2.ToString());
paragraph.ApplyStyle(BuiltinStyle.Heading2);

// Create a new ListStyle object with the Numbered list type.
ListStyle listSty2 = new ListStyle(document, ListType.Numbered);

// Iterate over the levels of the ListStyle and customize them.
foreach (ListLevel listLev in listSty2.Levels)
{
    listLev.UsePrevLevelPattern = true;
    listLev.NumberPrefix = "1.";
}

// Set the name of the ListStyle and add it to the document's ListStyles collection.
listSty2.Name = "MyStyle2";
document.ListStyles.Add(listSty2);

// Apply the ListStyle to the current Paragraph.
paragraph.ListFormat.ApplyStyle(listSty2.Name);

// Create another ListStyle object with the Numbered list type.
ListStyle listSty3 = new ListStyle(document, ListType.Numbered);

// Iterate over the levels of the ListStyle and customize them.
foreach (ListLevel listLev in listSty3.Levels)
{
    listLev.UsePrevLevelPattern = true;
    listLev.NumberPrefix = "1.1.";
}

// Set the name of the ListStyle and add it to the document's ListStyles collection.
listSty3.Name = "MyStyle3";
document.ListStyles.Add(listSty3);

// Add four Paragraphs to the Section and apply Heading3 style and ListStyle to each.
for (int i = 0; i < 4; i++)
{
    paragraph = section.AddParagraph();
    paragraph.AppendText(BuiltinStyle.Heading3.ToString());
    paragraph.ApplyStyle(BuiltinStyle.Heading3);
    paragraph.ListFormat.ApplyStyle(listSty3.Name);
}
```

---

# Spire.Doc C# Paragraph Processing
## Get paragraphs by style name from a Word document
```csharp
// Create a StringBuilder object to store the content.
StringBuilder content = new StringBuilder();

// Append a line of text to the StringBuilder.
content.AppendLine("Get paragraphs by style name \"Heading1\": ");

// Iterate over the Sections in the document.
foreach (Section section in document.Sections)
{
    // Iterate over the Paragraphs in each Section.
    foreach (Paragraph paragraph in section.Paragraphs)
    {
        // Check if the Paragraph has the style name "Heading1".
        if (paragraph.StyleName == "Heading1")
        {
            // Append the text of the Paragraph to the StringBuilder.
            content.AppendLine(paragraph.Text);
        }
    }
}
```

---

# spire.doc csharp revisions
## get paragraph and text range revision details from word document
```csharp
// Create a new Document object.
Document document = new Document();

// Iterate over the Sections in the document.
foreach (Section section in document.Sections)
{
    // Iterate over the Paragraphs in each Section.
    foreach (Paragraph paragraph in section.Paragraphs)
    {
        // Check if the Paragraph is a deleted revision.
        if (paragraph.IsDeleteRevision)
        {
            // Append information about the deleted revision.
            builder.AppendLine(string.Format("The section {0} paragraph {1} has been changed (deleted).", document.GetIndex(section), section.GetIndex(paragraph)));
            builder.AppendLine("Author: " + paragraph.DeleteRevision.Author);
            builder.AppendLine("DateTime: " + paragraph.DeleteRevision.DateTime);
            builder.AppendLine("Type: " + paragraph.DeleteRevision.Type);
            builder.AppendLine("");
        }
        // Check if the Paragraph is an inserted revision.
        else if (paragraph.IsInsertRevision)
        {
            // Append information about the inserted revision.
            builder.AppendLine(string.Format("The section {0} paragraph {1} has been changed (inserted).", document.GetIndex(section), section.GetIndex(paragraph)));
            builder.AppendLine("Author: " + paragraph.InsertRevision.Author);
            builder.AppendLine("DateTime: " + paragraph.InsertRevision.DateTime);
            builder.AppendLine("Type: " + paragraph.InsertRevision.Type);
            builder.AppendLine("");
        }
        else
        {
            // Iterate over the child DocumentObjects in the Paragraph.
            foreach (DocumentObject obj in paragraph.ChildObjects)
            {
                // Check if the child DocumentObject is a TextRange.
                if (obj.DocumentObjectType.Equals(DocumentObjectType.TextRange))
                {
                    TextRange textRange = obj as TextRange;
                    {
                        // Check if the TextRange is a deleted revision.
                        if (textRange.IsDeleteRevision)
                        {
                            // Append information about the deleted revision.
                            builder.AppendLine(string.Format("The section {0} paragraph {1} textrange {2} has been changed (deleted).", document.GetIndex(section), section.GetIndex(paragraph), paragraph.GetIndex(textRange)));
                            builder.AppendLine("Author: " + textRange.DeleteRevision.Author);
                            builder.AppendLine("DateTime: " + textRange.DeleteRevision.DateTime);
                            builder.AppendLine("Type: " + textRange.DeleteRevision.Type);
                            builder.AppendLine("Change Text: " + textRange.Text);
                            builder.AppendLine("");
                        }
                        // Check if the TextRange is an inserted revision.
                        else if (textRange.IsInsertRevision)
                        {
                            // Append information about the inserted revision.
                            builder.AppendLine(string.Format("The section {0} paragraph {1} textrange {2} has been changed (inserted).", document.GetIndex(section), section.GetIndex(paragraph), paragraph.GetIndex(textRange)));
                            builder.AppendLine("Author: " + textRange.InsertRevision.Author);
                            builder.AppendLine("DateTime: " + textRange.InsertRevision.DateTime);
                            builder.AppendLine("Type: " + textRange.InsertRevision.Type);
                            builder.AppendLine("Change Text: " + textRange.Text);
                            builder.AppendLine("");
                        }
                    }
                }
            }
        }
    }
}
```

---

# spire.doc csharp paragraph
## hide paragraph text in word document
```csharp
// Get the first Section of the document.
Section sec = document.Sections[0];

// Get the first Paragraph of the Section.
Paragraph para = sec.Paragraphs[0];

// Iterate over the child DocumentObjects in the Paragraph.
foreach (DocumentObject obj in para.ChildObjects)
{
    // Check if the child DocumentObject is a TextRange.
    if (obj is TextRange)
    {
        // Convert the child DocumentObject to a TextRange.
        TextRange range = obj as TextRange;

        // Set the Hidden property of the TextRange's CharacterFormat to true, hiding the text.
        range.CharacterFormat.Hidden = true;
    }
}
```

---

# Spire.Doc CSharp RTF Insertion
## Insert RTF string into a Word document
```csharp
// Create a new Document object.
Document document = new Document();

// Add a new Section to the document.
Section section = document.AddSection();

// Add a new Paragraph to the Section.
Paragraph para = section.AddParagraph();

// Define an RTF string containing formatted text.
String rtfString = @"{\rtf1\ansi\deff0 {\fonttbl {\f0 hakuyoxingshu7000;}}\f0\fs28 Hello, World}";

// Append the RTF string to the Paragraph, preserving the formatting.
para.AppendRTF(rtfString);
```

---

# Spire.Doc C# Paragraph Page Break
## Set page break before a paragraph in Word documents

```csharp
// Get the first section of the document
Section sec = document.Sections[0];

// Get the fifth paragraph of the section
Paragraph para = sec.Paragraphs[4];

// Set page break before the paragraph
para.Format.PageBreakBefore = true;
```

---

# spire.doc csharp paragraph removal
## remove all paragraphs from a Word document
```csharp
// Create a new instance of the Document class.
Document document = new Document();

// Iterate through each section in the document.
foreach (Section section in document.Sections)
{
    // Clear all paragraphs within the section.
    section.Paragraphs.Clear();
}
```

---

# spire.doc csharp remove empty lines
## remove empty paragraphs from word document
```csharp
// Iterate through each section in the document.
foreach (Section section in document.Sections)
{
    // Iterate through the child objects within the body of the section.
    for (int i = 0; i < section.Body.ChildObjects.Count; i++)
    {
        // Check if the child object is of type 'Paragraph'.
        if (section.Body.ChildObjects[i].DocumentObjectType == DocumentObjectType.Paragraph)
        {
            // Check if the text of the paragraph is empty or consists only of whitespace.
            if (String.IsNullOrEmpty((section.Body.ChildObjects[i] as Paragraph).Text.Trim()))
            {
                // Remove the empty paragraph from the child objects collection.
                section.Body.ChildObjects.Remove(section.Body.ChildObjects[i]);

                // Decrement the counter to account for the removed element.
                i--;
            }
        }
    }
}
```

---

# spire.doc csharp paragraph
## remove specific paragraph from document
```csharp
// Remove the paragraph at index 0 from the first section of the document.
document.Sections[0].Paragraphs.RemoveAt(0);
```

---

# spire.doc csharp paragraph formatting
## set before and after spacing for paragraph lines
```csharp
// Access the first section of the document
Section section = doc.Sections[0];

// Access the first paragraph in the section
Paragraph paragraph = section.Paragraphs[0];

// Set the spacing before the paragraph 
paragraph.Format.BeforeSpacingLines = 5f;

// Set the spacing after the paragraph
paragraph.Format.AfterSpacingLines = 15f;
```

---

# Spire.Doc Paragraph Indentation
## Set first line indentation for paragraphs in a document
```csharp
// Create a Paragraph object using the loaded document
Paragraph para = new Paragraph(document);

// Append text to the paragraph and customize its formatting
TextRange textRange1 = para.AppendText("This is an inserted paragraph.");
textRange1.CharacterFormat.TextColor = Color.Blue;
textRange1.CharacterFormat.FontSize = 15;

// Set the first line indent of the paragraph to 2 characters
para.Format.FirstLineIndentChars = 2;

// Alternatively, set the hanging indent as 2 characters
// para.Format.FirstLineIndentChars = -2;

// Reset the first line indent to 0 characters
para.Format.SetFirstLineIndentChars(0);

// Insert the paragraph at index 1 in the first section of the document
document.Sections[0].Paragraphs.Insert(1, para);
```

---

# spire.doc csharp frame position
## Set frame position in Word document
```csharp
//Get a paragraph
Paragraph paragraph = document.Sections[0].Paragraphs[0];

//Set the Frame's position
if (paragraph.Frame.IsFrame)
{
    paragraph.Frame.SetHorizontalPosition(150f);
    paragraph.Frame.SetVerticalPosition(150f);
}
```

---

# Spire.Doc C# Paragraph Indentation
## Set paragraph indentation by character units
```csharp
// Create a new Document object
Document document = new Document();

// Add a section to the document
Section sec = document.AddSection();

// Add a paragraph for the title
Paragraph para = sec.AddParagraph();
para.AppendText("Paragraph Formatting");
para.ApplyStyle(BuiltinStyle.Title);

// Add a paragraph with indent settings
para = sec.AddParagraph();
para.AppendText( "This paragraph is indent as follows: Indent 2 characters on the left and 5 characters on the right.");
para.Format.LeftIndentChars= 2f;
para.Format.RightIndentChars = 5f;
```

---

# Spire.Doc C# Paragraph Shading
## Set paragraph shading and text background color in Word documents
```csharp
// Get the first paragraph of the first section in the document.
Paragraph paragaph = document.Sections[0].Paragraphs[0];

// Set the background color of the paragraph to yellow.
paragaph.Format.BackColor = Color.Yellow;

// Get the third paragraph of the first section in the document.
paragaph = document.Sections[0].Paragraphs[2];

// Find the text "Christmas" within the paragraph, starting from the beginning, case-insensitive.
TextSelection selection = paragaph.Find("Christmas", true, false);

// Get the found text range as a single range.
TextRange range = selection.GetAsOneRange();

// Set the text background color of the range to yellow.
range.CharacterFormat.TextBackgroundColor = Color.Yellow;
```

---

# Spire.Doc C# Paragraph Formatting
## Set SnapToGrid property for a paragraph in a Word document
```csharp
// Create a new instance of the Document class.
Document doc = new Document();

// Add a new section to the document.
Section section = doc.AddSection();

// Set the grid type of the page setup in the section to "LinesOnly".
section.PageSetup.GridType = GridPitchType.LinesOnly;

// Set the number of lines per page in the section to 15.
section.PageSetup.LinesPerPage = 15;

// Add a new paragraph to the section.
Paragraph paragraph = section.AddParagraph();

// Append text to the paragraph.
paragraph.AppendText("With Spire.Doc, you can generate, modify, convert, render and print documents without utilizing Microsoft Word®. But you need MS Word viewer to view the resultant document. ");

// Set the "SnapToGrid" property of the paragraph's format to true.
paragraph.Format.SnapToGrid = true;
```

---

# spire.doc csharp text spacing
## set space between Asian and Latin text in Word document
```csharp
// Get the first paragraph of the first section in the document.
Paragraph para = document.Sections[0].Paragraphs[0];

// Set whether to automatically adjust space between Asian text and Latin text.
para.Format.AutoSpaceDE = false;

// Set whether to automatically adjust space between Asian text and numbers.
para.Format.AutoSpaceDN = true;
```

---

# Spire.Doc Paragraph Spacing
## Set paragraph spacing in Word document
```csharp
// Create a new paragraph object and associate it with the document.
Paragraph para = new Paragraph(document);

// Disable automatic spacing before the paragraph.
para.Format.BeforeAutoSpacing = false;
// Set the amount of spacing before the paragraph to 10 points.
para.Format.BeforeSpacing = 10;
// Disable automatic spacing after the paragraph.
para.Format.AfterAutoSpacing = false;
// Set the amount of spacing after the paragraph to 10 points.
para.Format.AfterSpacing = 10;

// Insert the newly created paragraph at index 1 within the paragraphs collection of the first section in the document.
document.Sections[0].Paragraphs.Insert(1, para);
```

---

# spire.doc csharp text formatting
## apply emphasis mark to text in word document
```csharp
// Find all occurrences of the string "Spire.Doc for .NET" in the document.
TextSelection[] textSelections = document.FindAllString("Spire.Doc for .NET", false, true);

// Iterate through each text selection result.
foreach (TextSelection selection in textSelections)
{
    // Get the found text range as a single range and apply an emphasis mark (dot) to its character format.
    selection.GetAsOneRange().CharacterFormat.EmphasisMark = Emphasis.Dot;
}
```

---

# Spire.Doc C# Text Case Conversion
## Change text case to all caps or small caps in a Word document
```csharp
// Get the second paragraph in the first section of the document.
Paragraph para1 = doc.Sections[0].Paragraphs[1];

// Iterate through each child object within the paragraph.
foreach (DocumentObject obj in para1.ChildObjects)
{
    // Check if the child object is a TextRange.
    if (obj is TextRange)
    {
        // Cast the child object to a TextRange and set the AllCaps property to true,
        // which converts the text to all capital letters.
        textRange = obj as TextRange;
        textRange.CharacterFormat.AllCaps = true;
    }
}

// Get the fourth paragraph in the first section of the document.
Paragraph para2 = doc.Sections[0].Paragraphs[3];

// Iterate through each child object within the paragraph.
foreach (DocumentObject obj in para2.ChildObjects)
{
    // Check if the child object is a TextRange.
    if (obj is TextRange)
    {
        // Cast the child object to a TextRange and set the IsSmallCaps property to true,
        // which converts the text to small capital letters.
        textRange = obj as TextRange;
        textRange.CharacterFormat.IsSmallCaps = true;
    }
}
```

---

# spire.doc csharp barcode
## create barcode in document using specific font
```csharp
// Create a new instance of the Document class.
Document doc = new Document();

// Add a new section to the document and get its first paragraph.
Paragraph p = doc.AddSection().AddParagraph();

// Append the text "H63TWX11072" to the paragraph and obtain the TextRange object.
TextRange txtRang = p.AppendText("H63TWX11072");

// Set the font name for the text range to "C39HrP60DlTt".
txtRang.CharacterFormat.FontName = "C39HrP60DlTt";

// Set the font size for the text range to 80.
txtRang.CharacterFormat.FontSize = 80;

// Set the text color for the text range to SeaGreen.
txtRang.CharacterFormat.TextColor = Color.SeaGreen;
```

---

# Spire.Doc C# Text Measurement
## Get the height and width of text in a document
```csharp
// Get the font used for the found text range.
Font font = selection.GetAsOneRange().CharacterFormat.Font;

// Create a fake image with dimensions of 1x1 pixels.
Image fakeImage = new Bitmap(1, 1);
Graphics graphics = Graphics.FromImage(fakeImage);

// Measure the size (height and width) of the specified text using the font.
SizeF size = graphics.MeasureString(text, font);
```

---

# Spire.Doc C# Get Page Number of Text
## Find text in a document and determine which page each occurrence appears on
```csharp
// Create a new Document object
Document document = new Document();

// Find all occurrences of the string "Spire" in the document
TextSelection[] textSelections = document.FindAllString("Spire", false, false);

// Create a FixedLayoutDocument object using the loaded document
FixedLayoutDocument layoutDoc = new FixedLayoutDocument(document);

// Initialize a counter for matched words
int count = 1;

// Create a StringBuilder to store the result
StringBuilder builder = new StringBuilder();

// Iterate through each TextSelection
foreach (TextSelection selection in textSelections)
{
    // Get the layout entities for the current selection
    foreach (FixedLayoutSpan line in layoutDoc.GetLayoutEntitiesOfNode(selection.GetRanges()[0]))
    {
        // Get the page index where the matched word is located
        int index = line.PageIndex;

        // Append the result to the StringBuilder
        builder.AppendLine("The matched word " + count + " is on page:" + index);

        // Increment the counter
        count++;
    }
}
```

---

# spire.doc csharp text extraction
## extract text from word document
```csharp
// Create a new instance of the Document class and load a Word document.
Document document = new Document("document.docx");

// Extract the text content from the document.
string text = document.GetText();
```

---

# spire.doc csharp text insertion
## insert new text after found string and highlight it
```csharp
// Find all occurrences of the string "Word" within the document.
// Perform a case-insensitive search and include whole word matches.
TextSelection[] selections = doc.FindAllString("Word", true, true);

// Initialize variables.
int index = 0;
TextRange range = new TextRange(doc);

// Iterate through each found text selection.
foreach (TextSelection selection in selections)
{
    // Get the entire range of the selected text.
    range = selection.GetAsOneRange();

    // Create a new TextRange object with the document.
    TextRange newrange = new TextRange(doc);

    // Set the text of the new TextRange to "(New text)".
    newrange.Text = "(New text)";

    // Get the index of the range within its owner paragraph.
    index = range.OwnerParagraph.ChildObjects.IndexOf(range);

    // Insert the new TextRange after the current range in the owner paragraph.
    range.OwnerParagraph.ChildObjects.Insert(index + 1, newrange);
}

// Find all occurrences of the string "New text" within the document.
// Perform a case-insensitive search and include whole word matches.
TextSelection[] text = doc.FindAllString("New text", true, true);

// Iterate through each found text selection.
foreach (TextSelection selection in text)
{
    // Set the highlight color of the text range to Yellow.
    selection.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
}
```

---

# spire.doc csharp symbol insertion
## insert unicode symbols into word document
```csharp
// Create a new instance of the Document class.
Document document = new Document();

// Add a new section to the document.
Section section = document.AddSection();

// Add a new paragraph to the section.
Paragraph paragraph = section.AddParagraph();

// Use a unicode character (U+00C4) to create the symbol Ä and append it to the paragraph.
TextRange tr = paragraph.AppendText('\u00C4'.ToString());

// Set the text color of the symbol Ä to red.
tr.CharacterFormat.TextColor = Color.Red;

// Append the symbol Ë to the paragraph using a unicode character (U+00CB).
paragraph.AppendText('\u00CB'.ToString());
```

---

# spire.doc csharp text encoding
## load text with specific encoding
```csharp
// Create a new instance of the Document class.
Document document = new Document();

// Load the text content from the specified input file using UTF-7 encoding.
document.LoadText(inputFile, Encoding.UTF7);
```

---

# spire.doc csharp superscript subscript
## set superscript and subscript text in word document
```csharp
// Create a new instance of the Document class.
Document document = new Document();

// Add a section to the document.
Section section = document.AddSection();

// Add a paragraph to the section.
Paragraph paragraph = section.AddParagraph();

// Append the text "E = mc" to the paragraph.
paragraph.AppendText("E = mc");

// Append the text "2" as a superscript to the paragraph.
TextRange range1 = paragraph.AppendText("2");
range1.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;

// Insert a line break in the paragraph.
paragraph.AppendBreak(BreakType.LineBreak);

// Append the text "F" to the paragraph.
paragraph.AppendText("F");

// Append the text "n" as a subscript to the paragraph.
TextRange range2 = paragraph.AppendText("n");
range2.CharacterFormat.SubSuperScript = SubSuperScript.SubScript;

// Append the text " = Fn-1 + Fn-2" with specific subscripts to the paragraph.
paragraph.AppendText(" = F");
paragraph.AppendText("n-1").CharacterFormat.SubSuperScript = SubSuperScript.SubScript;
paragraph.AppendText(" + F");
paragraph.AppendText("n-2").CharacterFormat.SubSuperScript = SubSuperScript.SubScript;

// Set the font size to 36 for all TextRange items in the paragraph.
foreach (var i in paragraph.Items)
{
    if (i is TextRange)
    {
        (i as TextRange).CharacterFormat.FontSize = 36;
    }
}
```

---

# spire.doc csharp text direction
## set text direction in document sections and table cells
```csharp
// Create a new instance of the Document class.
Document doc = new Document();

// Add a section to the document.
Section section1 = doc.AddSection();

// Set the text direction of section1 to right-to-left.
section1.TextDirection = TextDirection.RightToLeft;

// Add another section to the document.
Section section2 = doc.AddSection();

// Add a table to section2.
Table table = section2.AddTable();
table.ResetCells(1, 1);

// Access the first cell of the table.
TableCell cell = table.Rows[0].Cells[0];

// Set the height of the first row of the table to 150 points.
table.Rows[0].Height = 150;

// Set the width of the first cell of the table to 10 points.
table.Rows[0].Cells[0].SetCellWidth(10, CellWidthType.Point);

// Set the text direction of the cell to right-to-left rotated.
cell.CellFormat.TextDirection = TextDirection.RightToLeftRotated;
```

---

# Spire.Doc C# Text Splitting
## Add columns to a Word document and show lines between columns
```csharp
// Create a new instance of the Document class.
Document doc = new Document();

// Add a column to the first section of the document with specified widths.
doc.Sections[0].AddColumn(100f, 20f);

// Set the "ColumnsLineBetween" property of the page setup in the first section to true, indicating lines between columns.
doc.Sections[0].PageSetup.ColumnsLineBetween = true;
```

---

# Spire.Doc C# Document Operation
## Accept or reject tracked changes in a Word document
```csharp
// Create a new Document object
Document document = new Document();

// Get the first Section of the document
Section sec = document.Sections[0];

// Get the first Paragraph of the Section
Paragraph para = sec.Paragraphs[0];

// Accept all changes made in the document
para.Document.AcceptChanges();
```

---

# Spire.Doc C# Document Operation
## Add a section from one document to another
```csharp
// Get the second section from the source document.
Section Ssection = SouDoc.Sections[1];

// Clone the second section and add it to the target document.
TarDoc.Sections.Add(Ssection.Clone());
```

---

# Spire.Doc C# Document Variables
## Add variables to a Word document and set up field updates
```csharp
// Create a new document
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Add a paragraph to the section
Paragraph paragraph = section.AddParagraph();

// Append a field with the text "A1" and field type FieldDocVariable to the paragraph
paragraph.AppendField("A1", FieldType.FieldDocVariable);

// Add a variable named "A1" with a value of "12" to the document's variables collection
document.Variables.Add("A1", "12");

// Set the IsUpdateFields property of the document to true, enabling field updates
document.IsUpdateFields = true;
```

---

# spire.doc csharp language dictionary
## alter language dictionary for text in word document
```csharp
// Create a new instance of the Document class.
Document document = new Document();

// Add a section to the document.
Section sec = document.AddSection();

// Add a paragraph to the section.
Paragraph para = sec.AddParagraph();

// Append text "corrige según diccionario en inglés" to the paragraph.
TextRange txtRange = para.AppendText("corrige según diccionario en inglés");

// Set the LocaleIdASCII property of the CharacterFormat for the text range to 10250.
txtRange.CharacterFormat.LocaleIdASCII = 10250;
```

---

# Spire.Doc C# File Format Detection
## Detect and determine the format of a document file
```csharp
// Create a new Document object
Document doc = new Document();

// Load the document from a file
doc.LoadFromFile(filePath);

// Get the detected format type of the document
FileFormat ff = doc.DetectedFormatType;

// Initialize a string to hold the file format information
string fileFormat = "The file format is ";

// Use a switch statement to determine the file format and update the fileFormat string accordingly
switch (ff)
{
    case FileFormat.Doc:
        fileFormat += "Microsoft Word 97-2003 document.";
        break;
    case FileFormat.Dot:
        fileFormat += "Microsoft Word 97-2003 template.";
        break;
    case FileFormat.Docx:
        fileFormat += "Office Open XML WordprocessingML Macro-Free Document.";
        break;
    case FileFormat.Docm:
        fileFormat += "Office Open XML WordprocessingML Macro-Enabled Document.";
        break;
    case FileFormat.Dotx:
        fileFormat += "Office Open XML WordprocessingML Macro-Free Template.";
        break;
    case FileFormat.Dotm:
        fileFormat += "Office Open XML WordprocessingML Macro-Enabled Template.";
        break;
    case FileFormat.Rtf:
        fileFormat += "RTF format.";
        break;
    case FileFormat.WordML:
        fileFormat += "Microsoft Word 2003 WordprocessingML format.";
        break;
    case FileFormat.Html:
        fileFormat += "HTML format.";
        break;
    case FileFormat.WordXml:
        fileFormat += "Microsoft Word XML format for Word 2007-2013.";
        break;
    case FileFormat.Odt:
        fileFormat += "OpenDocument Text.";
        break;
    case FileFormat.Ott:
        fileFormat += "OpenDocument Text Template.";
        break;
    case FileFormat.DocPre97:
        fileFormat += "Microsoft Word 6 or Word 95 format.";
        break;
    default:
        fileFormat += "Unknown format.";
        break;
}

// Dispose of the Document object to release resources
doc.Dispose();
```

---

# Spire.Doc C# Document Cloning
## Clone a Word document using Spire.Doc library
```csharp
// Create a new instance of the Document class.
Document document = new Document();

// Load a document from the specified file path.
document.LoadFromFile("Template_Docx_1.docx");

// Clone the document and assign it to a new Document object.
Document newDoc = document.Clone();

// Save the cloned document to a file with the specified output file name and format.
newDoc.SaveToFile("Result-CloneWordDocument.docx", FileFormat.Docx2013);

// Clean up resources used by the documents.
document.Dispose();
newDoc.Dispose();
```

---

# spire.doc csharp document comparison
## Compare two Word documents and mark differences
```csharp
// Compare the contents of the two documents and mark differences using "E-iceblue" as the author name
doc1.Compare(doc2, "E-iceblue");
```

---

# Spire.Doc document comparison
## Compare two Word documents with specific options
```csharp
// Create CompareOptions object and set IgnoreFormatting property to true
CompareOptions compareOptions = new CompareOptions();
compareOptions.IgnoreFormatting = true;

// Compare the contents of the two documents with specified options and mark differences using "E-iceblue" as the author name and current date and time
doc1.Compare(doc2, "E-iceblue", System.DateTime.Now, compareOptions);
```

---

# Spire.Doc CSharp Document Content Copying
## Copy content from one Word document to another
```csharp
// Create source and destination documents
Document sourceDoc = new Document();
Document destinationDoc = new Document();

// Iterate through each section in the source document
foreach (Section sec in sourceDoc.Sections)
{
    // Iterate through each child object in the body of the section
    foreach (DocumentObject obj in sec.Body.ChildObjects)
    {
        // Clone the child object and add it to the destination document
        destinationDoc.Sections[0].Body.ChildObjects.Add(obj.Clone());
    }
}
```

---

# spire.doc csharp document variables
## count variables in a document
```csharp
// Create a new Document object
Document document = new Document();

// Load a Word document from the specified file path
document.LoadFromFile("document_path");

// Get the number of variables present in the document
int number = document.Variables.Count;

// Release the resources used by the document
document.Dispose();
```

---

# spire.doc csharp word count
## count words and characters in a document
```csharp
// Create a new Document object
Document document = new Document();

// Load the document from the specified file path
document.LoadFromFile("document.docx");

// Get the character count
int charCount = document.BuiltinDocumentProperties.CharCount;
int charCountWithSpace = document.BuiltinDocumentProperties.CharCountWithSpace;

// Get the word count
int wordCount = document.BuiltinDocumentProperties.WordCount;
```

---

# spire.doc csharp database operations
## store, read and delete word documents from database
```csharp
// Implementation of the StoreToDatabase method
public static void StoreToDatabase(String input, OleDbConnection connection)
{
    // Create a Document object from the specified input file
    Document doc = new Document(input);

    // Create a MemoryStream to store the document content
    MemoryStream stream = new MemoryStream();

    // Save the document to the MemoryStream in Docx format
    doc.SaveToStream(stream, FileFormat.Docx);

    // Get the file name from the input path
    string fileName = Path.GetFileName(input);

    // Define the SQL command string to insert the document into the database
    string commandString = "INSERT INTO Documents (FileName, FileContent) VALUES('" + fileName + "', @Doc)";

    // Create an OleDbCommand object with the command string and connection
    OleDbCommand command = new OleDbCommand(commandString, connection);

    // Set the parameter value for the document content using the MemoryStream
    command.Parameters.AddWithValue("Doc", stream.ToArray());

    // Execute the SQL command to store the document in the database
    command.ExecuteNonQuery();
}

// Implementation of the ReadFromDatabase method
public static Document ReadFromDatabase(string fileName, OleDbConnection mConnection)
{
    // Define the SQL command string to select the document from the database
    string commandString = "SELECT * FROM Documents WHERE FileName='" + fileName + "'";

    // Create an OleDbCommand object with the command string and connection
    OleDbCommand command = new OleDbCommand(commandString, mConnection);

    // Create an OleDbDataAdapter object with the command
    OleDbDataAdapter adapter = new OleDbDataAdapter(command);

    // Create a DataTable to store the retrieved data
    DataTable dataTable = new DataTable();

    // Fill the DataTable with the data from the database
    adapter.Fill(dataTable);

    // Check if any record matching the document is found in the DataTable
    if (dataTable.Rows.Count == 0)
        throw new ArgumentException(string.Format("Could not find any record matching the document \"{0}\" in the database.", fileName));

    // Get the file content from the first row of the DataTable
    byte[] buffer = (byte[])dataTable.Rows[0]["FileContent"];

    // Create a MemoryStream from the file content
    MemoryStream newStream = new MemoryStream(buffer);

    // Create a Document object from the MemoryStream
    Document doc = new Document(newStream);

    // Return the Document object
    return doc;
}

// Implementation of the DeleteFromDatabase method
public static void DeleteFromDatabase(string fileName, OleDbConnection mConnection)
{
    // Define the SQL command string to delete the document from the database
    string commandString = "DELETE * FROM Documents WHERE FileName='" + fileName + "'";

    // Create an OleDbCommand object with the command string and connection
    OleDbCommand command = new OleDbCommand(commandString, mConnection);

    // Execute the SQL command to delete the document from the database
    command.ExecuteNonQuery();
}
```

---

# spire.doc csharp document properties
## Set built-in document properties in a Word document
```csharp
// Set the Title property of the document
document.BuiltinDocumentProperties.Title = "Document Demo Document";

// Set the Subject property of the document
document.BuiltinDocumentProperties.Subject = "demo";

// Set the Author property of the document
document.BuiltinDocumentProperties.Author = "James";

// Set the Company property of the document
document.BuiltinDocumentProperties.Company = "e-iceblue";

// Set the Manager property of the document
document.BuiltinDocumentProperties.Manager = "Jakson";

// Set the Category property of the document
document.BuiltinDocumentProperties.Category = "Doc Demos";

// Set the Keywords property of the document
document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo";

// Set the Comments property of the document
document.BuiltinDocumentProperties.Comments = "This document is just a demo.";
```

---

# spire.doc download word from url
## Download a Word document from a URL and save it to a file
```csharp
// Create a new Document object
Document document = new Document();

// Create a new instance of WebClient
WebClient webClient = new WebClient();

// Download the Word file from the specified URL and store it in a MemoryStream
using (MemoryStream ms = new MemoryStream(webClient.DownloadData("http://www.e-iceblue.com/images/test.docx")))
{
    // Load the document from the MemoryStream in Docx format
    document.LoadFromStream(ms, FileFormat.Docx);
}

// Specify the file name for the downloaded result
String result = "Result-DownloadWordFileFromURL.docx";

// Save the downloaded document to the specified file path in Docx2013 format
document.SaveToFile(result, FileFormat.Docx2013);

// Dispose of the Document object to release resources
document.Dispose();
```

---

# Spire.Doc C# Track Changes
## Enable track changes in Word document
```csharp
// Create a new Document object
Document document = new Document();

// Enable tracking changes in the document
document.TrackChanges = true;
```

---

# Spire.Doc C# Document Properties
## Retrieve built-in and custom document properties from a Word document
```csharp
// Assuming document is a loaded Document object

// Retrieve built-in document properties
string title = document.BuiltinDocumentProperties.Title;
string comments = document.BuiltinDocumentProperties.Comments;
string author = document.BuiltinDocumentProperties.Author;
string keywords = document.BuiltinDocumentProperties.Keywords;
string company = document.BuiltinDocumentProperties.Company;

// Iterate through each custom document property
for (int i = 0; i < document.CustomDocumentProperties.Count; i++)
{
    string propertyName = document.CustomDocumentProperties[i].Name;
    object propertyValue = document.CustomDocumentProperties[i].Value;
}
```

---

# spire.doc csharp revisions
## get document revisions and track changes
```csharp
// Create a new Document object
Document document = new Document();

// Load a Word document from a file
document.LoadFromFile("document.docx");

// Create a StringBuilder to store inserted revisions
StringBuilder insertRevision = new StringBuilder();
insertRevision.AppendLine("Insert revisions:");
int index_insertRevision = 0;

// Create a StringBuilder to store deleted revisions
StringBuilder deleteRevision = new StringBuilder();
deleteRevision.AppendLine("Delete revisions:");
int index_deleteRevision = 0;

// Iterate through the sections in the document
foreach (Section sec in document.Sections)
{
    // Iterate through the child objects in the section's body
    foreach (DocumentObject docItem in sec.Body.ChildObjects)
    {
        // Check if the child object is a Paragraph
        if (docItem is Paragraph)
        {
            // Cast the child object to a Paragraph
            Paragraph para = (Paragraph)docItem;
            
            // Check if the paragraph contains an insert revision
            if (para.IsInsertRevision)
            {
                // Increment the insert revision index
                index_insertRevision++;
                insertRevision.AppendLine("Index: " + index_insertRevision);
                
                // Get the InsertRevision object for the paragraph
                EditRevision insRevison = para.InsertRevision;

                // Get the type of the insert revision
                EditRevisionType insType = insRevison.Type;
                insertRevision.AppendLine("Type: " + insType);
                
                // Get the author of the insert revision
                string insAuthor = insRevison.Author;
                insertRevision.AppendLine("Author: " + insAuthor);
            }
            // Check if the paragraph contains a delete revision
            else if (para.IsDeleteRevision)
            {
                // Increment the delete revision index
                index_deleteRevision++;
                deleteRevision.AppendLine("Index: " + index_deleteRevision);
                
                // Get the DeleteRevision object for the paragraph
                EditRevision delRevison = para.DeleteRevision;
                
                // Get the type of the delete revision
                EditRevisionType delType = delRevison.Type;
                deleteRevision.AppendLine("Type: " + delType);
                
                // Get the author of the delete revision
                string delAuthor = delRevison.Author;
                deleteRevision.AppendLine("Author: " + delAuthor);
            }
            
            // Iterate through the child objects in the paragraph
            foreach (DocumentObject obj in para.ChildObjects)
            {
                // Check if the child object is a TextRange
                if (obj is TextRange)
                {
                    // Cast the child object to a TextRange
                    TextRange textRange = (TextRange)obj;
                    
                    // Check if the text range contains an insert revision
                    if (textRange.IsInsertRevision)
                    {
                        // Increment the insert revision index
                        index_insertRevision++;
                        insertRevision.AppendLine("Index: " + index_insertRevision);
                        
                        // Get the InsertRevision object for the text range
                        EditRevision insRevison = textRange.InsertRevision;
                        
                        // Get the type of the insert revision
                        EditRevisionType insType = insRevison.Type;
                        insertRevision.AppendLine("Type: " + insType);
                        
                        // Get the author of the insert revision
                        string insAuthor = insRevison.Author;
                        insertRevision.AppendLine("Author: " + insAuthor);
                    }
                    // Check if the text range contains a delete revision
                    else if (textRange.IsDeleteRevision)
                    {
                        // Increment the delete revision index
                        index_deleteRevision++;
                        deleteRevision.AppendLine("Index: " + index_deleteRevision);
                        
                        // Get the DeleteRevision object for the text range
                        EditRevision delRevison = textRange.DeleteRevision;
                        
                        // Get the type of the delete revision
                        EditRevisionType delType = delRevison.Type;
                        deleteRevision.AppendLine("Type: " + delType);
                        
                        // Get the author of the delete revision
                        string delAuthor = delRevison.Author;
                        deleteRevision.AppendLine("Author: " + delAuthor);
                    }
                }
            }
        }
    }
}
```

---

# spire.doc csharp document variables
## extract and process document variables
```csharp
// Create a new Document object
Document document = new Document();

// Iterate through each key-value pair in the document's Variables collection
foreach (KeyValuePair<string, string> entry in document.Variables)
{
    // Extract the name and value from the current entry
    string name = entry.Key;
    string value = entry.Value;

    // Process the name and value
    // (Here you would typically do something with the variables)
}
```

---

# Spire.Doc Font Table Integration
## Integrate font tables from one document to another and clone sections
```csharp
// Create document instances
Document destDoc = new Document();
Document srcDoc = new Document();

// Integrate the current document font table to the destination document
srcDoc.IntegrateFontTableTo(destDoc);

// Iterate through each section in the source document.
foreach (Section section in srcDoc.Sections)
{
    // Clone each section and add it to the destination document.
    destDoc.Sections.Add(section.Clone());
}
```

---

# spire.doc csharp document format
## keep same format when cloning sections
```csharp
// Create a new instance of the Document class
Document srcDoc = new Document();
Document destDoc = new Document();

// Set the KeepSameFormat property of the source document to true.
srcDoc.KeepSameFormat = true;

// Iterate through each section in the source document.
foreach (Section section in srcDoc.Sections)
{
    // Clone each section and add it to the destination document.
    destDoc.Sections.Add(section.Clone());
}
```

---

# spire.doc csharp document operation
## link headers and footers between document sections
```csharp
// Link the header of the first section in the source document to the previous section's header.
srcDoc.Sections[0].HeadersFooters.Header.LinkToPrevious = true;

// Link the footer of the first section in the source document to the previous section's footer.
srcDoc.Sections[0].HeadersFooters.Footer.LinkToPrevious = true;

// Iterate through each section in the source document.
foreach (Section section in srcDoc.Sections)
{
    // Clone each section and add it to the destination document.
    dstDoc.Sections.Add(section.Clone());
}
```

---

# spire.doc document load and save
## load document from file and save to disk
```csharp
// Create a new instance of the Document class
Document doc = new Document();

// Load the document from the specified input file
doc.LoadFromFile(inputPath);

// Save the document to a file with the specified output file name and file format
doc.SaveToFile outputPath, FileFormat.Docx);

// Dispose of the document object to free up resources
doc.Dispose();
```

---

# spire.doc csharp stream operations
## load document from stream and save to stream
```csharp
// Create a new instance of the Document class by loading the document from the input stream
Document doc = new Document(stream);

// Create a new MemoryStream to store the document
MemoryStream newStream = new MemoryStream();

// Save the document to the new memory stream in RTF format
doc.SaveToStream(newStream, FileFormat.Rtf);

// Reset the position of the memory stream to the beginning
newStream.Position = 0;
```

---

# Spire.Doc Document Merge
## Merge two Word documents by cloning sections from one document to another
```csharp
// Create and load main document
Document document = new Document();
document.LoadFromFile(sourcePath, FileFormat.Doc);

// Create and load document to merge
Document documentMerge = new Document();
documentMerge.LoadFromFile(mergePath, FileFormat.Docx);

// Merge documents by cloning sections
foreach (Section sec in documentMerge.Sections)
{
    document.Sections.Add(sec.Clone());
}

// Save the merged document
document.SaveToFile(outputPath, FileFormat.Docx);

// Release resources
document.Dispose();
documentMerge.Dispose();
```

---

# spire.doc csharp document merge
## merge multiple documents on the same page
```csharp
// Create a new instance of the Document class.
Document document = new Document();
Document destinationDocument = new Document();

// Iterate through each section in the source document.
foreach (Section section in document.Sections)
{
    // Iterate through each child object in the body of the section.
    foreach (DocumentObject obj in section.Body.ChildObjects)
    {
        // Clone each child object and add it to the body of the first section in the destination document.
        destinationDocument.Sections[0].Body.ChildObjects.Add(obj.Clone());
    }
}

// Dispose of the source document and the destination document to release resources.
document.Dispose();
destinationDocument.Dispose();
```

---

# spire.doc csharp modify revision time
## modify revision timestamps in word document
```csharp
// Create a new Document object
Document document = new Document();

// Load a Word document from a file
document.LoadFromFile("ModifyRevisionTime.docx");

// Specify the date string and format
string dateString = "2023/3/1 00:00:00";
string format = "yyyy/M/d HH:mm:ss";

// Parse the date string into a DateTime object using the specified format
DateTime date = DateTime.ParseExact(dateString, format, null);

// Iterate through the sections in the document
foreach (Section sec in document.Sections)
{
    // Iterate through the child objects in the section's body
    foreach (DocumentObject docItem in sec.Body.ChildObjects)
    {
        // Check if the child object is a Paragraph
        if (docItem is Paragraph)
        {
            // Cast the child object to a Paragraph
            Paragraph para = (Paragraph)docItem;
            
            // Check if the paragraph contains an insert revision
            if (para.IsInsertRevision)
            {
                // Get the InsertRevision object for the paragraph
                EditRevision insRevison = para.InsertRevision;
                
                // Set the DateTime property of the insert revision to the specified date
                insRevison.DateTime = date;
            }
            // Check if the paragraph contains a delete revision
            else if (para.IsDeleteRevision)
            {
                // Get the DeleteRevision object for the paragraph
                EditRevision delRevison = para.DeleteRevision;
                
                // Set the DateTime property of the delete revision to the specified date
                delRevison.DateTime = date;
            }
            
            // Iterate through the child objects in the paragraph
            foreach (DocumentObject obj in para.ChildObjects)
            {
                // Check if the child object is a TextRange
                if (obj is TextRange)
                {
                    // Cast the child object to a TextRange
                    TextRange textRange = (TextRange)obj;
                    
                    // Check if the text range contains an insert revision
                    if (textRange.IsInsertRevision)
                    {
                        // Get the InsertRevision object for the text range
                        EditRevision insRevison = textRange.InsertRevision;
                        
                        // Set the DateTime property of the insert revision to the specified date
                        insRevison.DateTime = date;
                    }
                    // Check if the text range contains a delete revision
                    else if (textRange.IsDeleteRevision)
                    {
                        // Get the DeleteRevision object for the text range
                        EditRevision delRevison = textRange.DeleteRevision;
                        
                        // Set the DateTime property of the delete revision to the specified date
                        delRevison.DateTime = date;
                    }
                }
            }
        }
    }
}

// Save the modified document to a new file
document.SaveToFile("ModifyRevisionTime_out.docx", FileFormat.Docx);

// Dispose the Document object
document.Dispose();
```

---

# Spire.Doc Document Theme Preservation
## Clone themes and styles between Word documents
```csharp
// Create a new instance of the Document class.
Document doc = new Document();

// Load the document from a file.
doc.LoadFromFile(inputPath);

// Create another instance of the Document class.
Document newWord = new Document();

// Clone the default style, themes, and compatibility settings from the original document to the new document.
doc.CloneDefaultStyleTo(newWord);
doc.CloneThemesTo(newWord);
doc.CloneCompatibilityTo(newWord);

// Clone the first section from the original document and add it to the new document.
newWord.Sections.Add(doc.Sections[0].Clone());

// Save the new document to a file.
newWord.SaveToFile(outputPath, FileFormat.Docx);

// Dispose of the original document and the new document to release resources.
doc.Dispose();
newWord.Dispose();
```

---

# Spire.Doc Document Object Recursion
## Iterate through all document objects in a Word document
```csharp
// Create a new instance of the Document class
Document document = new Document();

// Create a StringBuilder object to store the output string
StringBuilder builder = new StringBuilder();

// Iterate through each section in the document
foreach (Section section in document.Sections)
{
    // Get the index of the current section
    int sectionIndex = document.GetIndex(section);
    
    // Append a formatted string indicating the section index and its child objects
    builder.AppendLine(string.Format("section index {0} has following ChildObjects", sectionIndex));

    // Iterate through each child object in the section's body
    foreach (DocumentObject obj in section.Body.ChildObjects)
    {
        // Get the index and type of the current child object
        builder.AppendLine(string.Format("Index: {0}, ChildObject Type: {1}", section.Body.GetIndex(obj), obj.DocumentObjectType));
        
        // Check if the child object is a paragraph
        if (obj.DocumentObjectType.Equals(DocumentObjectType.Paragraph))
        {
            // Convert the child object to a Paragraph
            Paragraph paragraph = obj as Paragraph;
            
            // Append a formatted string indicating the paragraph index and its child objects
            builder.AppendLine(string.Format("\tParagraph index {0} has following ChildObjects", section.Body.GetIndex(paragraph)));
            
            // Iterate through each child object in the paragraph
            foreach (DocumentObject obj2 in paragraph.ChildObjects)
            {
                // Append a formatted string indicating the index and type of the child object
                builder.AppendLine(string.Format("\tIndex: {0}, ChildObject Type: {1}", paragraph.GetIndex(obj2), obj2.DocumentObjectType));
            }
        }
    }
    
    // Append a blank line to separate sections
    builder.AppendLine(" ");
}
```

---

# spire.doc csharp variable removal
## remove document variables
```csharp
// Create a new Document object
Document document = new Document();

// Remove the variable named "A1" from the document's Variables collection
document.Variables.Remove("A1");

// Set the IsUpdateFields property of the document to true, enabling field updates
document.IsUpdateFields = true;
```

---

# Spire.Doc C# Retrieve Variables
## Retrieve document variables by index and by name
```csharp
// Create a new Document object
Document document = new Document();

// Get the name of the variable at index 0
string s1 = document.Variables.GetNameByIndex(0);

// Get the value of the variable at index 0
string s2 = document.Variables.GetValueByIndex(0);

// Get the value of the variable with the name "A1"
string s3 = document.Variables["A1"];
```

---

# Spire.Doc C# Section Break
## Make section breaks continuous in a Word document
```csharp
// Iterate through each section in the document
foreach (Section section in doc.Sections)
{
    // Set the break code of each section to NoBreak, which means no section break will be inserted
    section.BreakCode = SectionBreakType.NoBreak;
}
```

---

# spire.doc csharp document revisions
## set author for document revisions
```csharp
// Start track revisions
document.StartTrackRevisions("test");

// Set author for deleted revision
Paragraph para = document.LastParagraph;
para.Text = "";
for (int i = 0; i < para.ChildObjects.Count; i++)
{
    TextRange textRange = para.ChildObjects[i] as TextRange;
    if (textRange.IsDeleteRevision)
    {
        textRange.DeleteRevision.Author = "user1";
    }
}

// Set author for inserted revision
Paragraph paragraph = section.AddParagraph();
TextRange range = paragraph.AppendText("Added text");
range.InsertRevision.Author = "user2";

// Stop track revisions
document.StopTrackRevisions();
```

---

# Spire.Doc C# Document Properties
## Set built-in and custom document properties in Word documents
```csharp
// Set the values of various built-in document properties
document.BuiltinDocumentProperties.Title = "Document Demo Document";
document.BuiltinDocumentProperties.Author = "James";
document.BuiltinDocumentProperties.Company = "e-iceblue";
document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo";
document.BuiltinDocumentProperties.Comments = "This document is just a demo.";

// Get the collection of custom document properties
CustomDocumentProperties custom = document.CustomDocumentProperties;

// Add custom document properties to the collection
custom.Add("e-iceblue", true);
custom.Add("Authorized By", "John Smith");
custom.Add("Authorized Date", DateTime.Today);
```

---

# Spire.Doc C# Word View Modes
## Set Word document view modes and zoom settings
```csharp
// Set the document view type to WebLayout
document.ViewSetup.DocumentViewType = DocumentViewType.WebLayout;

// Set the zoom percentage to 150%
document.ViewSetup.ZoomPercent = 150;

// Set the zoom type to None
document.ViewSetup.ZoomType = ZoomType.None;
```

---

# Spire.Doc C# Document Insertion
## Insert text from one document into another document
```csharp
// Create a new instance of the Document class.
Document doc = new Document();

// Load the document from the specified file path.
doc.LoadFromFile("Template_N5.docx");

// Insert text from another document into the current document, specifying the file path and allowing automatic file format detection.
doc.InsertTextFromFile("Template_N3.docx", FileFormat.Auto);

// Save the modified document to a file with the specified file name and file format as Docx2013.
doc.SaveToFile("SimpleInsertFile_out.docx", FileFormat.Docx2013);

// Dispose of the document to release resources.
doc.Dispose();
```

---

# spire.doc document splitting
## split word document by page breaks
```csharp
// Create a new instance of the Document class to hold the original document
Document original = new Document();

// Create a new instance of the Document class to hold the modified document
Document newWord = new Document();

// Add a section to the new document
Section section = newWord.AddSection();

// Clone the default style, themes, and compatibility settings from the original document to the new document
original.CloneDefaultStyleTo(newWord);
original.CloneThemesTo(newWord);
original.CloneCompatibilityTo(newWord);

// Initialize an index variable to keep track of the split documents
int index = 0;

// Iterate through each section in the original document
foreach (Section sec in original.Sections)
{
    // Iterate through each object in the body of the section
    foreach (DocumentObject obj in sec.Body.ChildObjects)
    {
        // Check if the object is a paragraph
        if (obj is Paragraph)
        {
            // Cast the object as a Paragraph
            Paragraph para = obj as Paragraph;

            // Clone the section properties from the original section to the new section
            sec.CloneSectionPropertiesTo(section);

            // Add the cloned paragraph to the body of the new section
            section.Body.ChildObjects.Add(para.Clone());

            // Iterate through each object in the child objects of the paragraph
            foreach (DocumentObject parobj in para.ChildObjects)
            {
                // Check if the object is a page break
                if (parobj is Break && (parobj as Break).BreakType == BreakType.PageBreak)
                {
                    // Get the index of the page break within the paragraph
                    int i = para.ChildObjects.IndexOf(parobj);

                    // Remove the page break from the last paragraph in the section
                    section.Body.LastParagraph.ChildObjects.RemoveAt(i);

                    // Increment the index for the next split document
                    index++;

                    // Create a new instance of the Document class for the next split document
                    newWord = new Document();

                    // Add a section to the new document
                    section = newWord.AddSection();

                    // Clone the default style, themes, and compatibility settings from the original document to the new document
                    original.CloneDefaultStyleTo(newWord);
                    original.CloneThemesTo(newWord);
                    original.CloneCompatibilityTo(newWord);

                    // Clone the section properties from the original section to the new section
                    sec.CloneSectionPropertiesTo(section);

                    // Add the cloned paragraph to the body of the new section
                    section.Body.ChildObjects.Add(para.Clone());

                    // Check if the first paragraph in the section is empty and remove it if necessary
                    if (section.Paragraphs[0].ChildObjects.Count == 0)
                    {
                        section.Body.ChildObjects.RemoveAt(0);
                    }
                    else
                    {
                        // Remove all objects before the page break in the first paragraph of the section
                        while (i >= 0)
                        {
                            section.Paragraphs[0].ChildObjects.RemoveAt(i);
                            i--;
                        }
                    }
                }
            }
        }
        
        // Check if the object is a table and add it to the body of the section
        if (obj is Table)
        {
            section.Body.ChildObjects.Add(obj.Clone());
        }
    }
}
```

---

# spire.doc csharp document split
## split document by section breaks
```csharp
// Create a new instance of the Document class
Document document = new Document();

// Declare a new Document object
Document newWord;

// Iterate through each section in the document
for (int i = 0; i < document.Sections.Count; i++)
{
    // Create a new instance of the Document class to hold the split section
    newWord = new Document();

    // Clone the section at the current index and add it to the new document
    newWord.Sections.Add(document.Sections[i].Clone());

    // Dispose of the original and new document objects to release resources
    document.Dispose();
    newWord.Dispose();
}
```

---

# Spire.Doc C# Document Splitting
## Split Word document into multiple HTML pages based on headings
```csharp
// Split the document into multiple HTML pages
private static void SplitDocIntoMultipleHtml(String input, string outDirectory)
{
    // Load the document
    Document document = new Document();
    document.LoadFromFile(input);
    
    // Variable to hold the sub-document
    Document subDoc = null; 
    
    // Flag to check if it's the first element in the sub-document
    bool first = true; 
    
    // Index for naming the output HTML files
    int index = 0; 

    // Iterate through sections in the document
    foreach (Section sec in document.Sections)
    {
        // Iterate through elements in the section
        foreach (DocumentObject element in sec.Body.ChildObjects)
        {
            // Check if the element should be in the next document
            if (IsInNextDocument(element))
            {
                if (!first)
                {
                    // Save the previous sub-document as an HTML file
                    subDoc.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.Internal;
                    subDoc.HtmlExportOptions.ImageEmbedded = true;
                    subDoc.SaveToFile(Path.Combine(outDirectory, String.Format("out-{0}.html", index++)), FileFormat.Html);
                    subDoc = null;
                }
                first = false;
            }

            // Create a new sub-document if it doesn't exist
            if (subDoc == null)
            {
                subDoc = new Document();
                subDoc.AddSection();
            }

            // Add the element to the sub-document
            subDoc.Sections[0].Body.ChildObjects.Add(element.Clone());
        }
    }

    // Save the last sub-document as an HTML file, if it exists
    if (subDoc != null)
    {
        subDoc.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.Internal;
        subDoc.HtmlExportOptions.ImageEmbedded = true;
        subDoc.SaveToFile(Path.Combine(outDirectory, String.Format("out-{0}.html", index++)), FileFormat.Html);
    }
}

// Check if the document element should be in the next document
private static bool IsInNextDocument(DocumentObject element)
{
    if (element is Paragraph)
    {
        Paragraph p = element as Paragraph;
        if (p.StyleName == "Heading1")
        {
            return true;
        }
    }
    return false;
}
```

---

# spire.doc csharp revision tracking
## Start and stop tracking revisions in a Word document
```csharp
// Start track revisions
document.StartTrackRevisions("User01", DateTime.Now);

// Get the first paragraph and add content
document.Sections[0].Paragraphs[0].AppendText("User01 add new Text!");

// Delete a paragraph
document.Sections[0].Paragraphs.RemoveAt(2);

// Stop track revisions
document.StopTrackRevisions();
```

---

# spire.doc update document properties
## update last saved date of word document
```csharp
// Set the LastSaveDate property of the built-in document properties to the current local time converted to Greenwich time
document.BuiltinDocumentProperties.LastSaveDate = LocalTimeToGreenwishTime(DateTime.Now);

// Convert local time to Greenwich Mean Time (GMT)
public static DateTime LocalTimeToGreenwishTime(DateTime localTime)
{
    // Get the current local time zone
    TimeZone localTimeZone = TimeZone.CurrentTimeZone;

    // Get the time difference between local time and UTC (Coordinated Universal Time)
    TimeSpan timeSpan = localTimeZone.GetUtcOffset(localTime);

    // Subtract the time difference from the local time to get the Greenwich Mean Time (GMT)
    DateTime greenwishTime = localTime - timeSpan;

    // Return the calculated Greenwich Mean Time (GMT)
    return greenwishTime;
}
```

---

# Spire.Doc C# Gradient Background
## Set gradient background for Word document
```csharp
// Set the background type of the document to gradient
document.Background.Type = BackgroundType.Gradient;

// Get the BackgroundGradient object of the document's background
BackgroundGradient Test = document.Background.Gradient;

// Set the first color of the gradient background to white
Test.Color1 = Color.White;

// Set the second color of the gradient background to light blue
Test.Color2 = Color.LightBlue;

// Set the shading variant of the gradient background to ShadingDown
Test.ShadingVariant = GradientShadingVariant.ShadingDown;

// Set the shading style of the gradient background to Horizontal
Test.ShadingStyle = GradientShadingStyle.Horizontal;
```

---

# Spire.Doc Document Background
## Set image background for a Word document
```csharp
// Create a new Document object
Document document = new Document();

// Set the background type of the document to picture
document.Background.Type = BackgroundType.Picture;

// Set the background picture of the document
document.Background.Picture = Image.FromFile("Background.png");
```

---

# Spire.Doc C# Page Setup
## Add gutter to document section
```csharp
// Create a new Document object
Document document = new Document();

// Get the first section of the document
Section section = document.Sections[0];

// Set the gutter size of the section to 100f (floating-point value)
section.PageSetup.Gutter = 100f;
```

---

# Spire.Doc C# Line Numbering
## Add line numbers to Word document sections
```csharp
// Create a new instance of the Document class
Document document = new Document();

// Set the start value for line numbering in the first section of the document
document.Sections[0].PageSetup.LineNumberingStartValue = 1;

// Set the interval between line numbers in the first section of the document
document.Sections[0].PageSetup.LineNumberingStep = 6;

// Set the distance between line numbers and the main text in the first section of the document
document.Sections[0].PageSetup.LineNumberingDistanceFromText = 40f;

// Set the line numbering restart mode to continuous in the first section of the document
document.Sections[0].PageSetup.LineNumberingRestartMode = LineNumberingRestartMode.Continuous;
```

---

# Spire.Doc C# Page Borders
## Add borders to document pages
```csharp
// Create a new instance of the Document class
Document document = new Document();

// Get the first section of the document
Section section = document.Sections[0];

// Set the border type for the page setup of the section to DoubleWave
section.PageSetup.Borders.BorderType = Spire.Doc.Documents.BorderStyle.DoubleWave;

// Set the color of the borders to LightSeaGreen
section.PageSetup.Borders.Color = Color.LightSeaGreen;

// Set the left spacing for the borders of the page setup
section.PageSetup.Borders.Left.Space = 50;

// Set the right spacing for the borders of the page setup
section.PageSetup.Borders.Right.Space = 50;
```

---

# Spire.Doc C# Page Numbering
## Add page numbers in document sections
```csharp
// Iterate through the first three sections of the document
for (int i = 0; i < 3; i++)
{
    // Get the footer of the current section
    HeaderFooter footer = document.Sections[i].HeadersFooters.Footer;

    // Add a paragraph to the footer
    Paragraph footerParagraph = footer.AddParagraph();

    // Append a page number field to the footer paragraph
    footerParagraph.AppendField("page number", FieldType.FieldPage);

    // Append " of " text to the footer paragraph
    footerParagraph.AppendText(" of ");

    // Append a section pages field to the footer paragraph
    footerParagraph.AppendField("number of pages", FieldType.FieldSectionPages);

    // Set the horizontal alignment of the footer paragraph to right
    footerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

    // If it's the last iteration, exit the loop; otherwise, set up page numbering for the next section
    if (i == 2)
        break;
    else
    {
        // Restart page numbering for the next section
        document.Sections[i + 1].PageSetup.RestartPageNumbering = true;

        // Set the starting page number for the next section to 1
        document.Sections[i + 1].PageSetup.PageStartingNumber = 1;
    }
}
```

---

# Spire.Doc Page Setup
## Demonstrate how to set different page setup for document sections
```csharp
// Create a new instance of the Document class and load a Word document
Document doc = new Document();

// Get the second section of the document
Section SectionTwo = doc.Sections[1];

// Set the page orientation of the second section to Landscape
SectionTwo.PageSetup.Orientation = PageOrientation.Landscape;

// Uncomment the following line to set a custom page size for the second section
// SectionTwo.PageSetup.PageSize = new SizeF(800, 800);
```

---

# spire.doc csharp section break
## insert section break in word document
```csharp
// Create a new instance of the Document class
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Set page settings for the section
section.PageSetup.PageSize = PageSize.A4;
section.PageSetup.Margins.Top = 72f;
section.PageSetup.Margins.Bottom = 72f;
section.PageSetup.Margins.Left = 89.85f;
section.PageSetup.Margins.Right = 89.85f;

// Add another section to the document
section = document.AddSection();

// Insert a section break at the beginning of the section
section.AddParagraph().InsertSectionBreak(SectionBreakType.NewPage);
```

---

# Spire.Doc C# Page Break Insertion
## Insert page breaks after specific text in a Word document
```csharp
// Find all occurrences of the word "technology" in the document
TextSelection[] selections = document.FindAllString("technology", true, true);

// Iterate through each found text selection
foreach (TextSelection ts in selections)
{
    // Get the range of the text selection as one continuous range
    TextRange range = ts.GetAsOneRange();

    // Get the paragraph that contains the text range
    Paragraph paragraph = range.OwnerParagraph;

    // Get the index of the text range within the paragraph's child objects
    int index = paragraph.ChildObjects.IndexOf(range);

    // Insert a page break after the text range by creating a Break object with BreakType.PageBreak
    Break pageBreak = new Break(document, BreakType.PageBreak);

    // Insert the page break at the next index position in the paragraph's child objects
    paragraph.ChildObjects.Insert(index + 1, pageBreak);
}
```

---

# spire.doc csharp page break
## insert page break in word document
```csharp
// Create a new instance of the Document class
Document document = new Document();

// Get the first section of the document, third paragraph, and append a page break to it
document.Sections[0].Paragraphs[3].AppendBreak(BreakType.PageBreak);
```

---

# Spire.Doc C# Section Break
## Insert section break in Word document
```csharp
// Create a new Document object.
Document document = new Document();

// Insert a section break at a specific position in the document.
// There are five section break options: EvenPage, NewColumn, NewPage, NoBreak, OddPage.
document.Sections[0].Paragraphs[1].InsertSectionBreak(SectionBreakType.NoBreak);
```

---

# Spire.Doc Page Setup
## Configure page settings, headers, footers, and tables in a Word document
```csharp
// Create a new Document object.
Document document = new Document();

// Add a new section to the document.
Section section = document.AddSection();

// Set the page size of the section to A4.
section.PageSetup.PageSize = PageSize.A4;

// Set the top margin of the section to 72 points (1 inch).
section.PageSetup.Margins.Top = 72f;

// Set the bottom margin of the section to 72 points (1 inch).
section.PageSetup.Margins.Bottom = 72f;

// Set the left margin of the section to 89.85 points (approximately 1.27 cm).
section.PageSetup.Margins.Left = 89.85f;

// Set the right margin of the section to 89.85 points (approximately 1.27 cm).
section.PageSetup.Margins.Right = 89.85f;

// Insert headers and footers in the section.
HeaderFooter header = section.HeadersFooters.Header;
HeaderFooter footer = section.HeadersFooters.Footer;

// Add a paragraph to the header and insert text.
Paragraph headerParagraph = header.AddParagraph();
TextRange text = headerParagraph.AppendText("Demo of Spire.Doc");
text.CharacterFormat.FontName = "Arial";
text.CharacterFormat.FontSize = 10;
text.CharacterFormat.Italic = true;
headerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
headerParagraph.Format.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single;
headerParagraph.Format.Borders.Bottom.Space = 0.05F;

// Add a paragraph to the footer and insert fields for page numbering.
Paragraph footerParagraph = footer.AddParagraph();
footerParagraph.AppendField("page number", FieldType.FieldPage);
footerParagraph.AppendText(" of ");
footerParagraph.AppendField("number of pages", FieldType.FieldNumPages);
footerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
footerParagraph.Format.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;
footerParagraph.Format.Borders.Top.Space = 0.05F;

// Add a table to the section.
String[] tableHeader = { "Name", "Capital", "Continent", "Area", "Population" };
Spire.Doc.Table table = section.AddTable(true);

// Set the number of rows and columns in the table.
table.ResetCells(1, tableHeader.Length); // Only header row

// First Row (Table Header)
TableRow row = table.Rows[0];
row.IsHeader = true;
row.Height = 20;
row.HeightType = TableRowHeightType.Exactly;
for (int i = 0; i < row.Cells.Count; i++)
{
    row.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.Gray;
}

// Populate the header cells with text and formatting.
for (int i = 0; i < tableHeader.Length; i++)
{
    row.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
    Paragraph p = row.Cells[i].AddParagraph();
    p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
    TextRange txtRange = p.AppendText(tableHeader[i]);
    txtRange.CharacterFormat.Bold = true;
}
```

---

# spire.doc csharp page breaks
## remove page breaks from word document
```csharp
// Iterate through paragraphs in the first section of the document.
for (int j = 0; j < document.Sections[0].Paragraphs.Count; j++)
{
    // Get a reference to the current paragraph.
    Paragraph p = document.Sections[0].Paragraphs[j];

    // Iterate through child objects (elements) within the paragraph.
    for (int i = 0; i < p.ChildObjects.Count; i++)
    {
        // Get a reference to the current child object.
        DocumentObject obj = p.ChildObjects[i];

        // Check if the child object is a Break.
        if (obj.DocumentObjectType == DocumentObjectType.Break)
        {
            // Remove the Break from the paragraph's child objects.
            Break b = obj as Break;
            p.ChildObjects.Remove(b);
        }
    }
}
```

---

# spire.doc csharp page numbering
## reset page numbers in word document sections
```csharp
// Copy sections from document2 to document1.
foreach (Section sec in document2.Sections)
{
    document1.Sections.Add(sec.Clone());
}

// Copy sections from document3 to document1.
foreach (Section sec in document3.Sections)
{
    document1.Sections.Add(sec.Clone());
}

// Modify field types in footer sections of document1.
foreach (Section sec in document1.Sections)
{
    foreach (DocumentObject obj in sec.HeadersFooters.Footer.ChildObjects)
    {
        if (obj.DocumentObjectType == DocumentObjectType.StructureDocumentTag)
        {
            DocumentObject para = obj.ChildObjects[0];
            foreach (DocumentObject item in para.ChildObjects)
            {
                if (item.DocumentObjectType == DocumentObjectType.Field)
                {
                    if ((item as Field).Type == FieldType.FieldNumPages)
                    {
                        // Change the field type to FieldSectionPages.
                        (item as Field).Type = FieldType.FieldSectionPages;
                    }
                }
            }
        }
    }
}

// Reset page numbering for specific sections in document1.
document1.Sections[1].PageSetup.RestartPageNumbering = true;
document1.Sections[1].PageSetup.PageStartingNumber = 1;
document1.Sections[2].PageSetup.RestartPageNumbering = true;
document1.Sections[2].PageSetup.PageStartingNumber = 1;
```

---

# spire.doc csharp page setup
## set gutter position in document section
```csharp
// Get the first section of the document.
Section section = document.Sections[0];

// Set the top gutter option to true for the section's page setup.
section.PageSetup.IsTopGutter = true;

// Set the width of the gutter in points (100f).
section.PageSetup.Gutter = 100f;
```

---

# Spire.Doc C# EPUB Conversion
## Add cover image to EPUB document
```csharp
// Create a new instance of the Document class.
Document doc = new Document();

// Create a new DocPicture object with the document as its owner.
DocPicture picture = new DocPicture(doc);

// Specify the output file name for the EPUB file.
string result = "AddCoverImage.epub";

// Save the document as an EPUB file, including the cover image.
doc.SaveToEpub(result, picture);

// Dispose the document object to release resources.
doc.Dispose();
```

---

# Spire.Doc CSharp Document Conversion
## Convert document to byte array and back to document
```csharp
// Create a new instance of the Document class.
Document doc = new Document();

// Load the document from the specified input file.
doc.LoadFromFile(input);

// Create a new MemoryStream to store the document content.
MemoryStream outStream = new MemoryStream();

// Save the document to the MemoryStream in Docx format.
doc.SaveToStream(outStream, FileFormat.Docx);

// Convert the content of the MemoryStream to a byte array.
byte[] docBytes = outStream.ToArray();

// The bytes are now ready to be stored/transmitted.

// Create a new MemoryStream from the byte array.
MemoryStream inStream = new MemoryStream(docBytes);

// Create a new Document object by loading from the MemoryStream.
Document newDoc = new Document(inStream);

// Dispose the existing document object.
doc.Dispose();
```

---

# Spire.Doc Document Object to Image Conversion
## Convert various document objects (paragraphs, tables, rows, cells, shapes) to images
```csharp
private Image ConvertParagraphToImage(Paragraph obj)
{
    //Create a new document
    Document doc = new Document();

    //Add a new section
    Section section = doc.AddSection();

    //Add a deep clone of the paragraph to the section
    section.Body.ChildObjects.Add(obj.Clone());

    //Save the image
    Image image = doc.SaveToImages(0, ImageType.Bitmap);

    //Close the document
    doc.Close();
    return CutImageWhitePart(image as Bitmap, 1);
}

private Image ConvertTableToImage(Table obj)
{
    //Create a new document
    Document doc = new Document();

    //Add a section to the document
    Section section = doc.AddSection();

    //Add a deep clone of the table to the section
    section.Body.ChildObjects.Add(obj.Clone());

    //Save the image
    Image image = doc.SaveToImages(0, ImageType.Bitmap);

    //Close the document
    doc.Close();
    return CutImageWhitePart(image as Bitmap, 1);
}

private Image ConvertTableRowToImage(TableRow obj)
{
    //Create a new document
    Document doc = new Document();

    //Add a section to the document
    Section section = doc.AddSection();

    //Add a table to the section
    Table table = section.AddTable();

    //Add a deep clone of the row to the table
    table.Rows.Add(obj.Clone());

    //Save the image
    Image image = doc.SaveToImages(0, ImageType.Bitmap);
    doc.Close();
    return CutImageWhitePart(image as Bitmap, 1);
}

private Image ConvertTableCellToImage(TableCell obj)
{
    // Create a new document
    Document doc = new Document();

    //Add a section to the document
    Section section = doc.AddSection();

    //Add a table to the section
    Table table = section.AddTable();

    //Add a new row to the table and add a deep clone of the cell to it
    table.AddRow().Cells.Add(obj.Clone());

    //Save the image
    Image image = doc.SaveToImages(0, ImageType.Bitmap);
    doc.Close();
    return CutImageWhitePart(image as Bitmap, 1);
}

private Image ConvertShapeToImage(ShapeObject obj)
{
    //Create a new document
    Document doc = new Document();

    //Add a section to the document
    Section section = doc.AddSection();

    // Add a paragraph to the section and add a deep clone of the shape object to it
    section.AddParagraph().ChildObjects.Add(obj.Clone());

    //Create a MemoryStream
    MemoryStream ms = new MemoryStream();

    //Save the document to stream
    doc.SaveToStream(ms, FileFormat.Docx);

    //Load a document from stream
    doc.LoadFromStream(ms, FileFormat.Docx);

    //Save to image
    Image image = doc.SaveToImages(0, ImageType.Bitmap);

    //Close the document and stream
    ms.Close();
    doc.Close();
    return CutImageWhitePart(image as Bitmap, 1);
}

public Image CutImageWhitePart(Bitmap bmp, int WhiteBarRate)
{
    int top = 0, left = 0;
    int right = bmp.Width, bottom = bmp.Height;
    Color white = Color.White;

    for (int i = 0; i < bmp.Height; i++)
    {
        bool find = false;
        for (int j = 0; j < bmp.Width; j++)
        {
            Color c = bmp.GetPixel(j, i);
            if (IsWhite(c))
            {
                top = i;
                find = true;
                break;
            }
        }
        if (find) break;
    }

    for (int i = 0; i < bmp.Width; i++)
    {
        bool find = false;
        for (int j = top; j < bmp.Height; j++)
        {
            Color c = bmp.GetPixel(i, j);
            if (IsWhite(c))
            {
                left = i;
                find = true;
                break;
            }
        }
        if (find) break; ;
    }

    for (int i = bmp.Height - 1; i >= 0; i--)
    {
        bool find = false;
        for (int j = left; j < bmp.Width; j++)
        {
            Color c = bmp.GetPixel(j, i);
            if (IsWhite(c))
            {
                bottom = i;
                find = true;
                break;
            }
        }
        if (find) break;
    }

    for (int i = bmp.Width - 1; i >= 0; i--)
    {
        bool find = false;
        for (int j = 0; j <= bottom; j++)
        {
            Color c = bmp.GetPixel(i, j);
            if (IsWhite(c))
            {
                right = i;
                find = true;
                break;
            }
        }
        if (find) break;
    }
    int iWidth = right - left;
    int iHeight = bottom - left;
    int blockWidth = Convert.ToInt32(iWidth * WhiteBarRate / 100);
    bmp = Cut(bmp, left - blockWidth, top - blockWidth, right - left + 2 * blockWidth, bottom - top + 2 * blockWidth);

    return bmp;
}

public Bitmap Cut(Bitmap b, int StartX, int StartY, int iWidth, int iHeight)
{
    if (b == null)
    {
        return null;
    }
    int w = b.Width;
    int h = b.Height;
    if (StartX >= w || StartY >= h)
    {
        return null;
    }
    if (StartX + iWidth > w)
    {
        iWidth = w - StartX;
    }
    if (StartY + iHeight > h)
    {
        iHeight = h - StartY;
    }
    try
    {
        Bitmap bmpOut = new Bitmap(iWidth, iHeight, PixelFormat.Format24bppRgb);
        Graphics g = Graphics.FromImage(bmpOut);
        g.DrawImage(b, new Rectangle(0, 0, iWidth, iHeight), new Rectangle(StartX, StartY, iWidth, iHeight), GraphicsUnit.Pixel);
        g.Dispose();
        return bmpOut;
    }
    catch
    {
        return null;
    }
}

public bool IsWhite(Color c)
{
    if (c.R < 245 || c.G < 245 || c.B < 245)
        return true;
    else return false;
}
```

---

# spire.doc csharp disable hyperlinks
## disable hyperlinks when converting document to pdf
```csharp
// Create a ToPdfParameterList object to specify conversion parameters for PDF export
ToPdfParameterList pdf = new ToPdfParameterList();

// Set DisableLink to true to remove the hyperlink effect for the result PDF page
pdf.DisableLink = true;

// Save the document to PDF format with the specified parameters
document.SaveToFile(result, pdf);
```

---

# spire.doc csharp pdf conversion
## embed all fonts in pdf when converting word document
```csharp
// Create a new instance of the Document class.
Document document = new Document();

// Load a Word document from the specified file path.
document.LoadFromFile("document.docx");

// Create a new instance of the ToPdfParameterList class.
ToPdfParameterList ppl = new ToPdfParameterList();

// Set the IsEmbeddedAllFonts property to true, which embeds all fonts in the resulting PDF.
ppl.IsEmbeddedAllFonts = true;

// Save the document as a PDF file with the specified parameters.
document.SaveToFile("output.pdf", ppl);

// Dispose of the document to free up resources.
document.Dispose();
```

---

# Spire.Doc CSharp Font Embedding
## Embed non-installed fonts when converting Word to PDF
```csharp
// Create a new instance of the Document class.
Document document = new Document();

// Load a Word document from the specified file path.
document.LoadFromFile("input.docx");

// Create a new instance of the ToPdfParameterList class.
ToPdfParameterList parms = new ToPdfParameterList();

// Create a list to hold the paths of private fonts to be embedded.
List<PrivateFontPath> fonts = new List<PrivateFontPath>();

// Add a new PrivateFontPath object to the list, specifying the font name and its file path.
fonts.Add(new PrivateFontPath("PT Serif Caption", "PT_Serif-Caption-Web-Regular.ttf"));

// Set the PrivateFontPaths property of the parameter list to the created list of fonts.
parms.PrivateFontPaths = fonts;

// Save the document as a PDF file with the specified parameters.
document.SaveToFile("EmbedNoninstalledFonts.pdf", parms);

// Dispose of the document to free up resources.
document.Dispose();
```

---

# spire.doc csharp html to pdf conversion
## handle url loading during html to pdf conversion
```csharp
// Create a new instance of the Document class
Document document = new Document();

// Subscribe to the HtmlUrlLoadEvent to handle external resource loading.
document.HtmlUrlLoadEvent += MyDownloadEvent;

// Load an HTML file into the document. The file path and validation type are specified.
document.LoadFromFile("Template_HtmlFile3.html", FileFormat.Html, XHTMLValidationType.None);

// Save the loaded HTML content as a PDF file.
document.SaveToFile("HtmlFileToPDF.pdf", FileFormat.PDF);

private static void MyDownloadEvent(object sender, Document.HtmlUrlLoadEventArgs args)
{
    // Use WebClient to download external resources (e.g., images, CSS files) from URLs referenced in the HTML.
    using (WebClient webClient = new WebClient())
    {
        // Use the default credentials for authentication.
        webClient.Credentials = CredentialCache.DefaultCredentials;

        // Set a custom user-agent header to mimic a web browser during resource download.
        webClient.Headers.Set("user-agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:109.0) Gecko/20100101 Firefox/115.0");

        // Configure supported security protocols for SSL/TLS connections.
        // This ensures compatibility with different server configurations.
        // SystemDefault = 0, Ssl3 = 48, Tls = 192, Tls11 = 768, Tls12 = 3072, Tls13 = 12288
        ServicePointManager.SecurityProtocol = (SecurityProtocolType)0 |
                                               (SecurityProtocolType)12288 |  // Tls13
                                               (SecurityProtocolType)3072 |  // Tls12
                                               (SecurityProtocolType)768 |   // Tls11
                                               (SecurityProtocolType)192 |   // Tls
                                               (SecurityProtocolType)48;     // Ssl3

        // Download the resource data from the provided URL.
        byte[] webData = webClient.DownloadData(args.Url);

        // Set the downloaded data into the event arguments for further processing.
        args.DataBytes = webData;
    }
}
```

---

# spire.doc csharp html to word conversion
## convert HTML file to Word document
```csharp
// Create a new instance of the Document class.
Document document = new Document();

// Load an HTML file into the document object, with XHTML validation disabled.
document.LoadFromFile(@"..\..\..\..\..\..\..\Data\InputHtmlFile.html", FileFormat.Html, XHTMLValidationType.None);

// Save the document as a DOCX file named "HtmlFileToWord.docx".
document.SaveToFile("HtmlFileToWord.docx", FileFormat.Docx);

// Dispose the document object to release resources.
document.Dispose();
```

---

# Spire.Doc HTML to Word Conversion
## Convert HTML string to Word document
```csharp
// Create a new instance of the Document class.
Document document = new Document();

// Add a new section to the document.
Section sec = document.AddSection();

// Add a new paragraph to the section and append the HTML content to it.
sec.AddParagraph().AppendHTML(HTML);
```

---

# Spire.Doc C# HTML to Image Conversion
## Convert HTML document to image format
```csharp
//Create Word document.
Document document = new Document();

//Load the file from disk.
document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_HtmlFile1.html", FileFormat.Html, XHTMLValidationType.None);

string result = "Result-HtmlToImage.png";

//Save to image. You can convert HTML to BMP, JPEG, PNG, GIF, Tiff，etc.
Image image = document.SaveToImages(0, ImageType.Bitmap);
image.Save(result, ImageFormat.Png);

//Dispose the document
document.Dispose();
```

---

# spire.doc csharp html to pdf conversion
## convert HTML document to PDF format
```csharp
//Create a Word document
Document document = new Document();

//Load the file from disk
document.LoadFromFile("Template_HtmlFile.html", FileFormat.Html, XHTMLValidationType.None);

//Save to file
document.SaveToFile("Result-HtmlToPdf.pdf", FileFormat.PDF);

//Dispose the document
document.Dispose();
```

---

# spire.doc csharp html to xml conversion
## convert html file to xml format using spire.doc
```csharp
//Create Word document.
Document document = new Document();

//Load the file from disk.
document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_HtmlFile.html");

string result = "Result-HtmlToXml.xml";

//Save to file.
document.SaveToFile(result, FileFormat.Xml);

//Dispose the document
document.Dispose();
```

---

# Spire.Doc HTML to XPS Conversion
## Convert HTML document to XPS format using Spire.Doc library
```csharp
//Create Word document.
Document document = new Document();

//Load the file from disk.
document.LoadFromFile("Template_HtmlFile.html", FileFormat.Html, XHTMLValidationType.None);

//Save to file.
document.SaveToFile("Result-HtmlToXps.xps", FileFormat.XPS);

//Dispose the document
document.Dispose();
```

---

# spire.doc csharp image to pdf
## Convert image to PDF document using Spire.Doc
```csharp
//Create a new document
Document doc = new Document();

//Create a new section
Section section = doc.AddSection();

//Create a new paragraph
Paragraph paragraph = section.AddParagraph();

//Add a picture for paragraph
paragraph.AppendPicture(input);

//Set A4 page size
section.PageSetup.PageSize = PageSize.A4;

//Set the page margins
section.PageSetup.Margins.Top = 10f;
section.PageSetup.Margins.Left = 25f;
```

---

# Spire.Doc C# Keep Hidden Text
## Convert Word to PDF while preserving hidden text
```csharp
// Create a ToPdfParameterList object to specify conversion parameters
ToPdfParameterList pdf = new ToPdfParameterList();

// Set the 'IsHidden' parameter to true, which preserves any hidden text in the converted PDF
pdf.IsHidden = true;

// Save the document as a PDF using the specified conversion parameters
document.SaveToFile(result, pdf);
```

---

# spire.doc csharp markdown conversion
## Convert markdown to word or pdf
```csharp
// Create a new Document object
Document doc = new Document();

//Load .md file
doc.LoadFromFile(inputPath);

//Save to .pdf file
doc.SaveToFile("output.pdf", FileFormat.PDF);

// Dispose of the Document object
doc.Close();
```

---

# Spire.Doc ODT to Word Conversion
## Convert ODT files to Word documents using Spire.Doc library
```csharp
//Create Word document.
Document document = new Document();

//Load the file from disk.
document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_OdtFile.odt");

string result = "Result-OdtToDocOrDocx.docx";

//Save to Docx file.
document.SaveToFile(result, FileFormat.Docx);

//Dispose the document
document.Dispose();
```

---

# spire.doc csharp conversion
## preserve word bookmarks when converting to pdf
```csharp
// Create a new instance of the Document class
Document document = new Document();

// Load a Word document from the specified file path
document.LoadFromFile("Sample.doc");

// Create a ToPdfParameterList object to specify conversion parameters
ToPdfParameterList toPdf = new ToPdfParameterList();

// Set the 'CreateWordBookmarks' parameter to true to preserve bookmarks in the converted PDF
toPdf.CreateWordBookmarks = true;

// Set the 'WordBookmarksTitle' parameter to specify the title of the bookmarks in the PDF
toPdf.WordBookmarksTitle = "Bookmark";

// Set the 'WordBookmarksColor' parameter to specify the color of the bookmarks in the PDF
toPdf.WordBookmarksColor = Color.Gray;

// Attach an event handler to the BookmarkLayout event of the document
document.BookmarkLayout += new Spire.Doc.Documents.Rendering.BookmarkLevelHandler(document_BookmarkLayout);

// Save the document as a PDF with the specified conversion parameters
document.SaveToFile("PreserveBookmarks.pdf", toPdf);

// Define the event handler for the BookmarkLayout event
static void document_BookmarkLayout(object sender, Spire.Doc.Documents.Rendering.BookmarkLevelEventArgs args)
{
    // Customize the appearance of bookmarks based on their level
    if (args.BookmarkLevel.Level == 2)
    {
        args.BookmarkLevel.Color = Color.Red;
        args.BookmarkLevel.Style = BookmarkTextStyle.Bold;
    }
    else if (args.BookmarkLevel.Level == 3)
    {
        args.BookmarkLevel.Color = Color.Gray;
        args.BookmarkLevel.Style = BookmarkTextStyle.Italic;
    }
    else
    {
        args.BookmarkLevel.Color = Color.Green;
        args.BookmarkLevel.Style = BookmarkTextStyle.Regular;
    }
}
```

---

# Spire.Doc RTF to HTML Conversion
## Convert RTF document to HTML format using Spire.Doc library
```csharp
//Create Word document.
Document document = new Document();

//Load the file from disk.
document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_RtfFile.rtf");

string result = "Result-RtfToHtml.html";

//Save to file.
document.SaveToFile(result, FileFormat.Html);

//Dispose the document
document.Dispose();
```

---

# spire.doc csharp conversion
## convert RTF to PDF
```csharp
// Create Word document
Document document = new Document();

// Load the RTF file from disk
document.LoadFromFile(inputFile, FileFormat.Rtf);

// Save as PDF
document.SaveToFile(outputFile, FileFormat.PDF);

// Dispose the document
document.Dispose();
```

---

# Spire.Doc Comment Display Mode
## Set comment display mode when converting document to PDF
```csharp
// Set comment display mode when converting to pdf
document.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;
//document.LayoutOptions.CommentDisplayMode = CommentDisplayMode.Hide;
//document.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInBalloons;
```

---

# spire.doc csharp custom fonts
## set custom fonts for document conversion
```csharp
// Create a new Document instance
Document document = new Document();

// Create an InputStream for the custom font file
FileStream inputStream1 = new FileStream(@"..\..\..\..\..\..\..\Data\PT Serif Caption.ttf", FileMode.Open, FileAccess.Read);

// Create an array of InputStreams containing the custom font InputStream
Stream[] inputStreams = new Stream[] { inputStream1 };

// Set the custom fonts for the document
document.SetCustomFonts(inputStreams);

// Optionally set global custom fonts (commented out)
// Document.SetGlobalCustomFonts(inputStreams);

// Clear the custom fonts from the document
document.ClearCustomFonts();

// Optionally clear global custom fonts (commented out)
// Document.ClearGlobalCustomFonts();

// Dispose of the document to release resources
document.Dispose();
```

---

# spire.doc csharp font fallback
## set font fallback rules for document conversion
```csharp
// Create a new Document object
Document doc = new Document();

// Load the document from the specified file
doc.LoadFromFile("SetFontFallbackRule.docx");

// Load the font fallback rule settings from the XML file
doc.LoadFontFallbackRuleSettings("FontFallbackRule.xml");

// Save the document to a PDF file with the specified output file name
doc.SaveToFile("SetFontFallbackRule_output.pdf", FileFormat.PDF);

// Dispose the document object
doc.Dispose();
```

---

# Spire.Doc C# Image Quality
## Set JPEG quality for document conversion
```csharp
// Create a new instance of the Document class
Document document = new Document();

// Set the JPEG quality for saving images in the document to 40%
document.JPEGQuality = 40;
```

---

# Spire.Doc PDF Conversion with Embedded Fonts
## Specify which fonts should be embedded when converting a Word document to PDF
```csharp
// Create a ToPdfParameterList object to specify conversion parameters
ToPdfParameterList parms = new ToPdfParameterList();

// Create a list to store the names of embedded fonts to be used in the PDF
List<string> part = new List<string>();

// Add a font name, "PT Serif Caption", to the list of embedded fonts
part.Add("PT Serif Caption");

// Set the 'EmbeddedFontNameList' parameter to the list of embedded font names
parms.EmbeddedFontNameList = part;

// Save the document as a PDF using the specified conversion parameters
document.SaveToFile("output.pdf", parms);
```

---

# spire.doc csharp document conversion
## convert word document to epub format
```csharp
// Create a new instance of the Document class.
Document doc = new Document();

// Load a document from the specified file path.
doc.LoadFromFile("documentPath.doc");

// Specify the output file name for the EPUB file.
string result = "result.epub";

// Save the document as an EPUB file with the specified output file name and format.
doc.SaveToFile(result, FileFormat.EPub);

// Dispose the document object to release resources.
doc.Dispose();
```

---

# Spire.Doc C# Document Conversion
## Convert Word document to HTML format
```csharp
// Create a new instance of the Document class.
Document document = new Document();

// Load a Word document from the specified file path.
document.LoadFromFile(@"..\..\..\..\..\..\..\Data\ToHtmlTemplate.docx");

// Save the loaded document as an HTML file.
document.SaveToFile("Sample.html", FileFormat.Html);

// Release system resources associated with the Document object.
document.Dispose();
```

---

# spire.doc csharp html export
## configure HTML export options for Word document conversion
```csharp
// Set the file name for the CSS style sheet that will be used in the HTML export.
document.HtmlExportOptions.CssStyleSheetFileName = "sample.css";

// Specify that the CSS style sheet should be external.
document.HtmlExportOptions.CssStyleSheetType = CssStyleSheetType.External;

// Disable embedding images in the HTML output.
document.HtmlExportOptions.ImageEmbedded = false;

// Set the path where the exported HTML file will look for image resources.
document.HtmlExportOptions.ImagesPath = "Images";

// Treat text input form fields as plain text instead of form fields.
document.HtmlExportOptions.IsTextInputFormFieldAsText = true;

// Save the document as an HTML file with the specified file name and format.
document.SaveToFile("Sample.html", FileFormat.Html);
```

---

# Spire.Doc C# Document Conversion
## Convert Word document to image
```csharp
// Create word document
Document document = new Document();

// Load the file from disk
document.LoadFromFile(@"..\..\..\..\..\..\Data\ConvertedTemplate.docx");

// Save the first page to image
Image img = document.SaveToImages(0, ImageType.Bitmap);

// Save the image to file
img.Save("sample.png", ImageFormat.Png);

// Dispose the document
document.Dispose();
```

---

# spire.doc csharp conversion
## convert Word document to ODT format
```csharp
//Create word document
Document document = new Document();

//Load the file from disk.
document.LoadFromFile(@"..\..\..\..\..\..\Data\ToOdt.doc");

//Save to odt file.
document.SaveToFile("Sample.odt", FileFormat.Odt);

//Dispose the document
document.Dispose();
```

---

# spire.doc csharp document conversion
## convert document to pcl format
```csharp
//Create word document
Document doc = new Document();

//Load the file from disk.
doc.LoadFromFile("input file path");

//On Net4.6 and above platforms with adding the following external dependencies, you can set the UseHarfBuzzTextShaper which can better handling Thai and Tibetan characters
//external reference to:  
//HarfBuzzSharp >= 2.6.1.5
//System.Buffers >= 4.4.0
//System.Memory >= 4.5.3
//System.Numerics.Vectors >= 4.4.0
//System.Runtime.CompilerServices.Unsafe >= 4.5.2

//document.LayoutOptions.UseHarfBuzzTextShaper = true;

string result = "ToPCL.pcl";

//Save to PCL file
doc.SaveToFile(result, FileFormat.PCL);

//Dispose the document
doc.Dispose();
```

---

# Spire.Doc Word to PDF Conversion
## Convert Word documents to PDF format using Spire.Doc library
```csharp
// Create a new Document object
Document document = new Document();

// Load a Word document from the specified file path
document.LoadFromFile(inputFilePath);

// Save the document as a PDF file
document.SaveToFile(outputFilePath, FileFormat.PDF);

// Dispose of the Document object to free up resources
document.Dispose();
```

---

# spire.doc csharp pdf conversion
## convert document to pdf with bookmarks
```csharp
// Create a new Document instance
Document document = new Document();

// Load the document from the specified input file
document.LoadFromFile(inputFile);

// Create a ToPdfParameterList instance to configure PDF conversion options
ToPdfParameterList parames = new ToPdfParameterList();

// Enable the creation of bookmarks in the resulting PDF
parames.CreateWordBookmarks = true;

// Choose whether to create bookmarks using headings (true) or not (false)
parames.CreateWordBookmarksUsingHeadings = false;
// Uncomment this line to enable creating bookmarks using headings
//parames.CreateWordBookmarksUsingHeadings = true; 

// Save the document as a PDF file with the specified output file name and conversion parameters
document.SaveToFile(outFile, parames);
```

---

# spire.doc csharp pdf conversion
## disable hyperlinks when converting word to pdf
```csharp
//Create an instance of ToPdfParameterList.
ToPdfParameterList pdf = new ToPdfParameterList();

//Set DisableLink to true to remove the hyperlink effect for the result PDF page. 
//Set DisableLink to false to preserve the hyperlink effect for the result PDF page.
pdf.DisableLink = true;

//Save to file.
document.SaveToFile(result, pdf);
```

---

# spire.doc pdf conversion with font embedding
## convert document to pdf with all fonts embedded
```csharp
Document document = new Document();
document.LoadFromFile(@"..\..\..\..\..\..\..\Data\ConvertedTemplate.docx");
//embeds full fonts by default when IsEmbeddedAllFonts is set to true.
ToPdfParameterList ppl = new ToPdfParameterList();
ppl.IsEmbeddedAllFonts = true;

//Save doc file to pdf.
document.SaveToFile("Sample.pdf", ppl);
```

---

# spire.doc csharp pdf conversion
## embed non-installed fonts when converting word to pdf
```csharp
// Create a new document
Document document = new Document();
document.LoadFromFile("input.docx");

// Embed the non-installed fonts
ToPdfParameterList parms = new ToPdfParameterList();
List<PrivateFontPath> fonts = new List<PrivateFontPath>();
fonts.Add(new PrivateFontPath("PT Serif Caption", "path_to_font.ttf"));
parms.PrivateFontPaths = fonts;

// Save the document to a PDF file with embedded fonts
document.SaveToFile("output.pdf", parms);
```

---

# Spire.Doc C# Conversion
## Convert Word to PDF while keeping hidden text
```csharp
//Create Word document.
Document document = new Document();

//Load the file from disk.
document.LoadFromFile(@"..\..\..\..\..\..\..\Data\Template_Docx_5.docx");

//When convert to PDF file, set the property IsHidden as true.
ToPdfParameterList pdf = new ToPdfParameterList();
pdf.IsHidden = true;

String result = "Result-SaveTheHiddenTextToPDF.pdf";

//Save to file.
document.SaveToFile(result, pdf);
```

---

# spire.doc csharp pdf conversion
## convert document to pdf while preserving form fields
```csharp
// Preserve form field when converting to Pdf
ToPdfParameterList ppl = new ToPdfParameterList();
ppl.PreserveFormFields = true;

document.SaveToFile("ToPdfPreserveFormFields_output.pdf", ppl);
```

---

# spire.doc csharp pdf conversion
## convert word document to pdf while preserving bookmarks
```csharp
Document document = new Document();
document.LoadFromFile("Sample.doc");

ToPdfParameterList toPdf = new ToPdfParameterList();
toPdf.CreateWordBookmarks = true;
toPdf.WordBookmarksTitle = "Bookmark";
toPdf.WordBookmarksColor = Color.Gray;

// the event of BookmarkLayout occurs when drawing a bookmark
document.BookmarkLayout += new Spire.Doc.Documents.Rendering.BookmarkLevelHandler(document_BookmarkLayout);

// Save the document to a PDF file.
document.SaveToFile("PreserveBookmarks.pdf", toPdf);

static void document_BookmarkLayout(object sender, Spire.Doc.Documents.Rendering.BookmarkLevelEventArgs args)
{
    if (args.BookmarkLevel.Level == 2)
    {
        args.BookmarkLevel.Color = Color.Red;
        args.BookmarkLevel.Style = BookmarkTextStyle.Bold;
    }
    else if (args.BookmarkLevel.Level == 3)
    {
        args.BookmarkLevel.Color = Color.Gray;
        args.BookmarkLevel.Style = BookmarkTextStyle.Italic;
    }
    else
    {
        args.BookmarkLevel.Color = Color.Green;
        args.BookmarkLevel.Style = BookmarkTextStyle.Regular;
    }
}
```

---

# spire.doc csharp pdf conversion
## set image quality when converting document to pdf
```csharp
//Create Word document.
Document document = new Document();

//Set the output image quality to be 40% of the original image. The default set of the output image quality is 80% of the original.
document.JPEGQuality = 40;

//Save to file.
document.SaveToFile(result, FileFormat.PDF);
```

---

# Spire.Doc PDF Conversion with Embedded Fonts
## Convert document to PDF with specified embedded fonts
```csharp
// Specify embedded font
ToPdfParameterList parms = new ToPdfParameterList();
List<string> fontList = new List<string>();
fontList.Add("PT Serif Caption");
parms.EmbeddedFontNameList = fontList;
```

---

# Spire.Doc C# PDF Conversion with Password
## Convert Word document to PDF with password encryption
```csharp
// Create a new Document instance
Document document = new Document();

// Load the Word document
document.LoadFromFile("input.docx");

// Create a ToPdfParameterList instance to configure PDF conversion options
ToPdfParameterList toPdf = new ToPdfParameterList();

// Set a password for the PDF encryption
string password = "E-iceblue";
toPdf.PdfSecurity.Encrypt(password, password, PdfPermissionsFlags.Default, PdfEncryptionKeySize.Key128Bit);

// Save the document as a PDF file with encryption
document.SaveToFile("output.pdf", toPdf);

// Dispose the Document object after use
document.Dispose();
```

---

# Spire.Doc C# Document Conversion
## Convert Word document to PostScript format
```csharp
//Create word document
Document doc = new Document();

//Load the file from disk.
doc.LoadFromFile(@"..\..\..\..\..\..\Data\ConvertedTemplate.docx");

string result = "ToPostScript.ps";

//Save to PS file
doc.SaveToFile(result, FileFormat.PostScript);

//Dispose the document
doc.Dispose();
```

---

# spire.doc csharp document conversion
## convert word document to rtf format
```csharp
//Create word document
Document document = new Document();

//Load the file from disk.
document.LoadFromFile(@"..\..\..\..\..\..\Data\ToRtf.doc");

//Save to RTF file.
document.SaveToFile("Sample.rtf", FileFormat.Rtf);

//Dispose the document
document.Dispose();
```

---

# Spire.Doc Document to SVG Conversion
## Convert Word document to SVG format using Spire.Doc library
```csharp
//Create word document
Document document = new Document();

//Load the file from disk.
document.LoadFromFile(@"..\..\..\..\..\..\Data\ToSVGTemplate.docx");

//Save to SVG file.
document.SaveToFile("Sample.svg", FileFormat.SVG);

//Dispose the document
document.Dispose();
```

---

# Spire.Doc CSharp Word to TIFF Conversion
## Convert Word document pages to TIFF image format
```csharp
//Create word document
Document document = new Document();
document.LoadFromFile("input.docx");

//Save the document to a tiff image.
JoinTiffImages(SaveAsImage(document),"output.tif",EncoderValue.CompressionLZW);

//Dispose the document
document.Dispose();

private static Image[] SaveAsImage(Document document)
{    
    //Save all the pages in the document to images.
    Image[] images = document.SaveToImages(ImageType.Bitmap);    
    return images;
}

private static ImageCodecInfo GetEncoderInfo(string mimeType)
{
    //Set the code information for the images.
    ImageCodecInfo[] encoders = ImageCodecInfo.GetImageEncoders();
    for (int j = 0; j < encoders.Length; j++)
    {
        if (encoders[j].MimeType == mimeType)
            return encoders[j];
    }
    throw new Exception(mimeType + " mime type not found in ImageCodecInfo");
}

public static void JoinTiffImages(Image[] images, string outFile, EncoderValue compressEncoder)
{
    //Set the encoder parameters.
    System.Drawing.Imaging.Encoder enc = System.Drawing.Imaging.Encoder.SaveFlag;
    EncoderParameters ep = new EncoderParameters(2);
    ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.MultiFrame);
    ep.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)compressEncoder);
    Image pages = images[0];
    int frame = 0;
    ImageCodecInfo info = GetEncoderInfo("image/tiff");
    foreach (Image img in images)
    {
        if (frame == 0)
        {
            pages = img;
            //Save the first frame.
            pages.Save(outFile, info, ep);
        }
        else
        {
            //Save the intermediate frames.
            ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.FrameDimensionPage);
            pages.SaveAdd(img, ep);
        }
        if (frame == images.Length - 1)
        {
            //Flush and close.
            ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.Flush);
            pages.SaveAdd(ep);
        }
        frame++;
    }
}
```

---

# spire.doc csharp document conversion
## convert Word document to XML format
```csharp
//Create word document.
Document document = new Document();

//Load the file from disk.
document.LoadFromFile(@"..\..\..\..\..\..\Data\Summary_of_Science.doc");

//Save to a xml file.
document.SaveToFile("Sample.xml", FileFormat.Xml);

//Dispose the document
document.Dispose();
```

---

# spire.doc csharp document conversion
## convert word document to xps format
```csharp
//Create word document
Document document = new Document();

//Load the file from disk.
document.LoadFromFile(@"..\..\..\..\..\..\Data\ConvertedTemplate.docx");

//Save the document to a xps file.
document.SaveToFile("Sample.xps", FileFormat.XPS);

//Dispose the document
document.Dispose();
```

---

# spire.doc csharp text to word conversion
## convert text file to word document
```csharp
//Create Word document
Document document = new Document();

//Load the file from disk
document.LoadFromFile(inputTextFile);

//Save the file
document.SaveToFile(outputWordFile, FileFormat.Docx2013);
```

---

# Spire.Doc C# Word to EMF Conversion
## Convert Word document to EMF image format
```csharp
//Create a Word document.
Document document = new Document();

//Load the Word document
document.LoadFromFile(inputPath, FileFormat.Docx);

//Convert the first page of document to EMF image
System.Drawing.Image image = document.SaveToImages(0, Spire.Doc.Documents.ImageType.Metafile);

//Save as EMF format
image.Save(outputPath, ImageFormat.Emf);

//Dispose the document
document.Dispose();
```

---

# spire.doc csharp conversion
## convert Word document to Markdown format
```csharp
// Create a new document
Document doc = new Document();

// Load .docx file
doc.LoadFromFile(inputPath);

// Save to .md file
doc.SaveToFile(outputPath, FileFormat.Markdown);

// Dispose of the Document object
doc.Close();
```

---

# Spire.Doc Word to PDFA Conversion
## Convert Word document to PDF/A format using Spire.Doc library
```csharp
// Create a Word document
Document document = new Document();

// Load the file from disk
document.LoadFromFile("Template_Docx_1.docx");

// Create a ToPdfParameterList
ToPdfParameterList toPdf = new ToPdfParameterList();

// Set the Conformance-level of the Pdf file to PDF_A1B
toPdf.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B;

// Save the file
document.SaveToFile("Result-WordToPDFA.pdf", toPdf);

// Dispose the document
document.Dispose();
```

---

# Spire.Doc C# Word to Text Conversion
## Convert Word document to text format
```csharp
// Create a Word document
Document document = new Document();

// Load the file from disk
document.LoadFromFile("Template_Docx_1.docx");

// Save the file as text
document.SaveToFile("Result-WordToText.txt", FileFormat.Txt);

// Dispose the document
document.Dispose();
```

---

# Spire.Doc Word to Word XML Conversion
## Convert Word documents to Word XML formats (WordML and WordXml)
```csharp
//Create a Word document
Document document = new Document();

//Load the file from disk
document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_1.docx");

string result1 = "Result-WordToWordML.xml";
string result2 = "Result-WordToWordXML.xml";

//For word 2003:
document.SaveToFile(result1, FileFormat.WordML);

//For word 2007:
document.SaveToFile(result2, FileFormat.WordXml);

//Dispose the document
document.Dispose();
```

---

# Spire.Doc XML to PDF Conversion
## Convert XML document to PDF format using Spire.Doc library
```csharp
//Create a Word document.
Document document = new Document();

//Load the file from disk.
document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_XmlFile.xml", FileFormat.Xml);

string result = "Result-XMLToPDF.pdf";

//Save to file.
document.SaveToFile(result, FileFormat.PDF);

//Dispose the document
document.Dispose();
```

---

# Spire.Doc XML to Word Conversion
## Convert XML file to Word document using Spire.Doc library
```csharp
// Create a Word document
Document document = new Document();

// Load the XML file from disk
document.LoadFromFile("Template_XmlFile.xml", FileFormat.Xml);

// Save as Word document
document.SaveToFile("Result-XMLToWord.docx", FileFormat.Docx2013);

// Dispose the document
document.Dispose();
```

---

# spire.doc csharp font color
## change font color in word document
```csharp
//Get the first section 
Section section = doc.Sections[0];

//Get the first paragraph
Paragraph p1 = section.Paragraphs[0];

//Iterate through the childObjects of the paragraph 1 
foreach (DocumentObject childObj in p1.ChildObjects)
{
    //Determine if the childObj is TextRange
    if (childObj is TextRange)
    {
        //Change text color
        TextRange tr = childObj as TextRange;
        tr.CharacterFormat.TextColor = Color.RosyBrown;
    }
}

//Get the second paragraph
Paragraph p2 = section.Paragraphs[1];

//Iterate through the childObjects of the paragraph 2
foreach (DocumentObject childObj in p2.ChildObjects)
{
    //Determine if the childObj is TextRange
    if (childObj is TextRange)
    {
        //Change text color
        TextRange tr = childObj as TextRange;
        tr.CharacterFormat.TextColor = Color.DarkGreen;
    }
}
```

---

# spire.doc csharp font embedding
## Embed private font in Word document
```csharp
//Create a Word document
Document doc = new Document();

//Add a paragraph with text
Section section = doc.AddSection();
Paragraph p = section.AddParagraph();
TextRange range = p.AppendText("Spire.Doc for .NET is a professional Word.NET library specifically designed for developers to create, read, write, convert and print Word document files from any.NET platform with fast and high quality performance.");

//Set font properties
range.CharacterFormat.FontName = "PT Serif Caption";
range.CharacterFormat.FontSize = 20;

//Enable font embedding
doc.EmbedFontsInFile = true;

//Embed private font from file
doc.PrivateFontList.Add(new PrivateFontPath("PT Serif Caption", "path_to_font_file.ttf"));

//Save the document
doc.SaveToFile("EmbedPrivateFont.docx", FileFormat.Docx);
```

---

# spire.doc csharp font extraction
## get list of fonts used in a word document
```csharp
//Create a dictionary to store font and text range
Dictionary<Font, TextRange> font_obj = new Dictionary<Font, TextRange>() { };

//Create a Word document.
Document document = new Document();

//Load the file from disk.
document.LoadFromFile(input);

//Loop through the sections
foreach (Section section in document.Sections)
{
    //Loop through the paragraphs
    foreach (Paragraph paragraph in section.Body.Paragraphs)
    {
        //Loop through the child objects of the paragraph
        foreach (DocumentObject obj in paragraph.ChildObjects)
        {
            //Determine the Document Object Type of the child object
            if (obj.DocumentObjectType.Equals(DocumentObjectType.TextRange))
            {
                TextRange range = obj as TextRange;

                //Get the font 
                Font font = range.CharacterFormat.Font;

                // Determine if the font is already exists or not
                if (!font_obj.ContainsKey(font))
                {
                    font_obj.Add(font, range);
                }
            }
        }
    }
}

//Loop through dictionary
foreach (var item in font_obj)
{
    //Get the font
    Font font = item.Key;

    //Get the text range
    TextRange range = item.Value;

    //Format the font name, size,style and color
    string s = string.Format("Font Name: {0}, Size:{1}, Style:{2}, Color:{3}", font.Name, font.Size, font.Style, range.CharacterFormat.TextColor.Name);
    stringBuilder.AppendLine(s);
}
```

---

# Spire.Doc C# Font Setting
## Set font for text ranges in a Word document paragraph
```csharp
//Get the first section 
Section s = doc.Sections[0];

//Get the second paragraph
Paragraph p = s.Paragraphs[1];

//Create a characterFormat object
CharacterFormat format = new CharacterFormat(doc);
//Set font
format.Font = new Font("Arial", 16);

//Loop through the childObjects of paragraph 
foreach (DocumentObject childObj in p.ChildObjects)
{
    if (childObj is TextRange)
    {
        //Apply character format
        TextRange tr = childObj as TextRange;
        tr.ApplyCharacterFormat(format);
    }
}
```

---

# spire.doc csharp bullet style
## create bullet style using ASCII characters
```csharp
//Create a new document
Document document = new Document();
Section section = document.AddSection();

//Create a list style based on ASCII characters
ListStyle listStyle = new ListStyle(document, ListType.Bulleted);

//Set the style name
listStyle.Name = "liststyle";

//Set the bullet character
listStyle.Levels[0].BulletCharacter = "\x006e";

//Set the font name
listStyle.Levels[0].CharacterFormat.FontName = "Wingdings";

//Add the list style to the document
document.ListStyles.Add(listStyle);

//Create a paragraph
Paragraph paragraph = section.Body.AddParagraph();

//Append text
paragraph.AppendText("Spire.Doc for .NET");

//Apply the style
paragraph.ListFormat.ApplyStyle(listStyle.Name);
```

---

# spire.doc csharp character formatting
## apply various character formatting options to text in a Word document
```csharp
//Add a section
Section sec = document.AddSection();

//Add a paragraph
Paragraph titleParagraph = sec.AddParagraph();

//Append text
titleParagraph.AppendText("Font Styles and Effects ");

//Apply the builtin style
titleParagraph.ApplyStyle(BuiltinStyle.Title);

//Add a new paragraph
Paragraph paragraph = sec.AddParagraph();

//Append text
TextRange tr = paragraph.AppendText("Strikethough Text");

//Set strikeout style
tr.CharacterFormat.IsStrikeout = true;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Shadow Text");

//Set shadow property of text
tr.CharacterFormat.IsShadow = true;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Small caps Text");

//Set IsSmallCaps property of text
tr.CharacterFormat.IsSmallCaps = true;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Double Strikethough Text");

//Set DoubleStrike property of text
tr.CharacterFormat.DoubleStrike = true;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Outline Text");
tr.CharacterFormat.IsOutLine = true;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("AllCaps Text");
tr.CharacterFormat.AllCaps = true;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
paragraph.AppendText("Text");
tr = paragraph.AppendText("SubScript");

//Apply CharacterFormat
tr.CharacterFormat.SubSuperScript = SubSuperScript.SubScript;

//Append text
tr = paragraph.AppendText("And");
tr = paragraph.AppendText("SuperScript");

//Apply CharacterFormat
tr.CharacterFormat.SubSuperScript = SubSuperScript.SuperScript;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Emboss Text");

//Apply CharacterFormat
tr.CharacterFormat.Emboss = true;
tr.CharacterFormat.TextColor = Color.White;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
paragraph.AppendText("Hidden:");
tr = paragraph.AppendText("Hidden Text");

//Apply CharacterFormat
tr.CharacterFormat.Hidden = true;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Engrave Text");

//Apply CharacterFormat
tr.CharacterFormat.Engrave = true;
tr.CharacterFormat.TextColor = Color.White;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("WesternFonts╓╨╬─╫╓╠х");

//Apply CharacterFormat
tr.CharacterFormat.FontNameAscii = "Calibri";
tr.CharacterFormat.FontNameNonFarEast = "Calibri";
tr.CharacterFormat.FontNameFarEast = "Simsun";

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Font Size");

//Apply CharacterFormat
tr.CharacterFormat.FontSize = 20;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Font Color");

//Apply CharacterFormat
tr.CharacterFormat.TextColor = Color.Red;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Bold Italic Text");

//Apply CharacterFormat
tr.CharacterFormat.Bold = true;
tr.CharacterFormat.Italic = true;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Underline Style");

//Apply CharacterFormat
tr.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Highlight Text");

//Apply CharacterFormat
tr.CharacterFormat.HighlightColor = Color.Yellow;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Text has shading");

//Apply CharacterFormat
tr.CharacterFormat.TextBackgroundColor = Color.Green;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Border Around Text");

//Apply CharacterFormat
tr.CharacterFormat.Border.BorderType = Spire.Doc.Documents.BorderStyle.Single;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Text Scale");

//Apply CharacterFormat
tr.CharacterFormat.TextScale = 150;

//Append a line break
paragraph.AppendBreak(BreakType.LineBreak);

//Append text
tr = paragraph.AppendText("Character Spacing is 2 point");

//Apply CharacterFormat
tr.CharacterFormat.CharacterSpacing = 2;
```

---

# spire.doc csharp style management
## copy styles from one document to another
```csharp
//Create a Word document
Document srcDoc = new Document();

//Load the file from disk
srcDoc.LoadFromFile("source_document.docx");

//Create another Word document
Document destDoc = new Document();

//Load destination document from disk
destDoc.LoadFromFile("destination_document.docx");

//Get the style collections of source document
Spire.Doc.Collections.StyleCollection styles = srcDoc.Styles;

//Loop through the styles of source document
foreach (Style style in styles)
{
    //Add the style to destination document
    destDoc.Styles.Add(style);
}
```

---

# Spire.Doc C# Character Spacing
## Get character spacing and font name from a document paragraph
```csharp
//Create a document
Document document = new Document();

//Load the document from disk.
document.LoadFromFile(@"..\..\..\..\..\..\Data\Insert.docx");

//Get the first section of document
Section section = document.Sections[0];

//Get the first paragraph 
Paragraph paragraph = section.Paragraphs[0];

//Define two variables
string fontName = "";
float fontSpacing = 0;

//Traverse the ChildObjects 
foreach (DocumentObject docObj in paragraph.ChildObjects)
{
    //If it is TextRange
    if (docObj is TextRange)
    {
        TextRange textRange = docObj as TextRange;

        //Get the font name
        fontName = textRange.CharacterFormat.Font.Name;

        //Get the character spacing
        fontSpacing = textRange.CharacterFormat.CharacterSpacing;
    }
}

//Dispose the document
document.Dispose();
```

---

# Spire.Doc C# Get Text by Style Name
## Extract text from a Word document based on a specific style name
```csharp
//Create a Word document
Document doc = new Document();

//Load a document
doc.LoadFromFile("document.docx");

//Create string builder
StringBuilder builder = new StringBuilder();

//Loop through sections
foreach (Section section in doc.Sections)
{
    //Loop through paragraphs
    foreach (Paragraph para in section.Paragraphs)
    {
        //Find the paragraph whose style name is "Heading1"
        if (para.StyleName == "Heading1")
        {
            //Write the text of paragraph
            builder.AppendLine(para.Text);
        }
    }
}
```

---

# spire.doc csharp lists
## create numbered and bulleted lists in word document
```csharp
//Create a Word document
Document document = new Document();

//Add a section
Section sec = document.AddSection();

//Add paragraph and set list style
Paragraph paragraph = sec.AddParagraph();
paragraph.AppendText("Lists");
paragraph.ApplyStyle(BuiltinStyle.Title);

//Add numbered list header
paragraph = sec.AddParagraph();
paragraph.AppendText("Numbered List:").CharacterFormat.Bold = true;

//Create numbered list style
ListStyle numberList = new ListStyle(document, ListType.Numbered);
numberList.Name = "numberList";
numberList.Levels[1].NumberPrefix = "\x0000.";
numberList.Levels[1].PatternType = ListPatternType.Arabic;
numberList.Levels[2].NumberPrefix = "\x0000.\x0001.";
numberList.Levels[2].PatternType = ListPatternType.Arabic;

//Create bullet list style
ListStyle bulletList = new ListStyle(document, ListType.Bulleted);
bulletList.Name = "bulletList";

//Add the list styles to document
document.ListStyles.Add(numberList);
document.ListStyles.Add(bulletList);

//Create numbered list items
paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 1");
paragraph.ListFormat.ApplyStyle(numberList.Name);

paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2");
paragraph.ListFormat.ApplyStyle(numberList.Name);

//Create nested numbered list items
paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2.1");
paragraph.ListFormat.ApplyStyle(numberList.Name);
paragraph.ListFormat.ListLevelNumber = 1;

paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2.2");
paragraph.ListFormat.ApplyStyle(numberList.Name);
paragraph.ListFormat.ListLevelNumber = 1;

//Create deeper nested numbered list items
paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2.2.1");
paragraph.ListFormat.ApplyStyle(numberList.Name);
paragraph.ListFormat.ListLevelNumber = 2;

paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2.2.2");
paragraph.ListFormat.ApplyStyle(numberList.Name);
paragraph.ListFormat.ListLevelNumber = 2;

paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2.2.3");
paragraph.ListFormat.ApplyStyle(numberList.Name);
paragraph.ListFormat.ListLevelNumber = 2;

paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2.3");
paragraph.ListFormat.ApplyStyle(numberList.Name);
paragraph.ListFormat.ListLevelNumber = 1;

paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 3");
paragraph.ListFormat.ApplyStyle(numberList.Name);

//Add bullet list header
paragraph = sec.AddParagraph();
paragraph.AppendText("Bulleted List:").CharacterFormat.Bold = true;

//Create bullet list items
paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 1");
paragraph.ListFormat.ApplyStyle(bulletList.Name);

paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2");
paragraph.ListFormat.ApplyStyle(bulletList.Name);

//Create nested bullet list items
paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2.1");
paragraph.ListFormat.ApplyStyle(bulletList.Name);
paragraph.ListFormat.ListLevelNumber = 1;

paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2.2");
paragraph.ListFormat.ApplyStyle(bulletList.Name);
paragraph.ListFormat.ListLevelNumber = 1;

paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 3");
paragraph.ListFormat.ApplyStyle(bulletList.Name);
```

---

# spire.doc csharp document formatting
## apply multiple styles within a paragraph
```csharp
//Create a Word document
Document doc = new Document();

//Add a section
Section section = doc.AddSection();

//Add a paragraph
Paragraph para = section.AddParagraph();

//Add a text range
TextRange range = para.AppendText("Spire.Doc for .NET ");

//Set the font name
range.CharacterFormat.FontName = "Calibri";

//Set the font size
range.CharacterFormat.FontSize = 16f;

//Set the text color
range.CharacterFormat.TextColor = Color.Blue;

//Set the bold style
range.CharacterFormat.Bold = true;

//Set the underline Style
range.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;

//Append the text
range = para.AppendText("is a professional Word .NET library");

//Set the font name
range.CharacterFormat.FontName = "Calibri";

//Set the font size
range.CharacterFormat.FontSize = 15f;
```

---

# spire.doc csharp paragraph formatting
## demonstrates various paragraph formatting options including borders, alignment, indentation, and spacing
```csharp
//Add a section
Section sec = document.AddSection();

//Add a paragraph
Paragraph para = sec.AddParagraph();

//Append text
para.AppendText("Paragraph Formatting");

//Apply the Title style
para.ApplyStyle(BuiltinStyle.Title);

//Add a paragraph
para = sec.AddParagraph();

//Append text
para.AppendText("This paragraph is surrounded with borders.");

//Set the border type
para.Format.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;

//Set the border color
para.Format.Borders.Color = Color.Red;

para = sec.AddParagraph();
para.AppendText("The alignment of this paragraph is Left.");

//Set the horizontal alignment style
para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

para = sec.AddParagraph();
para.AppendText("The alignment of this paragraph is Center.");

//Set the horizontal alignment style
para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

para = sec.AddParagraph();
para.AppendText("The alignment of this paragraph is Right.");

//Set the horizontal alignment style
para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

para = sec.AddParagraph();
para.AppendText("The alignment of this paragraph is justified.");

//Set the horizontal alignment style
para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Justify;

para = sec.AddParagraph();
para.AppendText("The alignment of this paragraph is distributed.");

//Set the horizontal alignment style
para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Distribute;

para = sec.AddParagraph();
para.AppendText("This paragraph has the gray shadow.");

//Set the backcolor
para.Format.BackColor = Color.Gray;

para = sec.AddParagraph();
para.AppendText("This paragraph has the following indentations: Left indentation is 10pt, right indentation is 10pt, first line indentation is 15pt.");

//Set the indent
para.Format.SetLeftIndent(10);
para.Format.SetRightIndent(10);
para.Format.SetFirstLineIndent(15);

para = sec.AddParagraph();
para.AppendText("The hanging indentation of this paragraph is 15pt.");
//Negative value represents hanging indentation
para.Format.SetFirstLineIndent(-15);

para = sec.AddParagraph();
para.AppendText("This paragraph has the following spacing: spacing before is 10pt, spacing after is 20pt, line spacing is at least 10pt.");

//Set the spacing (in points) after the paragraph
para.Format.AfterSpacing = 20;

//Set the spacing (in points) before the paragraph
para.Format.BeforeSpacing = 10;

//Set the LineSpacingRule
para.Format.LineSpacingRule = LineSpacingRule.AtLeast;

//Set line spacing property of the paragraph.
para.Format.LineSpacing = 10;
```

---

# spire.doc csharp list numbering
## restart list numbering in word document
```csharp
//Create a numberList
ListStyle numberList = new ListStyle(document, ListType.Numbered);

//Set the name
numberList.Name = "Numbered1";

//Add the numberList to document
document.ListStyles.Add(numberList);

//Add paragraph and apply the list style
paragraph = section.AddParagraph();
paragraph.AppendText("List Item 1");
paragraph.ListFormat.ApplyStyle(numberList.Name);

paragraph = section.AddParagraph();
paragraph.AppendText("List Item 2");
paragraph.ListFormat.ApplyStyle(numberList.Name);

paragraph = section.AddParagraph();
paragraph.AppendText("List Item 3");
paragraph.ListFormat.ApplyStyle(numberList.Name);

paragraph = section.AddParagraph();
paragraph.AppendText("List Item 4");
paragraph.ListFormat.ApplyStyle(numberList.Name);

ListStyle numberList2 = new ListStyle(document, ListType.Numbered);
numberList2.Name = "Numbered2";
//set start number of second list
numberList2.Levels[0].StartAt = 10;
document.ListStyles.Add(numberList2);

//Add paragraph and apply the list style
paragraph = section.AddParagraph();
paragraph.AppendText("List Item 5");
paragraph.ListFormat.ApplyStyle(numberList2.Name);

paragraph = section.AddParagraph();
paragraph.AppendText("List Item 6");
paragraph.ListFormat.ApplyStyle(numberList2.Name);

paragraph = section.AddParagraph();
paragraph.AppendText("List Item 7");
paragraph.ListFormat.ApplyStyle(numberList2.Name);

paragraph = section.AddParagraph();
paragraph.AppendText("List Item 8");
paragraph.ListFormat.ApplyStyle(numberList2.Name);
```

---

# Spire.Doc C# Style Retrieval
## Retrieve style names from paragraphs in a Word document
```csharp
//Create and load a Word document
Document doc = new Document(@"..\..\..\..\..\..\Data\Styles.docx");

//Traverse all paragraphs in the document and get their style names through StyleName property
string styleName = null;

//Loop through all the sections
foreach (Section section in doc.Sections)
{
    //Loop through all the paragraphs
    foreach (Paragraph paragraph in section.Paragraphs)
    {
        //Get the style name
        styleName += paragraph.StyleName + "\r\n";
    }
}
```

---

# spire.doc csharp document styling
## create and apply styles to word document
```csharp
//Initialize a document
Document document = new Document();

//Add a section
Section sec = document.AddSection();

//Add default title style to document and modify
Style titleStyle = document.AddStyle(BuiltinStyle.Title);

//Set the font and font size
titleStyle.CharacterFormat.Font = new System.Drawing.Font("cambria", 28);

//Set the text color
titleStyle.CharacterFormat.TextColor = Color.FromArgb(42, 123, 136);

//judge if it is Paragraph Style and then set paragraph format
if (titleStyle is ParagraphStyle)
{
    ParagraphStyle ps = titleStyle as ParagraphStyle;

    //Set the BorderType
    ps.ParagraphFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single;

    //Set the color
    ps.ParagraphFormat.Borders.Bottom.Color = Color.FromArgb(42, 123, 136);

    //Set the line width
    ps.ParagraphFormat.Borders.Bottom.LineWidth = 1.5f;

    //Set the horizontal alignment style
    ps.ParagraphFormat.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
}

//Add default normal style and modify
Style normalStyle = document.AddStyle(BuiltinStyle.Normal);
normalStyle.CharacterFormat.Font = new System.Drawing.Font("cambria", 11);

//Add default heading1 style
Style heading1Style = document.AddStyle(BuiltinStyle.Heading1);
heading1Style.CharacterFormat.Font = new System.Drawing.Font("cambria", 14);
heading1Style.CharacterFormat.Bold = true;
heading1Style.CharacterFormat.TextColor = Color.FromArgb(42, 123, 136);

//Add default heading2 style
Style heading2Style = document.AddStyle(BuiltinStyle.Heading2);
heading2Style.CharacterFormat.Font = new System.Drawing.Font("cambria", 12);
heading2Style.CharacterFormat.Bold = true;

//Create a bulletList
ListStyle bulletList = new ListStyle(document, ListType.Bulleted);
bulletList.CharacterFormat.Font = new System.Drawing.Font("cambria", 12);

//Set the bulletList name
bulletList.Name = "bulletList";

//Add the style
document.ListStyles.Add(bulletList);

//Apply the Title style
Paragraph paragraph = sec.AddParagraph();
paragraph.AppendText("Your Name");
paragraph.ApplyStyle(BuiltinStyle.Title);

//Apply styles to additional paragraphs
paragraph = sec.AddParagraph();
paragraph.AppendText("Address, City, ST ZIP Code | Telephone | Email");
paragraph.ApplyStyle(BuiltinStyle.Normal);

paragraph = sec.AddParagraph();
paragraph.AppendText("Objective");
paragraph.ApplyStyle(BuiltinStyle.Heading1);

//Apply bullet list style
paragraph = sec.AddParagraph();
paragraph.AppendText("Major:Text");
paragraph.ListFormat.ApplyStyle("bulletList");
```

---

# Spire.Doc C# Mail Merge
## Add hyperlink to mail merged image
```csharp
// Define the field names and corresponding image file names
var fieldNames = new string[] { "MyImage" };
var fieldValues = new string[] { @"..\..\..\..\..\..\Data\mailmerge_logo.png" };
// Attach an event handler for the MergeImageField event
doc.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MailMerge_MergeImageField);
// Execute the mail merge with the field names and values
doc.MailMerge.Execute(fieldNames, fieldValues);

// Event handler for the MergeImageField event
private void MailMerge_MergeImageField(object sender, MergeImageFieldEventArgs field)
{
    string filePath = field.ImageFileName;  // FieldValue as string;
    if (!string.IsNullOrEmpty(filePath))
    {
        field.Image = Image.FromFile(filePath);
        // Set the hyperlink for the merged image field
        field.ImageLink = "https://www.e-iceblue.com/";
    }
}
```

---

# spire.doc csharp mail merge
## alternate row coloring in mail merge
```csharp
// Create a MergeFieldEventHandler
doc.MailMerge.MergeField += new MergeFieldEventHandler(MailMerge_MergeField);

// Fill mergedField with data from dataTable
doc.MailMerge.ExecuteWidthRegion(orderTable);

int rowIndex = 0;
void MailMerge_MergeField(object sender, MergeFieldEventArgs args)
{
    // Catch the beginning of a new row.
    if (args.CurrentMergeField.FieldName.Equals("Name"))
    {
        // Set the color depending on whether the row number is even or odd.
        Color rowColor;
        if (rowIndex % 2 == 0)
            rowColor = Color.FromArgb(215, 227, 235);
        else
            rowColor = Color.FromArgb(240, 242, 242);

        //Get the owner cell
        TableCell cell = args.CurrentMergeField.OwnerParagraph.Owner as TableCell;

        //Get the owner row
        TableRow row = cell.OwnerRow;

        //Set the back color
        for (int i = 0; i < row.Cells.Count; i++)
        {
            row.Cells[i].CellFormat.Shading.BackgroundPatternColor = rowColor;
        }
        rowIndex++;
    }
}
```

---

# Spire.Doc Mail Merge Culture Change
## Change locale for mail merge operations in Word documents
```csharp
// Store the current culture so it can be set back once mail merge is complete.
CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;

// Set the current thread culture
Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");

// Execute mail merge
document.MailMerge.Execute(fieldNames, fieldValues);

// Restore the thread culture
Thread.CurrentThread.CurrentCulture = currentCulture;
```

---

# spire.doc csharp conditional fields
## Execute conditional IF fields in mail merge
```csharp
// Create a new Document object
Document doc = new Document();

// Add a Section to the document
Section section = doc.AddSection();

// Add a Paragraph to the section
Paragraph paragraph = section.AddParagraph();

// Create a new IfField object
IfField ifField1 = new IfField(doc);

// Set the type and code of the IfField
ifField1.Type = FieldType.FieldIf;
ifField1.Code = "IF ";

// Add the IfField to the paragraph
paragraph.Items.Add(ifField1);

// Append the fields and text to the paragraph
paragraph.AppendField("Count", FieldType.FieldMergeField);
paragraph.AppendText(" > ");
paragraph.AppendText("\"1\" ");
paragraph.AppendText("\"Greater than one\" ");
paragraph.AppendText("\"Less than one\"");

// Create and add the end field mark
IParagraphBase end1 = doc.CreateParagraphItem(ParagraphItemType.FieldMark);
(end1 as FieldMark).Type = FieldMarkType.FieldEnd;
paragraph.Items.Add(end1);

// Set the end field mark for the IfField
ifField1.End = end1 as FieldMark;

// Add another paragraph to the section
paragraph = section.AddParagraph();

// Create a new IfField object
IfField ifField2 = new IfField(doc);

// Set the type and code of the IfField
ifField2.Type = FieldType.FieldIf;
ifField2.Code = "IF ";

// Add the IfField to the paragraph
paragraph.Items.Add(ifField2);

// Append the fields and text to the paragraph
paragraph.AppendField("Age", FieldType.FieldMergeField);
paragraph.AppendText(" > ");
paragraph.AppendText("\"50\" ");
paragraph.AppendText("\"The old man\" ");
paragraph.AppendText("\"The young man\"");

// Create and add the end field mark
IParagraphBase end2 = doc.CreateParagraphItem(ParagraphItemType.FieldMark);
(end2 as FieldMark).Type = FieldMarkType.FieldEnd;
paragraph.Items.Add(end2);

// Set the end field mark for the IfField
ifField2.End = end2 as FieldMark;

// Set up field names and values for mail merge
string[] fieldName = { "Count", "Age" };
string[] fieldValue = { "2", "30" };

// Execute the mail merge
doc.MailMerge.Execute(fieldName, fieldValue);

// Set IsUpdateFields property to true
doc.IsUpdateFields = true;

// Dispose the document object
doc.Dispose();
```

---

# Spire.Doc Mail Merge with DataTable
## Execute mail merge operation using a DataTable as data source
```csharp
// Create a Document 
Document doc = new Document();

//Load a mail merge template file
doc.LoadFromFile(input);

//Fill mergedField with data from dataTable
doc.MailMerge.ExecuteWidthRegion(orderTable);

//Save to file
string result = "ExecuteWithDataTable_out.doc";
doc.SaveToFile(result, FileFormat.Doc);

// Dispose the document object
doc.Dispose();
```

---

# Spire.Doc Mail Merge Hide Empty Regions
## This code demonstrates how to hide empty regions during mail merge in a Word document

```csharp
//Create word document
Document document = new Document();

//Prepare sample data
string[] filedNames = new string[] { "Contact Name", "Fax", "Date" };
string[] filedValues = new string[] { "John Smith", "+1 (69) 123456", DateTime.Now.Date.ToString() };

//Set the value to remove paragraphs which contain empty field.
document.MailMerge.HideEmptyParagraphs = true;

//Set the value to remove group which contain empty field.
document.MailMerge.HideEmptyGroup = true;

//Begin mail merge
document.MailMerge.Execute(filedNames, filedValues);
```

---

# Spire.Doc CSharp Mail Merge
## Identify merge field names in a Word document
```csharp
//Create Word document.
Document document = new Document();

//Get the collection of group names.
string[] GroupNames = document.MailMerge.GetMergeGroupNames();

//Get the collection of merge field names in a specific group.
string[] MergeFieldNamesWithinRegion = document.MailMerge.GetMergeFieldNames("Products");

//Get the collection of all the merge field names.
string[] MergeFieldNames = document.MailMerge.GetMergeFieldNames();

StringBuilder content = new StringBuilder();
content.AppendLine("----------------Group Names-----------------------------------------");
for (int i = 0; i < GroupNames.Length; i++)
{
    content.AppendLine(GroupNames[i]);
}

content.AppendLine("----------------Merge field names within a specific group-----------");
for (int j = 0; j < MergeFieldNamesWithinRegion.Length; j++)
{
    content.AppendLine(MergeFieldNamesWithinRegion[j]);
}

content.AppendLine("----------------All of the merge field names------------------------");
for (int k = 0; k < MergeFieldNames.Length; k++)
{
    content.AppendLine(MergeFieldNames[k]);
}
```

---

# Spire.Doc Mail Merge
## Execute mail merge operation in a Word document
```csharp
//Create a Word document
Document document = new Document();

//Prepare mail merge data
string[] fieldNames = new string[] { /* field names */ };
string[] fieldValues = new string[] { /* field values */ };

//Begin the mail merge process
document.MailMerge.Execute(fieldNames, fieldValues);
```

---

# Spire.Doc Mail Merge Form Fields
## Perform mail merge operations with form fields in Word documents
```csharp
// Define the field names for the mail merge
string[] fieldNames = new string[] { "Contact Name", "Fax", "Date", "Urgent", "Share", "Submit", "Body" };

// Define the field values for the mail merge
string[] fieldValues = new string[] { "John Smith", "+1 (69) 123456", DateTime.Now.Date.ToString(),
    "Yes","No","Yes",
    "<b>It's very urgent. Please deal with it ASAP. </b>" };

// Subscribe to the MergeField event
document.MailMerge.MergeField += new MergeFieldEventHandler(MailMerge_MergeField);

// Execute the mail merge using the field names and values
document.MailMerge.Execute(fieldNames, fieldValues);

void MailMerge_MergeField(object sender, MergeFieldEventArgs args)
{
    if (args.FieldValue.ToString() == "Yes")
    {
        // Get the checkbox name from the field name
        string checkBoxName = args.FieldName;

        // Get the owner paragraph of the current merge field
        Paragraph para = args.CurrentMergeField.OwnerParagraph;

        // Get the index of the current merge field within its parent paragraph
        int index = para.ChildObjects.IndexOf(args.CurrentMergeField);

        // Create a new CheckBoxFormField
        CheckBoxFormField field = para.AppendField(checkBoxName, FieldType.FieldFormCheckBox) as CheckBoxFormField;

        // Insert the new checkbox field at the same index as the current merge field
        para.ChildObjects.Insert(index, field);

        // Remove the current merge field from the paragraph
        para.ChildObjects.Remove(args.CurrentMergeField);

        // Set the checkbox field as checked
        field.Checked = true;
    }
    
    if (args.FieldValue.ToString() == "No")
    {
        // Get the checkbox name from the field name
        string checkBoxName = args.FieldName;

        // Get the owner paragraph of the current merge field
        Paragraph para = args.CurrentMergeField.OwnerParagraph;

        // Get the index of the current merge field within its parent paragraph
        int index = para.ChildObjects.IndexOf(args.CurrentMergeField);

        // Create a new CheckBoxFormField
        CheckBoxFormField field = para.AppendField(checkBoxName, FieldType.FieldFormCheckBox) as CheckBoxFormField;

        // Insert the new checkbox field at the same index as the current merge field
        para.ChildObjects.Insert(index, field);

        // Remove the current merge field from the paragraph
        para.ChildObjects.Remove(args.CurrentMergeField);

        // Set the checkbox field as unchecked
        field.Checked = false;
    }
   
    if (args.FieldName == "Body")
    {
        // Get the owner paragraph of the current merge field
        Paragraph para = args.CurrentMergeField.OwnerParagraph;

        // Append the HTML content as plain text to the paragraph
        para.AppendHTML(args.FieldValue.ToString());

        // Remove the current merge field from the paragraph
        para.ChildObjects.Remove(args.CurrentMergeField);
    }

    if (args.FieldName == "Date")
    {
        // Get the text input name from the field name
        string textInputName = args.FieldName;

        // Get the owner paragraph of the current merge field
        Paragraph para = args.CurrentMergeField.OwnerParagraph;

        // Create a new TextFormField
        TextFormField field = para.AppendField(textInputName, FieldType.FieldFormTextInput) as TextFormField;

        // Remove the current merge field from the paragraph
        para.ChildObjects.Remove(args.CurrentMergeField);

        // Set the text value for the text input field
        field.Text = args.FieldValue.ToString();
    }
}
```

---

# spire.doc csharp mail merge
## mail merge with images in word document
```csharp
//Create a Word document
Document spireDoc = new Document();

//Define the field names for the mail merge
string[] fieldNames = new string[] { "ImageFile" };

//Define the field values for the mail merge
string[] fieldValues = new string[] { "image_path" };

//Subscribe to the MergeField event
spireDoc.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MailMerge_MergeImageField);

//Execute the mail merge using the field names and values
spireDoc.MailMerge.Execute(fieldNames, fieldValues);

// Event handler for merging image fields
private void MailMerge_MergeImageField(object sender, MergeImageFieldEventArgs field)
{
    //Determine if the image exists or not
    string filePath = field.ImageFileName;
    if (!string.IsNullOrEmpty(filePath))
    {
        //Load the image from specified path
        field.Image = Image.FromFile(filePath);
    }
}
```

---

# spire.doc csharp mail merge
## execute mail merge with field names and values
```csharp
// Create a Word document
Document doc = new Document();

// Load a mail merge template file
doc.LoadFromFile(inputPath);

// Define the field names for the mail merge
string[] fieldName = new string[] { "XX_Name" };

// Define the field values for the mail merge
string[] fieldValue = new string[] { "Jason Tang" };

// Execute the mail merge using the field names and values
doc.MailMerge.Execute(fieldName, fieldValue);

// Save to file
doc.SaveToFile(resultPath, FileFormat.Docx);

// Dispose the document object
doc.Dispose();
```

---

# spire.doc mail merge event handler
## handle mail merge events and add page breaks
```csharp
// Subscribe to the MergeField event
document.MailMerge.MergeField += new MergeFieldEventHandler(MailMerge_MergeField);

// Execute the mail merge using the customerRecords list as the data source
document.MailMerge.ExecuteGroup(new MailMergeDataTable("Customer", customerRecords));

void MailMerge_MergeField(object sender, MergeFieldEventArgs args)
{
    // Check if the current row index is greater than the lastIndex
    if (args.RowIndex > lastIndex)
    {
        // Update the lastIndex with the current row index
        lastIndex = args.RowIndex;

        // Add a page break before the current merge field
        AddPageBreakForMergeField(args.CurrentMergeField);
    }
}

void AddPageBreakForMergeField(IMergeField mergeField)
{
    // Find position of needing to add page break
    bool foundGroupStart = false;
    Paragraph paragraph = mergeField.PreviousSibling.Owner as Paragraph;
    MergeField previousMergeField = null;

    // Find the group start merge field by traversing the previous sibling paragraphs
    while (!foundGroupStart)
    {
        paragraph = paragraph.PreviousSibling as Paragraph;

        for (int i = 0; i < paragraph.Items.Count; i++)
        {
            previousMergeField = paragraph.Items[i] as MergeField;

            if ((previousMergeField != null) && (previousMergeField.Prefix == "GroupStart"))
            {
                foundGroupStart = true;
                break;
            }
        }
    }

    // Append a page break to the paragraph
    paragraph.AppendBreak(BreakType.PageBreak);
}
```

---

# spire.doc csharp nested mail merge
## Implement nested mail merge functionality with Spire.Doc
```csharp
// Create a Document object
Document document = new Document();

// Load a Word document from file
document.LoadFromFile("template.docx");

// Create a list to store DictionaryEntry objects
List<DictionaryEntry> list = new List<DictionaryEntry>();

// Create a DictionaryEntry for "Customer" with an empty value and add it to the list
DictionaryEntry dictionaryEntry = new DictionaryEntry("Customer", string.Empty);
list.Add(dictionaryEntry);

// Create a DictionaryEntry for "Order" with a nested region condition and add it to the list
dictionaryEntry = new DictionaryEntry("Order", "Customer_Id = %Customer.Customer_Id%");
list.Add(dictionaryEntry);

// Execute mail merge with nested regions using the DataSet and list of DictionaryEntry objects
document.MailMerge.ExecuteWidthNestedRegion(dsData, list);

// Save the merged document to a file 
document.SaveToFile("Sample.docx", FileFormat.Docx);

// Dispose the Document object 
document.Dispose();
```

---

# Spire.Doc C# Bookmark Content Copying
## Copy content from a bookmark to another location in a Word document
```csharp
//Get the bookmark by name.
Bookmark bookmark = doc.Bookmarks["Test"];
DocumentObject docObj = null;

//Judge if the paragraph includes the bookmark exists in the table, if it exists in cell,
//Then need to find its outermost parent object(Table),
//and get the start/end index of current object on body.
if ((bookmark.BookmarkStart.Owner as Paragraph).IsInCell)
{
    //Get the table object
    docObj = bookmark.BookmarkStart.Owner.Owner.Owner.Owner;
}
else
{
    //Get the owner paragraph
    docObj = bookmark.BookmarkStart.Owner;
}

//Get the index of the docObj
int startIndex = doc.Sections[0].Body.ChildObjects.IndexOf(docObj);

//Judge the postion of the BookmarkEnd
if ((bookmark.BookmarkEnd.Owner as Paragraph).IsInCell)
{
    docObj = bookmark.BookmarkEnd.Owner.Owner.Owner.Owner;
}
else
{
    docObj = bookmark.BookmarkEnd.Owner;
}

//Get the index of the docObj
int endIndex = doc.Sections[0].Body.ChildObjects.IndexOf(docObj);

//Get the start/end index of the bookmark object on the paragraph.
Paragraph para = bookmark.BookmarkStart.Owner as Paragraph;

//Get the index of BookmarkStart
int pStartIndex = para.ChildObjects.IndexOf(bookmark.BookmarkStart);
para = bookmark.BookmarkEnd.Owner as Paragraph;

//Get the index of the BookmarkEnd
int pEndIndex = para.ChildObjects.IndexOf(bookmark.BookmarkEnd);

//Get the content of current bookmark and copy.
TextBodySelection select = new TextBodySelection(doc.Sections[0].Body, startIndex, endIndex, pStartIndex, pEndIndex);
TextBodyPart body = new TextBodyPart(select);
for (int i = 0; i < body.BodyItems.Count; i++)
{
    doc.Sections[0].Body.ChildObjects.Add(body.BodyItems[i].Clone());
}
```

---

# spire.doc csharp bookmark
## create bookmarks in word document
```csharp
private void CreateBookmark(Section section)
{
    // Add a Paragraph to the section
    Paragraph paragraph = section.AddParagraph();

    // Add text with formatting and make it italic
    TextRange txtRange = paragraph.AppendText("The following example demonstrates how to create bookmark in a Word document.");
    txtRange.CharacterFormat.Italic = true;

    // Add an empty paragraph for spacing
    section.AddParagraph();

    // Add another paragraph with heading style and colored text
    paragraph = section.AddParagraph();
    txtRange = paragraph.AppendText("Simple Create Bookmark.");
    txtRange.CharacterFormat.TextColor = Color.CornflowerBlue;
    paragraph.ApplyStyle(BuiltinStyle.Heading2);

    // Add an empty paragraph for spacing
    section.AddParagraph();

    // Add a paragraph with a bookmark
    paragraph = section.AddParagraph();
    paragraph.AppendBookmarkStart("SimpleCreateBookmark");
    paragraph.AppendText("This is a simple bookmark.");
    paragraph.AppendBookmarkEnd("SimpleCreateBookmark");

    // Add an empty paragraph for spacing
    section.AddParagraph();

    // Add another paragraph with heading style and colored text
    paragraph = section.AddParagraph();
    txtRange = paragraph.AppendText("Nested Create Bookmark.");
    txtRange.CharacterFormat.TextColor = Color.CornflowerBlue;
    paragraph.ApplyStyle(BuiltinStyle.Heading2);

    // Add an empty paragraph for spacing
    section.AddParagraph();

    // Add a paragraph with nested bookmarks
    paragraph = section.AddParagraph();
    paragraph.AppendBookmarkStart("Root");
    txtRange = paragraph.AppendText(" This is Root data ");
    txtRange.CharacterFormat.Italic = true;
    paragraph.AppendBookmarkStart("NestedLevel1");
    txtRange = paragraph.AppendText(" This is Nested Level1 ");
    txtRange.CharacterFormat.Italic = true;
    txtRange.CharacterFormat.TextColor = Color.DarkSlateGray;
    paragraph.AppendBookmarkStart("NestedLevel2");
    txtRange = paragraph.AppendText(" This is Nested Level2 ");
    txtRange.CharacterFormat.Italic = true;
    txtRange.CharacterFormat.TextColor = Color.DimGray;
    paragraph.AppendBookmarkEnd("NestedLevel2");
    paragraph.AppendBookmarkEnd("NestedLevel1");
    paragraph.AppendBookmarkEnd("Root");
}
```

---

# spire.doc csharp bookmark
## create bookmark for table in word document
```csharp
//Create bookmark for a table
private void CreateBookmarkForTable(Document doc, Section section)
{
    //Add a paragraph
    Paragraph paragraph = section.AddParagraph();

    //Append text for added paragraph
    TextRange txtRange = paragraph.AppendText("The following example demonstrates how to create bookmark for a table in a Word document.");

    //Set the font in italic
    txtRange.CharacterFormat.Italic = true;

    //Append bookmark start
    paragraph.AppendBookmarkStart("CreateBookmark");

    //Append bookmark end
    paragraph.AppendBookmarkEnd("CreateBookmark");

    //Add table
    Table table = section.AddTable(true);

    //Set the number of rows and columns
    table.ResetCells(2, 2);

    //Append text for table cells
    TextRange range = table[0, 0].AddParagraph().AppendText("sampleA");
    range = table[0, 1].AddParagraph().AppendText("sampleB");
    range = table[1, 0].AddParagraph().AppendText("120");
    range = table[1, 1].AddParagraph().AppendText("260");

    //Get the bookmark by index.
    Bookmark bookmark = doc.Bookmarks[0];

    //Get the name of bookmark.
    String bookmarkName = bookmark.Name;

    //Locate the bookmark by name.
    BookmarksNavigator navigator = new BookmarksNavigator(doc);
    navigator.MoveToBookmark(bookmarkName);

    //Add table to TextBodyPart
    TextBodyPart part = navigator.GetBookmarkContent();
    part.BodyItems.Add(table);

    //Replace bookmark content with table
    navigator.ReplaceBookmarkContent(part);
}
```

---

# spire.doc csharp bookmark
## extract text from bookmark in word document
```csharp
//Creates a BookmarkNavigator instance to access the bookmark
BookmarksNavigator navigator = new BookmarksNavigator(doc);

//Locate a specific bookmark by bookmark name
navigator.MoveToBookmark("Content");

//Get the bookmark content
TextBodyPart textBodyPart = navigator.GetBookmarkContent();

//Define a variable to store the text
string text = null;

//Iterate through the items in the bookmark content to get the text
foreach (var item in textBodyPart.BodyItems)
{
    if (item is Paragraph)
    {
        //Iterate through the child objects of the paragraph
        foreach (var childObject in (item as Paragraph).ChildObjects)
        {
            if (childObject is TextRange)
            {
                //Append the text
                text += (childObject as TextRange).Text;
            }
        }
    }
}
```

---

# Spire.Doc C# Bookmarks
## Get bookmarks from a Word document by index and by name
```csharp
//Create word document
Document document = new Document();

//Get the bookmark by index.
Bookmark bookmark1 = document.Bookmarks[0];

//Get the bookmark by name.
Bookmark bookmark2 = document.Bookmarks["Test2"];
```

---

# spire.doc csharp bookmark
## insert document content at bookmark location
```csharp
//Get the first section of the first document 
Section section1 = document1.Sections[0];

//Locate the bookmark
BookmarksNavigator bn = new BookmarksNavigator(document1);

//Find bookmark by name
bn.MoveToBookmark("Test", true, true);

//Get bookmarkStart
BookmarkStart start = bn.CurrentBookmark.BookmarkStart;

//Get the owner paragraph
Paragraph para = start.OwnerParagraph;

//Get the para index
int index = section1.Body.ChildObjects.IndexOf(para);

//Loop through the sections
foreach (Section section2 in document2.Sections)
{
    foreach (Paragraph paragraph in section2.Paragraphs)
    {
        //Insert the paragraphs of document2
        section1.Body.ChildObjects.Insert(index++ + 1, paragraph.Clone() as Paragraph);
    }
}
```

---

# spire.doc csharp bookmark
## insert image at bookmark location
```csharp
//Create a word document
Document doc = new Document();

//Create an instance of BookmarksNavigator
BookmarksNavigator bn = new BookmarksNavigator(doc);

//Find a bookmark named Test
bn.MoveToBookmark("Test", true, true);

//Add a section
Section section0 = doc.AddSection();

//Add a paragraph for the section
Paragraph paragraph = section0.AddParagraph();

//Load an image
Image image = Image.FromFile(@"..\..\..\..\..\..\Data\Word.png");

//Add a picture into the paragraph
DocPicture picture = paragraph.AppendPicture(image);

//Add the paragraph at the position of bookmark
bn.InsertParagraph(paragraph);

//Remove the section0
doc.Sections.Remove(section0);
```

---

# spire.doc csharp bookmark removal
## Remove a bookmark from a Word document without removing its content
```csharp
//Create a word document
Document document = new Document();

//Load the document from disk.
document.LoadFromFile("Bookmark.docx");

//Get the bookmark by name.
Bookmark bookmark = document.Bookmarks["Test"];

//Remove the bookmark, not its content.
document.Bookmarks.Remove(bookmark);

// Dispose the document
document.Dispose();
```

---

# spire.doc csharp bookmark
## remove content from bookmark in word document
```csharp
//Get the bookmark by name.            
Bookmark bookmark = document.Bookmarks["Test"];

//Get the owner paragraph of bookmark start
Paragraph para = bookmark.BookmarkStart.Owner as Paragraph;

//Get the index of the bookmark start
int startIndex = para.ChildObjects.IndexOf(bookmark.BookmarkStart);

//Get the owner paragraph of bookmark end
para = bookmark.BookmarkEnd.Owner as Paragraph;

//Get the index of the bookmark end
int endIndex = para.ChildObjects.IndexOf(bookmark.BookmarkEnd);

//Remove the content object, and Start from next of BookmarkStart object, end up with previous of BookmarkEnd object. 
//This method is only to remove the content of the bookmark.
for (int i = startIndex + 1; i < endIndex; i++)
{
    para.ChildObjects.RemoveAt(startIndex + 1);
}
```

---

# spire.doc csharp bookmark replacement
## Replace the content of a bookmark in a Word document using Spire.Doc
```csharp
//Create a word document
Document doc = new Document();

//Load the document from disk.
doc.LoadFromFile("Bookmark.docx");

//Create a BookmarksNavigator instance
BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(doc);

//Locate the bookmark.
bookmarkNavigator.MoveToBookmark("Test");

//Replace the context with new.
bookmarkNavigator.ReplaceBookmarkContent("This is replaced content.", false);

//Save the document.
doc.SaveToFile("ReplaceBookMarkContent.docx", FileFormat.Docx);

//Dispose the document
doc.Dispose();
```

---

# Spire.Doc C# Replace Bookmark with Table
## This code demonstrates how to replace a bookmark in a Word document with a table using Spire.Doc for .NET.

```csharp
//Create a table
Table table = new Table(doc, true);

//Create a BookmarksNavigator instance
BookmarksNavigator navigator = new BookmarksNavigator(doc);

//Get the specific bookmark by its name
navigator.MoveToBookmark("Test");

//Create a TextBodyPart instance 
TextBodyPart part = new TextBodyPart(doc);

//Add the table to the textpart
part.BodyItems.Add(table);

//Replace the current bookmark content with the TextBodyPart object
navigator.ReplaceBookmarkContent(part);
```

---

# spire.doc csharp bookmark color
## set different colors and styles for bookmarks in word document when converting to pdf
```csharp
//Create an instance of ToPdfParameterList
ToPdfParameterList toPdf = new ToPdfParameterList();

//Set CreateWordBookmarks to true to use word bookmarks when create the bookmarks
toPdf.CreateWordBookmarks = true;

//Set the title of word bookmarks
toPdf.WordBookmarksTitle = "Changed bookmark";

//Set the text color of word bookmarks
toPdf.WordBookmarksColor = Color.Gray;

//Call the event document_BookmarkLayout when drawing a bookmark
doc.BookmarkLayout += new Spire.Doc.Documents.Rendering.BookmarkLevelHandler(document_BookmarkLayout);

//Set bookmark layout 
void document_BookmarkLayout(object sender, Spire.Doc.Documents.Rendering.BookmarkLevelEventArgs args)
{
    //set the different color for different levels of bookmarks
    if (args.BookmarkLevel.Level == 2)
    {
        args.BookmarkLevel.Color = Color.Red;
        args.BookmarkLevel.Style = BookmarkTextStyle.Bold;
    }
    else if (args.BookmarkLevel.Level == 3)
    {
        args.BookmarkLevel.Color = Color.Gray;
        args.BookmarkLevel.Style = BookmarkTextStyle.Italic;
    }
    else
    {
        args.BookmarkLevel.Color = Color.Green;
        args.BookmarkLevel.Style = BookmarkTextStyle.Regular;
    }
}
```

---

# spire.doc csharp comment
## add comment for specific text in word document
```csharp
private void InsertComments(Document doc, string keystring)
{
    //Find the key string
    TextSelection find = doc.FindString(keystring, false, true);

    //Create the commentmarkStart and commentmarkEnd
    CommentMark commentmarkStart = new CommentMark(doc);

    //Set the comment Id
    commentmarkStart.CommentId = 1;

    //Set the start type
    commentmarkStart.Type = CommentMarkType.CommentStart;

    CommentMark commentmarkEnd = new CommentMark(doc);
    commentmarkEnd.CommentId = 1;
    commentmarkEnd.Type = CommentMarkType.CommentEnd;

    //Add the content for comment
    Comment comment = new Comment(doc);

    //Add the text to the paragraph
    comment.Body.AddParagraph().Text = "Test comments";

    //Add author information
    comment.Format.Author = "E-iceblue";

    //Get the textRange
    TextRange range = find.GetAsOneRange();

    //Get its paragraph
    Paragraph para = range.OwnerParagraph;

    //Get the index of textRange 
    int index = para.ChildObjects.IndexOf(range);

    //Add comment
    para.ChildObjects.Add(comment);

    //Insert the commentmarkStart and commentmarkEnd
    para.ChildObjects.Insert(index, commentmarkStart);
    para.ChildObjects.Insert(index + 2, commentmarkEnd);
}
```

---

# Spire.Doc C# Comments
## Add comments to Word document
```csharp
private void InsertComments(Section section)
{          
    //Get the second paragraph
    Paragraph paragraph = section.Paragraphs[1];

    //Add comment
    Spire.Doc.Fields.Comment comment = paragraph.AppendComment("Spire.Doc for .NET");

    //Add author information
    comment.Format.Author = "E-iceblue";

    //Set the user initials.
    comment.Format.Initial = "CM";
}
```

---

# Spire.Doc C# Comment Extraction
## Extract comments from a Word document
```csharp
// Create a word document
Document doc = new Document();

// Load the file from disk
doc.LoadFromFile(input);

// Create a StringBuilder instance
StringBuilder SB = new StringBuilder();

// Traverse all comments
foreach (Comment comment in doc.Comments)
{
    foreach (Paragraph p in comment.Body.Paragraphs)
    {
        // Append the comments to the StringBuilder instance
        SB.AppendLine(p.Text);
    }
}

// Save to TXT File
string output = "ExtractComment.txt";
File.WriteAllText(output, SB.ToString());

// Dispose the document
doc.Dispose();
```

---

# spire.doc csharp comment
## insert picture into comment
```csharp
//Get the third paragraph in the first section
Paragraph paragraph = doc.Sections[0].Paragraphs[2];

//Add comment
Comment comment = paragraph.AppendComment("This is a comment.");

//Add author information
comment.Format.Author = "E-iceblue";

//Create a DocPicture instance
DocPicture docPicture = new DocPicture(doc);

//Load a picture
docPicture.LoadImage(Image.FromFile(imagePath));

//Insert the picture into the comment body
comment.Body.AddParagraph().ChildObjects.Add(docPicture);
```

---

# Spire.Doc C# Comment Operations
## Demonstrates how to replace and remove comments in a Word document using Spire.Doc library
```csharp
//Replace the content of the first comment
doc.Comments[0].Body.Paragraphs[0].Replace("This is the title", "This comment is changed.", false, false);

//Remove the second comment
doc.Comments.RemoveAt(1);
```

---

# spire.doc csharp comment handling
## remove content associated with comments in word document
```csharp
//Get the first comment
Comment comment = document.Comments[0];

//Get the paragraph of obtained comment
Paragraph para = comment.OwnerParagraph;

//Get index of the CommentMarkStart 
int startIndex = para.ChildObjects.IndexOf(comment.CommentMarkStart);

//Get index of the CommentMarkEnd
int endIndex = para.ChildObjects.IndexOf(comment.CommentMarkEnd);

//Create a list
List<TextRange> list = new List<TextRange>();

//Get TextRanges between the indexes
for (int i = startIndex; i < endIndex; i++)
{
    if (para.ChildObjects[i] is TextRange)
    {
        //Add the text range
        list.Add(para.ChildObjects[i] as TextRange);
    }
}

//Insert a new TextRange
TextRange textRange = new TextRange(document);

//clear the text
textRange.Text = null;

//Insert the new textRange
para.ChildObjects.Insert(endIndex, textRange);

//Remove previous TextRanges
for (int i = 0; i < list.Count; i++)
{
    para.ChildObjects.Remove(list[i]);
}
```

---

# spire.doc csharp comment reply
## reply to a comment in a Word document and add an image
```csharp
//get the first comment.
Comment comment1 = doc.Comments[0];

//create a new comment
Comment replyComment1 = new Comment(doc);

//Set the author
replyComment1.Format.Author = "E-iceblue";

//Append text
replyComment1.Body.AddParagraph().AppendText("Spire.Doc is a professional Word .NET library on operating Word documents.");

//add the new comment as a reply to the selected comment.
comment1.ReplyToComment(replyComment1);

//Create a DocPicture instance
DocPicture docPicture = new DocPicture(doc);

//Load an image
docPicture.LoadImage(Image.FromFile(imagePath));

//insert a picture in the comment
replyComment1.Body.Paragraphs[0].ChildObjects.Add(docPicture);
```

---

# Spire.Doc C# Barcode Image
## Add barcode image to Word document
```csharp
//Add barcode image
DocPicture picture = document.Sections[0].AddParagraph().AppendPicture(Image.FromFile(imgPath));
```

---

# Spire.Doc C# Add Horizontal Line
## Demonstrates how to add a horizontal line to a Word document
```csharp
// Create a Word document
Document doc = new Document();

// Add a section
Section sec = doc.AddSection();

// Add a paragraph
Paragraph para = sec.AddParagraph();

// Append a horizontal line
para.AppendHorizonalLine();
```

---

# spire.doc csharp image manipulation
## add image to document footer
```csharp
//Add a picture in footer 
DocPicture picture = document.Sections[0].HeadersFooters.Footer.AddParagraph().AppendPicture(Image.FromFile(imgPath));

//Set the picture's position and style
picture.VerticalOrigin = VerticalOrigin.Page;
picture.HorizontalOrigin = HorizontalOrigin.Page;
picture.VerticalAlignment = ShapeVerticalAlignment.Bottom;
picture.TextWrappingStyle = TextWrappingStyle.None;

//Add a textbox in footer
Spire.Doc.Fields.TextBox textbox = document.Sections[0].HeadersFooters.Footer.AddParagraph().AppendTextBox(150, 20);

//Set the textbox's position and style
textbox.VerticalOrigin = VerticalOrigin.Page;
textbox.HorizontalOrigin = HorizontalOrigin.Page;
textbox.HorizontalPosition = 300;
textbox.VerticalPosition = 700;
textbox.Body.AddParagraph().AppendText("Welcome to E-iceblue");
```

---

# spire.doc csharp shape group
## create a shape group with multiple shapes including text boxes and arrows
```csharp
// Create a document
Document doc = new Document();

// Add a section to the document
Section sec = doc.AddSection();

// Add a paragraph to the section
Paragraph para = sec.AddParagraph();

// Create a shape group and set its width and height
ShapeGroup shapegroup = para.AppendShapeGroup(375, 462);

// Set the horizontal position of the shape group
shapegroup.HorizontalPosition = 180;

// Calculate the scaling factors for width and height
float X = (float)(shapegroup.Width / 1000.0f);
float Y = (float)(shapegroup.Height / 1000.0f);

// Add a text box to the shape group
Spire.Doc.Fields.TextBox txtBox = new Spire.Doc.Fields.TextBox(doc);
txtBox.SetShapeType(ShapeType.RoundRectangle);
txtBox.Width = 125 / X;
txtBox.Height = 54 / Y;
Paragraph paragraph = txtBox.Body.AddParagraph();
paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
paragraph.AppendText("Start");
txtBox.HorizontalPosition = 19 / X;
txtBox.VerticalPosition = 27 / Y;
txtBox.Format.LineColor = Color.Green;
shapegroup.ChildObjects.Add(txtBox);

// Add an arrow shape to the shape group
ShapeObject arrowLineShape = new ShapeObject(doc, ShapeType.DownArrow);
arrowLineShape.Width = 16 / X;
arrowLineShape.Height = 40 / Y;
arrowLineShape.HorizontalPosition = 69 / X;
arrowLineShape.VerticalPosition = 87 / Y;
arrowLineShape.StrokeColor = Color.Purple;
shapegroup.ChildObjects.Add(arrowLineShape);

// Add another text box to the shape group
txtBox = new Spire.Doc.Fields.TextBox(doc);
txtBox.SetShapeType(ShapeType.Rectangle);
txtBox.Width = 125 / X;
txtBox.Height = 54 / Y;
paragraph = txtBox.Body.AddParagraph();
paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
paragraph.AppendText("Step 1");
txtBox.HorizontalPosition = 19 / X;
txtBox.VerticalPosition = 131 / Y;
txtBox.Format.LineColor = Color.Blue;
shapegroup.ChildObjects.Add(txtBox);

// Add another arrow shape to the shape group
arrowLineShape = new ShapeObject(doc, ShapeType.DownArrow);
arrowLineShape.Width = 16 / X;
arrowLineShape.Height = 40 / Y;
arrowLineShape.HorizontalPosition = 69 / X;
arrowLineShape.VerticalPosition = 192 / Y;
arrowLineShape.StrokeColor = Color.Purple;
shapegroup.ChildObjects.Add(arrowLineShape);

// Add another text box to the shape group
txtBox = new Spire.Doc.Fields.TextBox(doc);
txtBox.SetShapeType(ShapeType.Parallelogram);
txtBox.Width = 149 / X;
txtBox.Height = 59 / Y;
paragraph = txtBox.Body.AddParagraph();
paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
paragraph.AppendText("Step 2");
txtBox.HorizontalPosition = 7 / X;
txtBox.VerticalPosition = 236 / Y;
txtBox.Format.LineColor = Color.BlueViolet;
shapegroup.ChildObjects.Add(txtBox);

// Add another arrow shape to the shape group
arrowLineShape = new ShapeObject(doc, ShapeType.DownArrow);
arrowLineShape.Width = 16 / X;
arrowLineShape.Height = 40 / Y;
arrowLineShape.HorizontalPosition = 66 / X;
arrowLineShape.VerticalPosition = 300 / Y;
arrowLineShape.StrokeColor = Color.Purple;
shapegroup.ChildObjects.Add(arrowLineShape);

// Add another text box to the shape group
txtBox = new Spire.Doc.Fields.TextBox(doc);
txtBox.SetShapeType(ShapeType.Rectangle);
txtBox.Width = 125 / X;
txtBox.Height = 54 / Y;
paragraph = txtBox.Body.AddParagraph();
paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
paragraph.AppendText("Step 3");
txtBox.HorizontalPosition = 19 / X;
txtBox.VerticalPosition = 345 / Y;
txtBox.Format.LineColor = Color.Blue;
shapegroup.ChildObjects.Add(txtBox);
```

---

# Spire.Doc C# Shape Addition
## Add various shapes to a Word document
```csharp
// Create a new document
Document doc = new Document();

// Add a section to the document
Section sec = doc.AddSection();

// Add a paragraph to the section
Paragraph para = sec.AddParagraph();

int x = 60, y = 40, lineCount = 0;
for (int i = 1; i < 20; i++)
{
    // Check if the current line count is a multiple of 8
    if (lineCount > 0 && lineCount % 8 == 0)
    {
        // Append a page break to start a new page
        para.AppendBreak(BreakType.PageBreak);
        x = 60;
        y = 40;
        lineCount = 0;
    }

    // Append a shape to the paragraph
    ShapeObject shape = para.AppendShape(50, 50, (ShapeType)i);
    shape.HorizontalOrigin = HorizontalOrigin.Page;
    shape.HorizontalPosition = x;
    shape.VerticalOrigin = VerticalOrigin.Page;
    shape.VerticalPosition = y + 50;
    x = x + (int)shape.Width + 50;

    // Check if the shape count is a multiple of 5
    if (i > 0 && i % 5 == 0)
    {
        // Adjust the vertical position and line count
        y = y + (int)shape.Height + 120;
        lineCount++;
        x = 60;
    }
}
```

---

# Spire.Doc C# SVG Image
## Add SVG image to Word document
```csharp
// Create a new Document object
Document document = new Document();

// Add a new Section to the document
Section section = document.AddSection();

// Add a new Paragraph to the section
Paragraph paragraph = section.AddParagraph();

// Append the picture (SVG) to the paragraph
paragraph.AppendPicture(svgFilePath);
```

---

# spire.doc csharp shape alignment
## align shapes horizontally in a word document
```csharp
//Get the first section
Section section = doc.Sections[0];

//Loop through the paragraphs in the section
foreach (Paragraph para in section.Paragraphs)
{
    // Loop through the child objects in the paragraph
    foreach (DocumentObject obj in para.ChildObjects)
    {
        if (obj is ShapeObject)
        {
            //Set the horizontal alignment as center
            (obj as ShapeObject).HorizontalAlignment = ShapeHorizontalAlignment.Center;
        }
    }
}
```

---

# spire.doc csharp image extraction
## extract images from word document
```csharp
// Create a new Document object
Document document = new Document(@"..\..\..\..\..\..\Data\Template.docx");

// Create a queue to store composite objects
Queue<ICompositeObject> nodes = new Queue<ICompositeObject>();

// Enqueue the document as the initial node
nodes.Enqueue(document);

// Create a list to store images
IList<Image> images = new List<Image>();

// Traverse through the composite objects in the document
while (nodes.Count > 0)
{
    // Dequeue the next node
    ICompositeObject node = nodes.Dequeue();

    // Iterate through the child objects of the node
    foreach (IDocumentObject child in node.ChildObjects)
    {
        // If the child is a composite object, enqueue it for further processing
        if (child is ICompositeObject)
        {
            nodes.Enqueue(child as ICompositeObject);

            // If the child is a picture, add its image to the list
            if (child.DocumentObjectType == DocumentObjectType.Picture)
            {
                DocPicture picture = child as DocPicture;
                images.Add(picture.Image);
            }
        }
    }
}
```

---

# spire.doc csharp get alternative text
## extract alternative text from shapes in word document
```csharp
//Create string builder
StringBuilder builder = new StringBuilder();

//Loop through shapes and get the AlternativeText
foreach (Section section in document.Sections)
{
    //Loop through the paragraphs in the section
    foreach (Paragraph para in section.Paragraphs)
    {
        //Loop through the child objects in the paragraph
        foreach (DocumentObject obj in para.ChildObjects)
        {
            //If the shape is a shape object
            if (obj is ShapeObject)
            {
                string text = (obj as ShapeObject).AlternativeText;
                //Append the alternative text in builder
                builder.AppendLine(text);
            }
        }
    }
}
```

---

# spire.doc csharp image
## insert image into word document
```csharp
// Add a new paragraph to the section
Paragraph paragraph = section.AddParagraph();
paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

// Append the image to the paragraph and set its width and height
DocPicture picture = paragraph.AppendPicture(ima);
picture.Width = 100;
picture.Height = 100;

// Add a new paragraph to the section
paragraph = section.AddParagraph();
paragraph.Format.LineSpacing = 20f;

// Add text to the paragraph with specified formatting
TextRange tr = paragraph.AppendText("Spire.Doc for .NET is a professional Word .NET library specially designed for developers to create, read, write, convert and print Word document files from any .NET (C#, VB.NET, ASP.NET) platform with fast and high-quality performance.");
tr.CharacterFormat.FontName = "Arial";
tr.CharacterFormat.FontSize = 14;

// Add an empty paragraph to create spacing
section.AddParagraph();

// Add a new paragraph to the section
paragraph = section.AddParagraph();
paragraph.Format.LineSpacing = 20f;

// Add text to the paragraph with specified formatting
tr = paragraph.AppendText("As an independent Word .NET component, Spire.Doc for .NET doesn't need Microsoft Word to be installed on the machine. However, it can incorporate Microsoft Word document creation capabilities into any developers' .NET applications.");
tr.CharacterFormat.FontName = "Arial";
tr.CharacterFormat.FontSize = 14;
```

---

# Spire.Doc C# Image Insertion
## Insert and configure an image in a Word document

```csharp
//Create a word document
Document doc = new Document();

//Get the first section
Section section = doc.Sections[0];

//Add a new section or get the first section
Paragraph paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();

//Append text
paragraph.AppendText("The sample demonstrates how to insert an image into a document.");

//Apply style
paragraph.ApplyStyle(BuiltinStyle.Heading2);

//Add a new paragraph
paragraph = section.AddParagraph();

//Append text
paragraph.AppendText("The above is a picture.");

//Load an image 
Bitmap p = new Bitmap(Image.FromFile(@"..\..\..\..\..\..\Data\Word.png"));

//rotate image and insert image to word document
p.RotateFlip(RotateFlipType.Rotate90FlipX);

//Create a DocPicture instance
DocPicture picture = new DocPicture(doc);

//Load the image
picture.LoadImage(p);
//set image's position
picture.HorizontalPosition = 50.0F;
picture.VerticalPosition = 60.0F;

//set image's size
picture.Width = 200;
picture.Height = 200;

//set textWrappingStyle with image;
picture.TextWrappingStyle = TextWrappingStyle.Through;

//Insert the picture at the beginning of the second paragraph
paragraph.ChildObjects.Insert(0, picture);
```

---

# spire.doc csharp wordart
## insert WordArt into Word document
```csharp
//Add a paragraph.
Paragraph paragraph = doc.Sections[0].AddParagraph();

//Add a shape.
ShapeObject shape = paragraph.AppendShape(250, 70, ShapeType.TextWave4);

//Set the position of the shape.
shape.VerticalPosition = 20;
shape.HorizontalPosition = 80;

//set the text of WordArt.
shape.WordArt.Text = "Thanks for reading.";

//Set the fill color.
shape.FillColor = Color.Red;

//Set the border color of the text.
shape.StrokeColor = Color.Yellow;
```

---

# spire.doc remove shapes
## remove shape objects from word document
```csharp
//Get the first section
Section section = doc.Sections[0];

//Get all the child objects of paragraph
foreach (Paragraph para in section.Paragraphs)
{
    for (int i = 0; i < para.ChildObjects.Count; i++)
    {
        //If the child objects is shape object
        if (para.ChildObjects[i] is ShapeObject)
        {
            //Remove the shape object
            para.ChildObjects.RemoveAt(i);
        }
    }
}
```

---

# Spire.Doc C# Image Replacement
## Replace images with text in a Word document
```csharp
//Replace all pictures with texts
int j = 1;
foreach (Section sec in doc.Sections)
{
    foreach (Paragraph para in sec.Paragraphs)
    {
        List<DocumentObject> pictures = new List<DocumentObject>();
        //Get all pictures in the Word document
        foreach (DocumentObject docObj in para.ChildObjects)
        {
            if (docObj.DocumentObjectType == DocumentObjectType.Picture)
            {
                pictures.Add(docObj);
            }
        }

        //Replace pitures with the text "Here was image {image index}"
        foreach (DocumentObject pic in pictures)
        {
            //Get the index of the picture
            int index = para.ChildObjects.IndexOf(pic);

            //Create a new TextRange
            TextRange range = new TextRange(doc);

            //Format the text
            range.Text = string.Format("Here was image {0}", j);

            //Insert the textrange
            para.ChildObjects.Insert(index, range);

            //Remove the picture
            para.ChildObjects.Remove(pic);
            j++;
        }
    }
}
```

---

# spire.doc csharp image resize
## reset image size in word document
```csharp
//Get the first secion
Section section = doc.Sections[0];

//Get the first paragraph
Paragraph paragraph = section.Paragraphs[0];

//Reset the image size of the first paragraph
foreach (DocumentObject docObj in paragraph.ChildObjects)
{
    if (docObj is DocPicture)
    {
        DocPicture picture = docObj as DocPicture;

        //Set the width
        picture.Width = 50f;

        //Set the height
        picture.Height = 50f;
    }
}
```

---

# spire.doc csharp shape manipulation
## reset shape size in word document
```csharp
//Get the first section 
Section section = doc.Sections[0];

//Get the first paragraph
Paragraph para = section.Paragraphs[0];

//Get the second shape
ShapeObject shape = para.ChildObjects[1] as ShapeObject;

//Reset the width and height of the shape
shape.Width = 200;
shape.Height = 200;
```

---

# spire.doc csharp shape rotation
## rotate shape objects in word document
```csharp
//Get the first section
Section section = doc.Sections[0];

//Traverse the word document 
foreach (Paragraph para in section.Paragraphs)
{
    foreach (DocumentObject obj in para.ChildObjects)
    {
        if (obj is ShapeObject)
        {
            //Set the shape rotation as 20
            (obj as ShapeObject).Rotation = 20.0;
        }
    }
}
```

---

# Spire.Doc C# Line Shape Styling
## Set style properties for a line shape in a Word document
```csharp
//create a document
Document doc = new Document();

//Add a section
Section sec = doc.AddSection();

//Add a new paragraph
Paragraph para = sec.AddParagraph();

//Add a line shape
ShapeObject shape = para.AppendShape(100, 100, ShapeType.Line);

//Set style of Line shape
shape.FillColor = Color.Orange;
shape.StrokeColor = Color.Black;
shape.LineStyle = ShapeLineStyle.Single;
shape.LineDashing = LineDashing.LongDashDotDotGEL;
```

---

# spire.doc csharp text wrapping
## set text wrap style for images in word document
```csharp
//Loop through the sections
foreach (Section sec in doc.Sections)
{
    //Loop through the paragraphs
    foreach (Paragraph para in sec.Paragraphs)
    {
        //Create a list to store the pictures
        List<DocumentObject> pictures = new List<DocumentObject>();

        //Get all pictures in the Word document
        foreach (DocumentObject docObj in para.ChildObjects)
        {
            if (docObj.DocumentObjectType == DocumentObjectType.Picture)
            {
                pictures.Add(docObj);
            }
        }

        //Set text wrap styles for each piture
        foreach (DocumentObject pic in pictures)
        {
            //Create a DocPicture instance
            DocPicture picture = pic as DocPicture;

            //Set the wrap style and type
            picture.TextWrappingStyle = TextWrappingStyle.Through;
            picture.TextWrappingType = TextWrappingType.Both;
        }
    }
}
```

---

# spire.doc textbox transparency
## Set transparency for a textbox in a Word document
```csharp
//Create a word document
Document doc = new Document();

//Create a new section
Section section = doc.AddSection();

//Create a new paragraph
Paragraph paragraph = section.AddParagraph();

//Append TextBox
Spire.Doc.Fields.TextBox textbox1 = paragraph.AppendTextBox(100, 50);

//Set fill color
textbox1.Format.FillColor = Color.Red;

//Set fill transparency
textbox1.FillTransparency = 0.45;
```

---

# spire.doc csharp image transparency
## Set transparent color for images in Word document
```csharp
//Get the first paragraph in the first section
Paragraph paragraph = doc.Sections[0].Paragraphs[0];

//Loop through the child objects of the paragraph
foreach (DocumentObject obj in paragraph.ChildObjects)
{
    if (obj is DocPicture)
    {
        //Set the blue color of the image(s) in the paragraph to transparent
        DocPicture picture = obj as DocPicture;
        picture.TransparentColor = Color.Blue;
    }
}
```

---

# spire.doc csharp image update
## update images in word document
```csharp
//Create a list to store the pictures
List<DocumentObject> pictures = new List<DocumentObject>();

//Loop through the sections
foreach (Section sec in doc.Sections)
{
    //Loop through the paragraphs
    foreach (Paragraph para in sec.Paragraphs)
    {
        //Loop through the child objects of the paragraph
        foreach (DocumentObject docObj in para.ChildObjects)
        {
            //Determine if the type is picture or not
            if (docObj.DocumentObjectType == DocumentObjectType.Picture)
            {
                //Add the picture to list
                pictures.Add(docObj);
            }
        }
    }
}

//Create a DocPicture instance
DocPicture picture = pictures[0] as DocPicture;

//Replace the first picture with a new image file
picture.LoadImage(Image.FromFile(@"..\..\..\..\..\..\Data\E-iceblue.png"));
```

---

# Spire.Doc C# Header Management
## Add header only to the first page of a Word document
```csharp
// Get the header from the first section
HeaderFooter header = doc1.Sections[0].HeadersFooters.Header;

// Get the first page header of the destination document
HeaderFooter firstPageHeader = doc2.Sections[0].HeadersFooters.FirstPageHeader;

// Loop the sections of doc2
foreach (Section section in doc2.Sections)
{
    // Specify that the current section has a different header/footer for the first page
    section.PageSetup.DifferentFirstPageHeaderFooter = true;
}

// Removes all child objects in firstPageHeader
firstPageHeader.Paragraphs.Clear();

// Loop through the child objects of the header
foreach (DocumentObject obj in header.ChildObjects)
{
    // Add all child objects of the header to firstPageHeader
    firstPageHeader.ChildObjects.Add(obj.Clone());
}
```

---

# spire.doc csharp header footer height
## Adjust header and footer height in Word document
```csharp
//Get the first section
Section section = doc.Sections[0];

//Adjust the height of headers in the section
section.PageSetup.HeaderDistance = 100;

//Adjust the height of footers in the section
section.PageSetup.FooterDistance = 100;
```

---

# spire.doc csharp header footer
## copy header and footer between word documents
```csharp
//Get the header section from the source document
HeaderFooter header = doc1.Sections[0].HeadersFooters.Header;

//Loop through the sections of destination document
foreach (Section section in doc2.Sections)
{
    //Loop through the child objects of header
    foreach (DocumentObject obj in header.ChildObjects)
    {
        //Copy each object in the header of source file to destination file
        section.HeadersFooters.Header.ChildObjects.Add(obj.Clone());
    }
}
```

---

# spire.doc csharp header footer
## create different first page header and footer
```csharp
//Get the section
Section section = doc.Sections[0];

//specify that the current section has a different header/footer for the first page
section.PageSetup.DifferentFirstPageHeaderFooter = true;

//Set the first page header. Here we append a picture in the header
Paragraph paragraph1 = section.HeadersFooters.FirstPageHeader.AddParagraph();

//Set horizontal alignment for the paragraph
paragraph1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

//Append a picture
DocPicture headerimage = paragraph1.AppendPicture(Image.FromFile(imagePath));

//Set the first page footer
Paragraph paragraph2 = section.HeadersFooters.FirstPageFooter.AddParagraph();

//Set horizontal alignment for the paragraph
paragraph2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

//Append text
TextRange FF = paragraph2.AppendText("First Page Footer");

//Set font size
FF.CharacterFormat.FontSize = 10;

//Set the other header & footer. If you only need the first page header & footer, don't set this
Paragraph paragraph3 = section.HeadersFooters.Header.AddParagraph();

//Set horizontal alignment for the paragraph
paragraph3.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

//Append text
TextRange NH = paragraph3.AppendText("Spire.Doc for .NET");

//Set font size
NH.CharacterFormat.FontSize = 10;

//Add a paragraph
Paragraph paragraph4 = section.HeadersFooters.Footer.AddParagraph();

//Set horizontal alignment for the paragraph
paragraph4.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

//Append text
TextRange NF = paragraph4.AppendText("E-iceblue");

//Set font size
NF.CharacterFormat.FontSize = 10;
```

---

# Spire.Doc Header and Footer Implementation
## Add headers and footers with images, text, and page numbers to a Word document
```csharp
private void InsertHeaderAndFooter(Section section)
{
    //Get the header
    HeaderFooter header = section.HeadersFooters.Header;

    //Get the footer
    HeaderFooter footer = section.HeadersFooters.Footer;

    // Create a new paragraph for the header and add an image
    Paragraph headerParagraph = header.AddParagraph();
    DocPicture headerPicture = headerParagraph.AppendPicture(Image.FromFile("Header.png"));

    // Add text to the header paragraph and set its formatting properties
    TextRange text = headerParagraph.AppendText("Demo of Spire.Doc");
    text.CharacterFormat.FontName = "Arial";
    text.CharacterFormat.FontSize = 10;
    text.CharacterFormat.Italic = true;
    headerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

    // Set border properties for the bottom border of the header paragraph
    headerParagraph.Format.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.Single;
    headerParagraph.Format.Borders.Bottom.Space = 0.05F;

    // Set the text wrapping style and alignment properties for the header picture
    headerPicture.TextWrappingStyle = TextWrappingStyle.Behind;
    headerPicture.HorizontalOrigin = HorizontalOrigin.Page;
    headerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left;
    headerPicture.VerticalOrigin = VerticalOrigin.Page;
    headerPicture.VerticalAlignment = ShapeVerticalAlignment.Top;

    // Create a new paragraph for the footer and add an image
    Paragraph footerParagraph = footer.AddParagraph();
    DocPicture footerPicture = footerParagraph.AppendPicture(Image.FromFile("Footer.png"));

    // Set the text wrapping style and alignment properties for the footer picture
    footerPicture.TextWrappingStyle = TextWrappingStyle.Behind;
    footerPicture.HorizontalOrigin = HorizontalOrigin.Page;
    footerPicture.HorizontalAlignment = ShapeHorizontalAlignment.Left;
    footerPicture.VerticalOrigin = VerticalOrigin.Page;
    footerPicture.VerticalAlignment = ShapeVerticalAlignment.Bottom;

    // Add fields for page number and total number of pages to the footer paragraph
    footerParagraph.AppendField("page number", FieldType.FieldPage);
    footerParagraph.AppendText(" of ");
    footerParagraph.AppendField("number of pages", FieldType.FieldNumPages);
    footerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

    // Set border properties for the top border of the footer paragraph
    footerParagraph.Format.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Single;
    footerParagraph.Format.Borders.Top.Space = 0.05F;
}
```

---

# Spire.Doc C# Header and Footer
## Add images to document header and footer
```csharp
//Get the header of the first page
HeaderFooter header = doc.Sections[0].HeadersFooters.Header;

//Add a paragraph for the header
Paragraph paragraph = header.AddParagraph();

//Set the format of the paragraph
paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

//Append a picture in the paragraph
DocPicture headerimage = paragraph.AppendPicture(Image.FromFile(imagePath1));
headerimage.VerticalAlignment = ShapeVerticalAlignment.Bottom;

//Get the footer of the first section
HeaderFooter footer = doc.Sections[0].HeadersFooters.Footer;

//Add a paragraph for the footer
Paragraph paragraph2 = footer.AddParagraph();

//Set the format of the paragraph
paragraph2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

//Append a picture in the paragraph
paragraph2.AppendPicture(Image.FromFile(imagePath2));

//Append text in the paragraph and set its character format
TextRange TR = paragraph2.AppendText("Copyright © 2013 e-iceblue. All Rights Reserved.");
TR.CharacterFormat.FontName = "Arial";
TR.CharacterFormat.FontSize = 10;
TR.CharacterFormat.TextColor = Color.Black;
```

---

# Spire.Doc C# Header Protection
## Lock document content while allowing header/footer editing
```csharp
//Get the first section
Section section = doc.Sections[0];

//Protect the document and set the ProtectionType as AllowOnlyFormFields
doc.Protect(ProtectionType.AllowOnlyFormFields, "123");

//Set the ProtectForm as false to unprotect the section
section.ProtectForm = false;
```

---

# Spire.Doc C# Header Footer
## Create different headers and footers for odd and even pages
```csharp
//Get the first section
Section section = doc.Sections[0];

//Set the DifferentOddAndEvenPagesHeaderFooter property to true
section.PageSetup.DifferentOddAndEvenPagesHeaderFooter = true;

//Add odd header
Paragraph P3 = section.HeadersFooters.OddHeader.AddParagraph();

//Append text
TextRange OH = P3.AppendText("Odd Header");

//Set the HorizontalAlignment for the paragraph
P3.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

//Set the font name and font size
OH.CharacterFormat.FontName = "Arial";
OH.CharacterFormat.FontSize = 10;

//Add even header
Paragraph P4 = section.HeadersFooters.EvenHeader.AddParagraph();

//Append text
TextRange EH = P4.AppendText("Even Header from E-iceblue Using Spire.Doc");

//Set the HorizontalAlignment for the paragraph
P4.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

//Set the font name and font size
EH.CharacterFormat.FontName = "Arial";
EH.CharacterFormat.FontSize = 10;

//Add odd footer
Paragraph P2 = section.HeadersFooters.OddFooter.AddParagraph();

//Append text
TextRange OF = P2.AppendText("Odd Footer");

//Set the HorizontalAlignment for the paragraph
P2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

//Set the font name and font size
OF.CharacterFormat.FontName = "Arial";
OF.CharacterFormat.FontSize = 10;

//Add even footer
Paragraph P1 = section.HeadersFooters.EvenFooter.AddParagraph();

//Append text
TextRange EF = P1.AppendText("Even Footer from E-iceblue Using Spire.Doc");

//Set the font name and font size
EF.CharacterFormat.FontName = "Arial";
EF.CharacterFormat.FontSize = 10;

//Set the HorizontalAlignment for the paragraph
P1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
```

---

# spire.doc csharp page border
## configure page border surround for header and footer
```csharp
// Create a new document
Document doc = new Document();

// Add a section to the document
Section section = doc.AddSection();

// Set the page border properties
section.PageSetup.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Wave;
section.PageSetup.Borders.Color = Color.Green;
section.PageSetup.Borders.Left.Space = 20;
section.PageSetup.Borders.Right.Space = 20;

// Add a header paragraph to the section
Paragraph paragraph1 = section.HeadersFooters.Header.AddParagraph();

// Set horizontal alignment for the paragraph
paragraph1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

// Append text
TextRange headerText = paragraph1.AppendText("Header isn't included in page border");

// Set the character format for the text
headerText.CharacterFormat.FontName = "Calibri";
headerText.CharacterFormat.FontSize = 20;
headerText.CharacterFormat.Bold = true;

// Add a footer paragraph to the section
Paragraph paragraph2 = section.HeadersFooters.Footer.AddParagraph();

// Set horizontal alignment for the paragraph
paragraph2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

// Append text
TextRange footerText = paragraph2.AppendText("Footer is included in page border");

// Set the character format for the text
footerText.CharacterFormat.FontName = "Calibri";
footerText.CharacterFormat.FontSize = 20;
footerText.CharacterFormat.Bold = true;

// Configure page setup properties
section.PageSetup.PageBorderIncludeHeader = false;
section.PageSetup.HeaderDistance = 40;
section.PageSetup.PageBorderIncludeFooter = true;
section.PageSetup.FooterDistance = 40;
```

---

# spire.doc csharp remove footer
## remove footer from word document
```csharp
//Get the first section
Section section = doc.Sections[0];

//Clear footer in the first page
HeaderFooter footer;
footer = section.HeadersFooters[HeaderFooterType.FooterFirstPage];
if (footer != null)
    footer.ChildObjects.Clear();
//Clear footer in the odd page
footer = section.HeadersFooters[HeaderFooterType.FooterOdd];
if (footer != null)
    footer.ChildObjects.Clear();
//Clear footer in the even page
footer = section.HeadersFooters[HeaderFooterType.FooterEven];
if (footer != null)
    footer.ChildObjects.Clear();
```

---

# spire.doc csharp header
## remove headers from word document
```csharp
//Get the first section of the document
Section section = doc.Sections[0];

//Traverse the word document and clear all headers in different type
foreach (Paragraph para in section.Paragraphs)
{
    foreach (DocumentObject obj in para.ChildObjects)
    {
        //Clear header in the first page
        HeaderFooter header;
        header = section.HeadersFooters[HeaderFooterType.HeaderFirstPage];
        if (header != null)
            header.ChildObjects.Clear();
        //Clear header in the odd page
        header = section.HeadersFooters[HeaderFooterType.HeaderOdd];
        if (header != null)
            header.ChildObjects.Clear();
        //Clear header in the even page
        header = section.HeadersFooters[HeaderFooterType.HeaderEven];
        if (header != null)
            header.ChildObjects.Clear();
    }
}
```

---

# Spire.Doc C# Table Alternative Text
## Add title and description to a table as alternative text
```csharp
//Get the first section
Section section = doc.Sections[0];

//Get the first table in the section
Table table = section.Tables[0] as Table;

//Set the table title
table.Title = "Table 1";

//Add description
table.TableDescription = "Description Text";
```

---

# spire.doc csharp table operations
## add and delete rows in a word table
```csharp
//Get the first section
Section section = document.Sections[0];

//Get the first table
Table table = section.Tables[0] as Table;

//Delete the eighth row
table.Rows.RemoveAt(7);

//Add a row and insert it into specific position
TableRow row = new TableRow(document);
for (int i = 0; i < table.Rows[0].Cells.Count; i++)
{
    //Add a cell
    TableCell tc = row.AddCell();

    //Add a paragraph for the cell
    Paragraph paragraph = tc.AddParagraph();

    //Set horizontal alignment for the paragraph
    paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

    //Append text
    paragraph.AppendText("Added");
}

//Insert the new row
table.Rows.Insert(2, row);

//Add a row at the end of table
table.AddRow();
```

---

# Spire.Doc C# Table Column Operations
## Add or remove columns from a Word table
```csharp
private void AddColumn(Table table, int columnIndex)
{
    for (int r = 0; r < table.Rows.Count; r++)
    {
        //Create a new table cell
        TableCell addCell = new TableCell(table.Document);

        //Insert the new cell into the specified position
        table.Rows[r].Cells.Insert(columnIndex, addCell);
    }
}

private void RemoveColumn(Table table, int columnIndex)
{
    for (int r = 0; r < table.Rows.Count; r++)
    {
        //Remove the cell from specified position
        table.Rows[r].Cells.RemoveAt(columnIndex);
    }
}
```

---

# Spire.Doc C# Add Picture to Table Cell
## Insert an image into a specific cell of a Word table and set its dimensions
```csharp
//Get the first table from the first section of the document
Table table = (Table)doc.Sections[0].Tables[0];

//Add a picture to the specified table cell
DocPicture picture = table.Rows[1].Cells[2].Paragraphs[0].AppendPicture(Image.FromFile(imagePath));

//Set picture width
picture.Width = 100;

//Set picture height
picture.Height = 100;
```

---

# spire.doc csharp table
## create table from datatable
```csharp
//Add a table
Table table = section.AddTable(true);

//Set its width
table.PreferredWidth = new PreferredWidth(WidthType.Percentage, 100);

//Fill table with the data of datatable
FillTableUsingDataTable(table, dataTable);

//Set table style
table.Format.Paddings.All = 5;

for (int i = 0; i < table.FirstRow.Cells.Count; i++)
{
    table.FirstRow.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.CadetBlue;
}

private static void FillTableUsingDataTable(Table table, DataTable dataTable)
{
    //Get the count of the columns
    int columnCount = dataTable.Columns.Count;

    //Loop through the rows of data table
    foreach (DataRow dataRow in dataTable.Rows)
    {
        TableRow row = table.AddRow(columnCount);
        foreach (DataColumn dataColumn in dataTable.Columns)
        {
            //Get the column index
            int columnIndex = dataTable.Columns.IndexOf(dataColumn);

            //Get the value 
            string value = dataRow[dataColumn].ToString();

            //Get the cell object
            TableCell cell = row.Cells[columnIndex];
            //Add paragraph for cell
            Paragraph para = cell.AddParagraph();
            //Append text from datatable
            para.AppendText(value);
            //Set the alignment of cell
            cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
            para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
        }
    }
}
```

---

# Spire.Doc C# Table Formatting
## Allow table rows to break across pages
```csharp
//Get the first section
Section section = document.Sections[0];

//Get the first table
Table table = section.Tables[0] as Table;

//Loop through the table rows
foreach (TableRow row in table.Rows)
{
    //Allow break across pages
    row.RowFormat.IsBreakAcrossPages = true;
}
```

---

# spire.doc csharp table auto fit
## automatically fit table to contents in word document
```csharp
//Get the first section
Section section = document.Sections[0];

//Get the table from the section
Table table = section.Tables[0] as Table;

//Automatically fit the table to the cell content
table.AutoFit(AutoFitBehaviorType.AutoFitToContents);
```

---

# spire.doc csharp table
## auto fit table to fixed column widths
```csharp
//Get the first section
Section section = document.Sections[0];

//Get the first table
Table table = section.Tables[0] as Table;

//The table is set to a fixed size
table.AutoFit(AutoFitBehaviorType.FixedColumnWidths);
```

---

# Spire.Doc C# Table AutoFit
## Automatically fit a table to the window width in a Word document
```csharp
// Create a document
Document document = new Document();

// Get the first section
Section section = document.Sections[0];

// Get the first table
Table table = section.Tables[0] as Table;

// Automatically fit the table to the active window width
table.AutoFit(AutoFitBehaviorType.AutoFitToWindow);
```

---

# spire.doc csharp table cell merge status
## check the merge status of table cells in a Word document
```csharp
//Get the first section
Section section = doc.Sections[0];

//Get the first table in the section
Table table = section.Tables[0] as Table;

//Create a StringBuilder instance
StringBuilder stringBuidler = new StringBuilder();

//Loop through the table rows
for (int i = 0; i < table.Rows.Count; i++)
{
    //Get the table rows
    TableRow tableRow = table.Rows[i];

    //Loop through the cells of the row
    for (int j = 0; j < tableRow.Cells.Count; j++)
    {
        //Get each cell
        TableCell tableCell = tableRow.Cells[j];

        //Returns the way of vertical merging of the cell
        CellMerge verticalMerge = tableCell.CellFormat.VerticalMerge;

        //Get the status of cell merge 
        short horizontalMerge = tableCell.GridSpan;
        if (verticalMerge == CellMerge.None && horizontalMerge == 1)
        {
            stringBuidler.Append("Row " + i + ", cell " + j + ": ");
            stringBuidler.AppendLine("This cell isn't merged.");
        }
        else
        {
            stringBuidler.Append("Row " + i + ", cell " + j + ": ");
            stringBuidler.AppendLine("This cell is merged.");
        }
    }

    //Append an empty line
    stringBuidler.AppendLine();
}
```

---

# spire.doc csharp table row
## clone table row in word document
```csharp
//Get the first section
Section se = doc.Sections[0];

//Get the first row of the first table
TableRow firstRow = se.Tables[0].Rows[0];

//Copy the first row to clone_FirstRow via TableRow.clone()
TableRow clone_FirstRow = firstRow.Clone();

//Add a table row to collection
se.Tables[0].Rows.Add(clone_FirstRow);
```

---

# Spire.Doc C# Table Cloning
## Clone a table in a Word document and modify its content
```csharp
//Get the first section
Section se = doc.Sections[0];

//Get the first table
Table original_Table = (Table)se.Tables[0];

//Copy the existing table to copied_Table via Table.clone()
Table copied_Table = original_Table.Clone();
string[] st = new string[] { "Spire.Presentation for .Net", "A professional " +
    "PowerPoint® compatible library that enables developers to create, read, " +
    "write, modify, convert and Print PowerPoint documents on any .NET framework, " +
    ".NET Core platform." };
//Get the last row of table
TableRow lastRow = copied_Table.Rows[copied_Table.Rows.Count - 1];

//Change last row's data
for (int i = 0; i < lastRow.Cells.Count - 1; i++)
{
    lastRow.Cells[i].Paragraphs[0].Text = st[i];
}
//Add copied_Table to the section
se.Tables.Add(copied_Table);
```

---

# spire.doc csharp table operations
## combine and split tables in word document
```csharp
// Combine tables
//Get the first and second table
Table table1 = section.Tables[0] as Table;
Table table2 = section.Tables[1] as Table;

//Add the rows of table2 to table1
for (int i = 0; i < table2.Rows.Count; i++)
{
    table1.Rows.Add(table2.Rows[i].Clone());
}

//Remove the table2
section.Tables.Remove(table2);

// Split table
//Get the first table
Table table = section.Tables[0] as Table;

//We will split the table at the third row;
int splitIndex = 2;

//Create a new table for the split table
Table newTable = new Table(section.Document);

//Add rows to the new table
for (int i = splitIndex; i < table.Rows.Count; i++)
{
    newTable.Rows.Add(table.Rows[i].Clone());
}

//Remove rows from original table
for (int i = table.Rows.Count - 1; i >= splitIndex; i--)
{
    table.Rows.RemoveAt(i);
}

//Add the new table in section
section.Tables.Add(newTable);
```

---

# spire.doc csharp table
## create nested table in word document
```csharp
//Create a new document
Document doc = new Document();

//Add a new section
Section section = doc.AddSection();

//Add a table
Table table = section.AddTable(true);

//Set the number of rows and columns
table.ResetCells(2, 2);

//Set column width
table.Rows[0].Cells[0].SetCellWidth(70, CellWidthType.Point);
table.Rows[1].Cells[0].SetCellWidth(70, CellWidthType.Point);

//Determines how Microsoft Word resizes a table when the AutoFit feature is used
table.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

//Insert content to cells
table[0, 0].AddParagraph().AppendText("Spire.Doc for .NET");
string text = "Spire.Doc for .NET is a professional Word" +
".NET library specifically designed for developers to create," +
"read, write, convert and print Word document files from any .NET" +
"platform with fast and high quality performance.";
table[0, 1].AddParagraph().AppendText(text);

//Add a nested table to cell(first row, second column)
Table nestedTable = table[0, 1].AddTable(true);

//Set the number of rows and columns
nestedTable.ResetCells(4, 3);

//Determines how Microsoft Word resizes a table when the AutoFit feature is used
nestedTable.AutoFit(AutoFitBehaviorType.AutoFitToContents);

//Add content to nested cells
nestedTable[0, 0].AddParagraph().AppendText("NO.");
nestedTable[0, 1].AddParagraph().AppendText("Item");
nestedTable[0, 2].AddParagraph().AppendText("Price");

//Add content to nested cells
nestedTable[1, 0].AddParagraph().AppendText("1");
nestedTable[1, 1].AddParagraph().AppendText("Pro Edition");
nestedTable[1, 2].AddParagraph().AppendText("$799");

//Add content to nested cells
nestedTable[2, 0].AddParagraph().AppendText("2");
nestedTable[2, 1].AddParagraph().AppendText("Standard Edition");
nestedTable[2, 2].AddParagraph().AppendText("$599");

//Add content to nested cells
nestedTable[3, 0].AddParagraph().AppendText("3");
nestedTable[3, 1].AddParagraph().AppendText("Free Edition");
nestedTable[3, 2].AddParagraph().AppendText("$0");
```

---

# spire.doc csharp table
## create a formatted table in word document
```csharp
// Create a new document
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Define the table headers
String[] header = { "Name", "Capital", "Continent", "Area", "Population" };

// Sample data for the table
String[][] data =
    {
        new String[]{"Argentina", "Buenos Aires", "South America", "2777815", "32300003"},
        new String[]{"Bolivia", "La Paz", "South America", "1098575", "7300000"},
        new String[]{"Brazil", "Brasilia", "South America", "8511196", "150400000"}
    };

// Create a new table in the section
Spire.Doc.Table table = section.AddTable(true);
table.ResetCells(data.Length + 1, header.Length);

// Set the properties for the first row (header row)
TableRow headerRow = table.Rows[0];
headerRow.IsHeader = true;
headerRow.Height = 20;
headerRow.HeightType = TableRowHeightType.Exactly;
for (int i = 0; i < headerRow.Cells.Count; i++)
{
    headerRow.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.Gray;
}

// Populate the cells in the header row with the header values
for (int i = 0; i < header.Length; i++)
{
    headerRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
    Paragraph p = headerRow.Cells[i].AddParagraph();
    p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
    TextRange txtRange = p.AppendText(header[i]);
    txtRange.CharacterFormat.Bold = true;
}

// Populate the table rows with data
for (int r = 0; r < data.Length; r++)
{
    TableRow dataRow = table.Rows[r + 1];
    dataRow.Height = 20;
    dataRow.HeightType = TableRowHeightType.Exactly;
    for (int i = 0; i < dataRow.Cells.Count; i++)
    {
        dataRow.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.Empty;
    }

    // Populate the cells in the data rows with the corresponding data values
    for (int c = 0; c < data[r].Length; c++)
    {
        dataRow.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
        dataRow.Cells[c].AddParagraph().AppendText(data[r][c]);
    }
}

// Apply background color to alternate rows
for (int j = 1; j < table.Rows.Count; j++)
{
    if (j % 2 == 0)
    {
        TableRow row2 = table.Rows[j];
        for (int f = 0; f < row2.Cells.Count; f++)
        {
            row2.Cells[f].CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        }
    }
}
```

---

# spire.doc csharp table
## create table directly in word document
```csharp
// Create a new Document object
Document doc = new Document();

// Add a new section to the document
Section section = doc.AddSection();

// Create a new table with the document as its parent
Table table = new Table(doc);
table.ResetCells(1, 2);

// Set the preferred width of the table to 100% of the page width
table.PreferredWidth = new PreferredWidth(WidthType.Percentage, (short)100);

// Set the border type of the table to single line
table.Format.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;

// Create a new row for the table
TableRow row = table.Rows[0];

// Set the height of the row to 50.0f
row.Height = 50.0f; 

// Create the first cell of the row
TableCell cell1 = table.Rows[0].Cells[0];
Paragraph para1 = cell1.AddParagraph();
// Add text to the cell
para1.AppendText("Row 1, Cell 1"); 
// Set the horizontal alignment of the text
para1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center; 
// Set the background color of the cell
cell1.CellFormat.Shading.BackgroundPatternColor = Color.CadetBlue;
// Set the vertical alignment of the content in the cell
cell1.CellFormat.VerticalAlignment = VerticalAlignment.Middle; 

// Create the second cell of the row
TableCell cell2 = table.Rows[0].Cells[1];
Paragraph para2 = cell2.AddParagraph();
// Add text to the cell
para2.AppendText("Row 1, Cell 2"); 
// Set the horizontal alignment of the text
para2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
// Set the background color of the cell
cell2.CellFormat.Shading.BackgroundPatternColor = Color.CadetBlue; 
// Set the vertical alignment of the content in the cell
cell2.CellFormat.VerticalAlignment = VerticalAlignment.Middle; 

// Add the table to the section
section.Tables.Add(table);
```

---

# Spire.Doc C# HTML Table
## Create table from HTML in Word document
```csharp
// Create a new Document object
Document document = new Document();

// Add a new section to the document
Section section = document.AddSection();

// Append the HTML content to the section as a paragraph
section.AddParagraph().AppendHTML(HTML);
```

---

# Spire.Doc C# Vertical Table
## Create a vertical table in a Word document with rotated text
```csharp
// Create a new Document object
Document document = new Document();

// Add a new section to the document
Section section = document.AddSection();

// Add a table to the section
Table table = section.AddTable();
table.ResetCells(1, 1);

// Get the first cell of the table
TableCell cell = table.Rows[0].Cells[0];

// Set the height of the table row
table.Rows[0].Height = 150;

// Add a paragraph with text to the cell
cell.AddParagraph().AppendText("Draft copy in vertical style");

// Set the text direction of the cell to right-to-left rotated
cell.CellFormat.TextDirection = TextDirection.RightToLeftRotated;

// Enable wrap text around the table
table.Format.WrapTextAround = true;

// Set the vertical position of the table relative to the page
table.Format.Positioning.VertRelationTo = VerticalRelation.Page;

// Set the horizontal position of the table relative to the page
table.Format.Positioning.HorizRelationTo = HorizontalRelation.Page;

// Set the horizontal position of the table
table.Format.Positioning.HorizPosition = section.PageSetup.PageSize.Width - table.Width;

// Set the vertical position of the table
table.Format.Positioning.VertPosition = 200;
```

---

# Spire.Doc C# Table Borders
## Set different border styles for tables and cells in a Word document
```csharp
// Get the first table in the document's first section
Table table = document.Sections[0].Tables[0] as Table;

// Set borders for the entire table
setTableBorders(table);

// Set borders for a specific cell in the table
setCellBorders(table.Rows[2].Cells[0]);

private void setTableBorders(Table table)
{
    table.Format.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;
    table.Format.Borders.LineWidth = 3.0F;
    table.Format.Borders.Color = Color.Red;
}

private void setCellBorders(TableCell tableCell)
{
    tableCell.CellFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.DotDash;
    tableCell.CellFormat.Borders.LineWidth = 1.0F;
    tableCell.CellFormat.Borders.Color = Color.Green;
}
```

---

# Spire.Doc C# Table Formatting
## Format merged cells in a Word document table
```csharp
// Create a new ParagraphStyle and customize its formatting properties
ParagraphStyle style = new ParagraphStyle(document);
style.Name = "Style";
style.CharacterFormat.TextColor = Color.DeepSkyBlue;
style.CharacterFormat.Italic = true;
style.CharacterFormat.Bold = true;
style.CharacterFormat.FontSize = 13;
document.Styles.Add(style);

// Apply horizontal merge for the cells in the first row from column index 0 to 1
table.ApplyHorizontalMerge(0, 0, 1);

// Apply the custom style to the paragraph in the first cell of the first row
table.Rows[0].Cells[0].Paragraphs[0].ApplyStyle(style.Name);

// Set the vertical alignment and horizontal alignment of the first cell in the first row
table.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
table.Rows[0].Cells[0].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

// Apply vertical merge for the cells in the second row from row index 1 to 3
table.ApplyVerticalMerge(0, 1, 3);

// Apply the custom style to the paragraph in the first cell of the second row
table.Rows[1].Cells[0].Paragraphs[0].ApplyStyle(style.Name);

// Set the vertical alignment and horizontal alignment of the first cell in the second row
table.Rows[1].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
table.Rows[1].Cells[0].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

// Set the width of the first cell in the second row as a percentage of the table width
table.Rows[1].Cells[0].SetCellWidth(20, CellWidthType.Percentage);
```

---

# spire.doc csharp table border
## get diagonal border information from table cell
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Get the first table in the section
Table table = section.Tables[0] as Table;

// Get the DiagonalUp border type of cell (0,0) in the table
Spire.Doc.Documents.BorderStyle bs_UP = table[0, 0].CellFormat.Borders.DiagonalUp.BorderType;

// Get the DiagonalUp border color of cell (0,0) in the table
Color color_UP = table[0, 0].CellFormat.Borders.DiagonalUp.Color;

// Get the line width of the DiagonalUp border of cell (0,0) in the table
float width_UP = table[0, 0].CellFormat.Borders.DiagonalUp.LineWidth;

// Get the DiagonalDown border type of cell (0,0) in the table
Spire.Doc.Documents.BorderStyle bs_Down = table[0, 0].CellFormat.Borders.DiagonalDown.BorderType;

// Get the DiagonalDown border color of cell (0,0) in the table
Color color_Down = table[0, 0].CellFormat.Borders.DiagonalDown.Color;

// Get the line width of the DiagonalDown border of cell (0,0) in the table
float width_Down = table[0, 0].CellFormat.Borders.DiagonalDown.LineWidth;
```

---

# Spire.Doc C# Get Table Index
## Retrieve table, row, and cell indices from a Word document
```csharp
// Get the first section of the document
Section section = doc.Sections[0];

// Get the first table in the section
Table table = section.Tables[0] as Table;

// Get the collection of tables in the section
Spire.Doc.Collections.TableCollection collections = section.Tables;

// Get the index of the table in the collection
int tableIndex = collections.IndexOf(table);

// Get the last row in the table and its index
TableRow row = table.LastRow;
int rowIndex = row.GetRowIndex();

// Get the last cell in the row and its index
TableCell cell = row.LastChild as TableCell;
int cellIndex = cell.GetCellIndex();
```

---

# Spire.Doc C# Table Position
## Get table positioning information from a Word document
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Get the first table in the section
Table table = section.Tables[0] as Table;

// Check if text wrapping is enabled around the table
if (table.Format.WrapTextAround)
{
    // Get the positioning information for the table
    TablePositioning position = table.Format.Positioning;
    
    // Horizontal positioning information
    float horizPosition = position.HorizPosition;
    HorizontalPositionAbs horizPositionAbs = position.HorizPositionAbs;
    HorizontalRelation horizRelationTo = position.HorizRelationTo;
    
    // Vertical positioning information
    float vertPosition = position.VertPosition;
    VerticalPositionAbs vertPositionAbs = position.VertPositionAbs;
    VerticalRelation vertRelationTo = position.VertRelationTo;
    
    // Distance from surrounding text
    float distanceFromTop = position.DistanceFromTop;
    float distanceFromLeft = position.DistanceFromLeft;
    float distanceFromBottom = position.DistanceFromBottom;
    float distanceFromRight = position.DistanceFromRight;
}
```

---

# spire.doc csharp table operations
## merge and split table cells in word document
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Get the first table in the section
Table table = section.Tables[0] as Table;

// Apply horizontal merging to cells in the range (6, 2) to (6, 3)
table.ApplyHorizontalMerge(6, 2, 3);

// Apply vertical merging to cells in the range (2, 4) to (2, 5)
table.ApplyVerticalMerge(2, 4, 5);

// Split the cell at row index 8 and column index 3 into a 2x2 grid
table.Rows[8].Cells[3].SplitCell(2, 2);
```

---

# spire.doc csharp table formatting
## modify table, row, and cell formats in word documents
```csharp
// Modify the table format
private static void MoidfyTableFormat(Table table)
{
    // Set the preferred width of the table
    table.PreferredWidth = new PreferredWidth(WidthType.Twip, (short)6000);

    // Apply a specific table style to the table
    table.ApplyStyle(DefaultTableStyle.ColorfulGridAccent3);

    // Set padding for all cells in the table
    table.Format.Paddings.All = 5;

    // Set the title and description of the table
    table.Title = "Spire.Doc for .NET";
    table.TableDescription = "Spire.Doc for .NET is a professional Word .NET library";
}

// Modify the row format
private static void ModifyRowFormat(Table table)
{
    // Set the cell spacing of the first row
    table.Format.CellSpacing = 2;

    // Set the height of the second row
    table.Rows[1].HeightType = TableRowHeightType.Exactly;
    table.Rows[1].Height = 20f;

    // Set the background color of the third row
    for (int i = 0; i < table.Rows[2].Cells.Count; i++)
    {
        table.Rows[2].Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.DarkSeaGreen;
    }
}

// Modify the cell format
private static void ModifyCellFormat(Table table)
{
    // Set the vertical alignment and horizontal alignment of the first cell in the first row
    table.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
    table.Rows[0].Cells[0].Paragraphs[0].Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

    // Set the background color of the first cell in the second row
    table.Rows[1].Cells[0].CellFormat.Shading.BackgroundPatternColor = Color.DarkSeaGreen;

    // Set borders for the first cell in the third row
    table.Rows[2].Cells[0].CellFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Single;
    table.Rows[2].Cells[0].CellFormat.Borders.LineWidth = 1f;
    table.Rows[2].Cells[0].CellFormat.Borders.Left.Color = Color.Red;
    table.Rows[2].Cells[0].CellFormat.Borders.Right.Color = Color.Red;
    table.Rows[2].Cells[0].CellFormat.Borders.Top.Color = Color.Red;
    table.Rows[2].Cells[0].CellFormat.Borders.Bottom.Color = Color.Red;

    // Set the text direction of the first cell in the fourth row
    table.Rows[3].Cells[0].CellFormat.TextDirection = TextDirection.RightToLeft;
}
```

---

# spire.doc csharp table formatting
## prevent page breaks in word table
```csharp
// Get the first table in the first section of the document
Table table = document.Sections[0].Tables[0] as Table;

// Iterate through each row in the table
foreach (TableRow row in table.Rows)
{
    // Iterate through each cell in the row
    foreach (TableCell cell in row.Cells)
    {
        // Iterate through each paragraph in the cell
        foreach (Paragraph p in cell.Paragraphs)
        {
            // Set "Keep with next" property to true to prevent page breaks within paragraphs
            p.Format.KeepFollow = true;
        }
    }
}
```

---

# spire.doc csharp table operation
## remove table from word document
```csharp
// Remove the first Table
doc.Sections[0].Tables.RemoveAt(0);
```

---

# spire.doc csharp table
## repeat header rows on each page in word document
```csharp
// Create a new Word document
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Add a table to the section
Table table = section.AddTable(true);

// Set the preferred width of the table to 100%
PreferredWidth width = new PreferredWidth(WidthType.Percentage, 100);
table.PreferredWidth = width;

// Add a header row to the table
TableRow row = table.AddRow();
row.IsHeader = true;  // This makes the row repeat on each page
// Add a cell to the header row
TableCell cell = row.AddCell();
cell.SetCellWidth(100, CellWidthType.Percentage);

// Style the header row
for (int i = 0; i < row.Cells.Count; i++)
{
    row.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
}

// Add a paragraph to the cell with text "Row Header 1"
Paragraph paragraph = cell.AddParagraph();
paragraph.AppendText("Row Header 1");
paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

// Add another header row to the table
row = table.AddRow(false, 1);
row.IsHeader = true;  // This makes the row repeat on each page
for (int i = 0; i < row.Cells.Count; i++)
{
    row.Cells[i].CellFormat.Shading.BackgroundPatternColor = Color.Ivory;
}
row.Height = 30;
cell = row.Cells[0];
cell.SetCellWidth(100, CellWidthType.Percentage);
cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;

// Add a paragraph to the cell with text "Row Header 2"
paragraph = cell.AddParagraph();
paragraph.AppendText("Row Header 2");
paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

// Add rows and cells to the table
for (int i = 0; i < 70; i++)
{
    row = table.AddRow(false, 2);
    cell = row.Cells[0];
    cell.SetCellWidth(50, CellWidthType.Percentage);
    cell.AddParagraph().AppendText("Column 1 Text");
    cell = row.Cells[1];
    cell.SetCellWidth(50, CellWidthType.Percentage);
    cell.AddParagraph().AppendText("Column 2 Text");
}

// Set background color for alternating rows
for (int j = 1; j < table.Rows.Count; j++)
{
    if (j % 2 == 0)
    {
        TableRow row2 = table.Rows[j];
        for (int f = 0; f < row2.Cells.Count; f++)
        {
            row2.Cells[f].CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;         
        }
    }
}
```

---

# spire.doc csharp table text replacement
## replace text in word table using regex and string matching
```csharp
// Get the first section of the document
Section section = doc.Sections[0];

// Get the first table in the section
Table table = section.Tables[0] as Table;

// Create a regular expression pattern for matching text within curly braces
System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"{[^\}]+\}");

// Replace text in the table that matches the regular expression pattern with "E-iceblue"
table.Replace(regex, "E-iceblue");

// Replace the text "Beijing" with "Component" in the table, case-insensitive and match whole words only
table.Replace("Beijing", "Component", false, true);
```

---

# spire.doc csharp table column width
## set the width of a specific column in a word table using spire.doc library
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Get the first table in the section
Table table = section.Tables[0] as Table;

// Set the width of the first column in each row to 200 points
for (int i = 0; i < table.Rows.Count; i++)
{
    table.Rows[i].Cells[0].SetCellWidth(200, CellWidthType.Point);
}
```

---

# spire.doc csharp table positioning
## set table outside position in word document header
```csharp
// Create a new document object
Document doc = new Document();

// Add a section to the document
Section sec = doc.AddSection();

// Get the header of the first section in the document
HeaderFooter header = doc.Sections[0].HeadersFooters.Header;

// Add a paragraph to the header with left-aligned text
Paragraph paragraph = header.AddParagraph();
paragraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

// Add a table to the header
Table table = header.AddTable();
table.ResetCells(4, 2);

// Set table properties for text wrapping and positioning
table.Format.WrapTextAround = true;
table.Format.Positioning.HorizPositionAbs = HorizontalPosition.Outside;
table.Format.Positioning.VertRelationTo = VerticalRelation.Margin;
table.Format.Positioning.VertPosition = 43;

// Fill the table with data and set cell widths
for (int r = 0; r < 4; r++)
{
    TableRow dataRow = table.Rows[r];
    for (int c = 0; c < 2; c++)
    {
        if (c == 0)
        {
            // Add left-aligned text to the cell
            Paragraph par = dataRow.Cells[c].AddParagraph();
            par.AppendText("Sample Text");
            par.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
            dataRow.Cells[c].SetCellWidth(180, CellWidthType.Point);
        }
        else
        {
            // Add right-aligned text to the cell
            Paragraph par = dataRow.Cells[c].AddParagraph();
            par.AppendText("Sample Text");
            par.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
            dataRow.Cells[c].SetCellWidth(180, CellWidthType.Point);
        }
    }
}
```

---

# spire.doc csharp table style and border
## Set table style and customize borders in a Word document
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Get the first table in the section
Table table = section.Tables[0] as Table;

// Apply a predefined table style to the table
table.ApplyStyle(DefaultTableStyle.ColorfulList);

// Set the right border of the table
table.Format.Borders.Right.BorderType = Spire.Doc.Documents.BorderStyle.Hairline;
table.Format.Borders.Right.LineWidth = 1.0F;
table.Format.Borders.Right.Color = Color.Red;

// Set the top border of the table
table.Format.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.Hairline;
table.Format.Borders.Top.LineWidth = 1.0F;
table.Format.Borders.Top.Color = Color.Green;

// Set the left border of the table
table.Format.Borders.Left.BorderType = Spire.Doc.Documents.BorderStyle.Hairline;
table.Format.Borders.Left.LineWidth = 1.0F;
table.Format.Borders.Left.Color = Color.Yellow;

// Set the bottom border of the table
table.Format.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.DotDash;

// Set the vertical borders of the table
table.Format.Borders.Vertical.BorderType = Spire.Doc.Documents.BorderStyle.Dot;
table.Format.Borders.Vertical.Color = Color.Orange;

// Set the horizontal borders of the table to none
table.Format.Borders.Horizontal.BorderType = Spire.Doc.Documents.BorderStyle.None;
```

---

# spire.doc csharp table
## set vertical alignment in table cells
```csharp
// Add a table to the section with auto-fit behavior
Table table = section.AddTable(true);

// Reset the table cells to 3 rows and 3 columns
table.ResetCells(3, 3);

// Apply vertical merging to the first column of the table, spanning 3 rows
table.ApplyVerticalMerge(0, 0, 2);

// Set the vertical alignment of cells in the table
table.Rows[0].Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
table.Rows[0].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Top;
table.Rows[0].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Top;
table.Rows[1].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
table.Rows[1].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
table.Rows[2].Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Bottom;
table.Rows[2].Cells[2].CellFormat.VerticalAlignment = VerticalAlignment.Bottom;
```

---

# spire.doc csharp hyperlink
## create image hyperlink in document
```csharp
// Create a new Document object
Document doc = new Document();

// Get the first section of the document
Section section = doc.Sections[0];

// Add a new paragraph in the section
Paragraph paragraph = section.AddParagraph();

// Create a new DocPicture object with the loaded image
DocPicture picture = new DocPicture(doc);

// Load the image into the DocPicture object
picture.LoadImage(image);

// Append a hyperlink to the paragraph with the specified URL and the picture as the display element
paragraph.AppendHyperlink("https://www.example.com", picture, HyperlinkType.WebLink);
```

---

# spire.doc csharp hyperlinks
## find hyperlinks in document
```csharp
// Create a new Document object
Document doc = new Document();

// Load the document from the specified file path
doc.LoadFromFile(input);

// Create a list to store the hyperlinks and a variable to hold the text of the hyperlinks
List<Field> hyperlinks = new List<Field>();
string hyperlinksText = null;

// Iterate through the sections in the document
foreach (Section section in doc.Sections)
{
    // Iterate through the child objects in the body of the section
    foreach (DocumentObject sec in section.Body.ChildObjects)
    {
        // Check if the child object is a paragraph
        if (sec.DocumentObjectType == DocumentObjectType.Paragraph)
        {
            // Iterate through the child objects in the paragraph
            foreach (DocumentObject para in (sec as Paragraph).ChildObjects)
            {
                // Check if the child object is a field
                if (para.DocumentObjectType == DocumentObjectType.Field)
                {
                    // Cast the child object to a Field
                    Field field = para as Field;
                    
                    // Check if the field is a hyperlink
                    if (field.Type == FieldType.FieldHyperlink)
                    {
                        // Add the field to the list of hyperlinks
                        hyperlinks.Add(field);
                        
                        // Append the field's text to the hyperlinksText variable
                        hyperlinksText += field.FieldText + "\r\n";
                    }
                }
            }
        }
    }
}
```

---

# Spire.Doc C# Hyperlink
## Create and insert hyperlinks in a Word document
```csharp
// Create a new Document object
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Add a paragraph to the section, or get the first paragraph if it exists
Paragraph paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();

// Set the text content and apply a built-in style to the paragraph
paragraph.AppendText("Spire.Doc for .NET \r\n e-iceblue company Ltd. 2002-2010 All rights reserverd");
paragraph.ApplyStyle(BuiltinStyle.Heading2);

// Add a new paragraph to the section
paragraph = section.AddParagraph();

// Set the text content and apply a built-in style to the paragraph
paragraph.AppendText("Home page");
paragraph.ApplyStyle(BuiltinStyle.Heading2);

// Add a hyperlink to the paragraph with the specified URL and display text
paragraph = section.AddParagraph();
paragraph.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink);

// Add a new paragraph to the section
paragraph = section.AddParagraph();

// Set the text content and apply a built-in style to the paragraph
paragraph.AppendText("Contact US");
paragraph.ApplyStyle(BuiltinStyle.Heading2);

// Add a hyperlink to the paragraph with the specified email address and display text
paragraph = section.AddParagraph();
paragraph.AppendHyperlink("mailto:support@e-iceblue.com", "support@e-iceblue.com", HyperlinkType.EMailLink);

// Add a new paragraph to the section
paragraph = section.AddParagraph();

// Set the text content and apply a built-in style to the paragraph
paragraph.AppendText("Forum");
paragraph.ApplyStyle(BuiltinStyle.Heading2);

// Add a hyperlink to the paragraph with the specified URL and display text
paragraph = section.AddParagraph();
paragraph.AppendHyperlink("www.e-iceblue.com/forum/", "www.e-iceblue.com/forum/", HyperlinkType.WebLink);

// Add a new paragraph to the section
paragraph = section.AddParagraph();

// Set the text content and apply a built-in style to the paragraph
paragraph.AppendText("Download Link");
paragraph.ApplyStyle(BuiltinStyle.Heading2);

// Add a hyperlink to the paragraph with the specified URL and display text
paragraph = section.AddParagraph();
paragraph.AppendHyperlink("www.e-iceblue.com/Download/download-word-for-net-now.html", "www.e-iceblue.com/Download/download-word-for-net-now.html", HyperlinkType.WebLink);

// Add a new paragraph to the section
paragraph = section.AddParagraph();

// Set the text content and apply a built-in style to the paragraph
paragraph.AppendText("Insert Link On Image");
paragraph.ApplyStyle(BuiltinStyle.Heading2);

// Add an image to the paragraph and append a hyperlink to it with the specified URL and link type
paragraph = section.AddParagraph();
DocPicture picture = paragraph.AppendPicture(System.Drawing.Image.FromFile(@"..\..\..\..\..\..\Data\Spire.Doc.png"));
paragraph.AppendHyperlink("www.e-iceblue.com/Download/download-word-for-net-now.html", picture, HyperlinkType.WebLink);
```

---

# Spire.Doc C# Hyperlink Modification
## Find and modify hyperlink text in a Word document
```csharp
// Create a list to store the hyperlinks
List<Field> hyperlinks = new List<Field>();

// Iterate through the sections in the document
foreach (Section section in doc.Sections)
{
    // Iterate through the child objects in the body of the section
    foreach (DocumentObject sec in section.Body.ChildObjects)
    {
        // Check if the child object is a paragraph
        if (sec.DocumentObjectType == DocumentObjectType.Paragraph)
        {
            // Iterate through the child objects in the paragraph
            foreach (DocumentObject para in (sec as Paragraph).ChildObjects)
            {
                // Check if the child object is a field
                if (para.DocumentObjectType == DocumentObjectType.Field)
                {
                    // Cast the child object to a Field
                    Field field = para as Field;

                    // Check if the field is a hyperlink
                    if (field.Type == FieldType.FieldHyperlink)
                    {
                        // Add the field to the list of hyperlinks
                        hyperlinks.Add(field);
                    }
                }
            }
        }
    }
}

// Modify the text of the first hyperlink field
hyperlinks[0].FieldText = "Spire.Doc component";
```

---

# spire.doc csharp hyperlink removal
## Remove hyperlinks from Word document while preserving text
```csharp
// Method to find all hyperlinks in the document and return them as a list
private List<Field> FindAllHyperlinks(Document document)
{
    List<Field> hyperlinks = new List<Field>();

    foreach (Section section in document.Sections)
    {
        foreach (DocumentObject sec in section.Body.ChildObjects)
        {
            if (sec.DocumentObjectType == DocumentObjectType.Paragraph)
            {
                foreach (DocumentObject para in (sec as Paragraph).ChildObjects)
                {
                    if (para.DocumentObjectType == DocumentObjectType.Field)
                    {
                        Field field = para as Field;
                        if (field.Type == FieldType.FieldHyperlink)
                        {
                            hyperlinks.Add(field);
                        }
                    }
                }
            }
        }
    }
    return hyperlinks;
}

// Method to flatten a hyperlink, removing the hyperlink functionality but keeping the text
private void FlattenHyperlinks(Field field)
{
    // Store the indices of relevant objects for later removal
    int ownerParaIndex = field.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.OwnerParagraph);
    int fieldIndex = field.OwnerParagraph.ChildObjects.IndexOf(field);
    Paragraph sepOwnerPara = field.Separator.OwnerParagraph;
    int sepOwnerParaIndex = field.Separator.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.Separator.OwnerParagraph);
    int sepIndex = field.Separator.OwnerParagraph.ChildObjects.IndexOf(field.Separator);
    int endIndex = field.End.OwnerParagraph.ChildObjects.IndexOf(field.End);
    int endOwnerParaIndex = field.End.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.End.OwnerParagraph);

    // Format the text between the separator and the end of the field result
    FormatFieldResultText(field.Separator.OwnerParagraph.OwnerTextBody, sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex);

    // Remove the end field marker
    field.End.OwnerParagraph.ChildObjects.RemoveAt(endIndex);

    // Remove the field and its associated objects in reverse order
    for (int i = sepOwnerParaIndex; i >= ownerParaIndex; i--)
    {
        if (i == sepOwnerParaIndex && i == ownerParaIndex)
        {
            // Remove objects from the same paragraph as the field
            for (int j = sepIndex; j >= fieldIndex; j--)
            {
                field.OwnerParagraph.ChildObjects.RemoveAt(j);
            }
        }
        else if (i == ownerParaIndex)
        {
            // Remove objects from the field's paragraph but after the field
            for (int j = field.OwnerParagraph.ChildObjects.Count - 1; j >= fieldIndex; j--)
            {
                field.OwnerParagraph.ChildObjects.RemoveAt(j);
            }
        }
        else if (i == sepOwnerParaIndex)
        {
            // Remove objects from the separator's paragraph
            for (int j = sepIndex; j >= 0; j--)
            {
                sepOwnerPara.ChildObjects.RemoveAt(j);
            }
        }
        else
        {
            // Remove objects from other paragraphs
            field.OwnerParagraph.OwnerTextBody.ChildObjects.RemoveAt(i);
        }
    }
}

// Method to format the text between the separator and the end of a field result in the document body
private void FormatFieldResultText(Body ownerBody, int sepOwnerParaIndex, int endOwnerParaIndex, int sepIndex, int endIndex)
{
    for (int i = sepOwnerParaIndex; i <= endOwnerParaIndex; i++)
    {
        // Get the paragraph at the current index
        Paragraph para = ownerBody.ChildObjects[i] as Paragraph;
        
        if (i == sepOwnerParaIndex && i == endOwnerParaIndex)
        {
            // Format objects within the same paragraph as the separator and the end of the field
            for (int j = sepIndex + 1; j < endIndex; j++)
            {
                FormatText(para.ChildObjects[j] as TextRange);
            }
        }
        else if (i == sepOwnerParaIndex)
        {
            // Format objects after the separator in the separator's paragraph
            for (int j = sepIndex + 1; j < para.ChildObjects.Count; j++)
            {
                FormatText(para.ChildObjects[j] as TextRange);
            }
        }
        else if (i == endOwnerParaIndex)
        {
            // Format objects before the end of the field in the end paragraph
            for (int j = 0; j < endIndex; j++)
            {
                FormatText(para.ChildObjects[j] as TextRange);
            }
        }
        else
        {
            // Format all objects in other paragraphs
            for (int j = 0; j < para.ChildObjects.Count; j++)
            {
                FormatText(para.ChildObjects[j] as TextRange);
            }
        }
    }
}

// Method to format the text range by setting its color to black and removing underline style
private void FormatText(TextRange tr)
{
    tr.CharacterFormat.TextColor = Color.Black;
    tr.CharacterFormat.UnderlineStyle = UnderlineStyle.None;
}
```

---

# spire.doc csharp hyperlink
## set hyperlink format in word document
```csharp
// Create a new Document object
Document doc = new Document();

// Get the first section of the document
Section section = doc.Sections[0];

// Add a paragraph to the section and append regular text
Paragraph para1 = section.AddParagraph();
para1.AppendText("Regular Link: ");

// Append a hyperlink to the paragraph with the specified URL and display text
TextRange txtRange1 = para1.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink);
txtRange1.CharacterFormat.FontName = "Times New Roman";
txtRange1.CharacterFormat.FontSize = 12;

// Add a blank paragraph as separation
Paragraph blankPara1 = section.AddParagraph();

// Add another paragraph to the section and append text
Paragraph para2 = section.AddParagraph();
para2.AppendText("Change Color: ");

// Append a hyperlink to the paragraph with the specified URL and display text
TextRange txtRange2 = para2.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink);
txtRange2.CharacterFormat.FontName = "Times New Roman";
txtRange2.CharacterFormat.FontSize = 12;
txtRange2.CharacterFormat.TextColor = Color.Red;

// Add a blank paragraph as separation
Paragraph blankPara2 = section.AddParagraph();

// Add another paragraph to the section and append text
Paragraph para3 = section.AddParagraph();
para3.AppendText("Remove Underline: ");

// Append a hyperlink to the paragraph with the specified URL and display text
TextRange txtRange3 = para3.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.WebLink);
txtRange3.CharacterFormat.FontName = "Times New Roman";
txtRange3.CharacterFormat.FontSize = 12;
txtRange3.CharacterFormat.UnderlineStyle = UnderlineStyle.None;
```

---

# Spire.Doc Digital Signature
## Add digital signature to Word document
```csharp
// Create a new Document object
Document doc = new Document();

// Load the document from the specified file path
doc.LoadFromFile("document_path.doc");

// Save the document with digital signature using certificate and password
doc.SaveToFile("output_path.docx", FileFormat.Docx, "certificate_path.pfx", "certificate_password");

// Dispose the document object to free up resources
doc.Dispose();
```

---

# spire.doc digital signature verification
## check if Word document has digital signature
```csharp
bool hasDigitalSignature = Document.HasDigitalSignature(@"..\..\..\..\..\..\Data\CheckDigitalSignature.docx");

// Use a switch statement to determine the file format and update the fileFormat string accordingly
if(hasDigitalSignature)
{
    MessageBox.Show("This Word document has digital signature");
}else
{
    MessageBox.Show("This Word document has not digital signature");
}
```

---

# spire.doc csharp decrypt
## Decrypt password-protected Word document
```csharp
// Create a new Document object
Document document = new Document();

// Load the document from the specified file path using the provided password
document.LoadFromFile("TemplateWithPassword.docx", FileFormat.Docx, "E-iceblue");

// Save the document to the specified file path in DOCX format
document.SaveToFile("Sample.docx", FileFormat.Docx);
```

---

# Spire.Doc C# Document Encryption Check
## Determine if a Word document is encrypted using Spire.Doc library
```csharp
// Check if document is encrypted
bool isEncrypted = Document.IsEncrypted(@"..\..\..\..\..\..\Data\TemplateWithPassword.docx");
if(isEncrypted == true)
{
    MessageBox.Show("This document is encrypted. ");
}
else
{
    MessageBox.Show("This document is unencrypted. ");
}
```

---

# Spire.Doc Document Encryption
## Encrypt a Word document with password protection
```csharp
// Create a new Document object
Document document = new Document();

// Encrypt the document with the provided password
document.Encrypt("E-iceblue");
```

---

# spire.doc csharp security
## Lock specified sections in a document
```csharp
// Add two sections to the document
Section s1 = document.AddSection();
Section s2 = document.AddSection();

// Protect the document with a password and allow only form fields
document.Protect(ProtectionType.AllowOnlyFormFields, "123");

// Disable form field protection for section 2
s2.ProtectForm = false;
```

---

# spire.doc csharp security
## remove editable ranges from word document
```csharp
// Iterate through each section in the document
foreach (Section section in document.Sections)
{
    // Iterate through each paragraph in the section's body
    foreach (Paragraph paragraph in section.Body.Paragraphs)
    {
        // Loop through the child objects of the paragraph
        for (int i = 0; i < paragraph.ChildObjects.Count; )
        {
            DocumentObject obj = paragraph.ChildObjects[i];
            
            // Check if the child object is a PermissionStart or PermissionEnd element
            if (obj is PermissionStart || obj is PermissionEnd)
            {
                // Remove the PermissionStart or PermissionEnd element from the paragraph
                paragraph.ChildObjects.Remove(obj);
            }
            else
            {
                // Move to the next child object
                i++;
            }
        }
    }
}
```

---

# Spire.Doc C# Security
## Remove read-only restriction from Word document
```csharp
// Create a new Document object
Document document = new Document();

// Load the Word document file
document.LoadFromFile("RemoveReadOnlyRestriction.docx");

// Remove the read-only restriction from the document
document.Protect(ProtectionType.NoProtection);
```

---

# spire.doc csharp security
## set editable range in word document
```csharp
// Create a new Document object
Document document = new Document();

// Set the document protection to allow only reading with a password
document.Protect(ProtectionType.AllowOnlyReading, "password");

// Create a PermissionStart object to mark the start of an editable range with a specific ID
PermissionStart start = new PermissionStart(document, "testID");
// Create a PermissionEnd object to mark the end of the editable range with the same ID
PermissionEnd end = new PermissionEnd(document, "testID");

// Insert the PermissionStart object at the beginning of the first paragraph in the first section
document.Sections[0].Paragraphs[0].ChildObjects.Insert(0, start);
// Add the PermissionEnd object to the end of the first paragraph in the first section
document.Sections[0].Paragraphs[0].ChildObjects.Add(end);

// Dispose the Document object to free resources
document.Dispose();
```

---

# Spire.Doc Document Protection
## Set document protection type with password
```csharp
// Set the document protection to allow only reading with the specified password
document.Protect(ProtectionType.AllowOnlyReading, "123456");
```

---

# Spire.Doc Word to PDF Encryption
## Convert Word document to encrypted PDF
```csharp
// Create a new Document object
Document document = new Document();

// Load the Word document file from the specified path
document.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Docx_2.docx");

// Create a ToPdfParameterList object to specify PDF conversion parameters
ToPdfParameterList toPdf = new ToPdfParameterList();

// Encrypt the PDF with the specified password "e-iceblue"
toPdf.PdfSecurity.Encrypt("e-iceblue");

// Specify the output file name for the converted PDF
String result = "Result-WordToPdfEncrypt.pdf";

// Save the document as a PDF with the specified encryption settings
document.SaveToFile(result, toPdf);

// Dispose the Document object to free resources
document.Dispose();
```

---

# spire.doc csharp tc field
## add Table of Contents entry field to document
```csharp
// Create a new Document object
Document document = new Document();

// Add a new section to the document
Section section = document.AddSection();

// Add a new paragraph to the section
Paragraph paragraph = section.AddParagraph();

// Append a TC (Table of Contents) field to the paragraph with the specified entry text
Field field = paragraph.AppendField("TC", FieldType.FieldTOCEntry);
field.Code = @"TC " + "\"Entry Text\"" + " \\f" + " t";
```

---

# spire.doc csharp convert equation
## convert equation fields to OfficeMath objects in Word document
```csharp
// Get the first paragraph of the first section in the document
Paragraph paragraph = document.Sections[0].Paragraphs[0];

// Iterate through the child objects of the paragraph
for (int i = 0; i < paragraph.ChildObjects.Count; i++)
{
    // Get the current document object
    DocumentObject documentObject = paragraph.ChildObjects[i];

    // Check if the document object is a field of type Equation
    if (documentObject is Field && ((Field)documentObject).Type == FieldType.FieldEquation)
    {
        // Convert the field to an OfficeMath object
        OfficeMath officeMath = OfficeMath.FromEqField((Field)documentObject);

        // If conversion is successful, replace the field with the OfficeMath object
        if (officeMath != null)
        {
            paragraph.ChildObjects.Remove(documentObject);
            paragraph.ChildObjects.Insert(i, officeMath);
        }
    }
}
```

---

# spire.doc csharp field conversion
## convert form fields to body text in word document
```csharp
// Iterate through each form field in the first section of the document's body
foreach (FormField field in sourceDocument.Sections[0].Body.FormFields)
{
    // Check if the form field is of type FieldFormTextInput
    if (field.Type == FieldType.FieldFormTextInput)
    {
        // Get the owner paragraph of the form field
        Paragraph paragraph = field.OwnerParagraph;

        // Initialize variables for start and end index of bookmark objects
        int startIndex = 0;
        int endIndex = 0;

        // Create a TextRange object using the source document
        TextRange textRange = new TextRange(sourceDocument);

        // Set the text of the TextRange to the text of the paragraph
        textRange.Text = paragraph.Text;

        // Iterate through each child object of the paragraph
        foreach (DocumentObject obj in paragraph.ChildObjects)
        {
            // Check if the child object is a BookmarkStart object
            if (obj.DocumentObjectType == DocumentObjectType.BookmarkStart)
            {
                // Store the index of the BookmarkStart object
                startIndex = paragraph.ChildObjects.IndexOf(obj);
            }

            // Check if the child object is a BookmarkEnd object
            if (obj.DocumentObjectType == DocumentObjectType.BookmarkEnd)
            {
                // Store the index of the BookmarkEnd object
                endIndex = paragraph.ChildObjects.IndexOf(obj);
            }
        }

        // Remove the form fields or child objects between the start and end index
        for (int i = endIndex; i > startIndex; i--)
        {
            if (paragraph.ChildObjects[i] is TextFormField)
            {
                // Remove the TextFormField object
                TextFormField textFormField = paragraph.ChildObjects[i] as TextFormField;
                paragraph.ChildObjects.Remove(textFormField);
            }
            else
            {
                // Remove other child objects
                paragraph.ChildObjects.RemoveAt(i);
            }
        }

        // Insert the modified TextRange at the start index of the paragraph
        paragraph.ChildObjects.Insert(startIndex, textRange);

        // Exit the loop after processing the first FieldFormTextInput
        break;
    }
}
```

---

# spire.doc csharp field conversion
## convert document fields to text
```csharp
// Get the collection of fields in the document
FieldCollection fields = document.Fields;
int count = fields.Count;

// Iterate through each field in the collection
for (int i = 0; i < count; i++)
{
    // Get the first field in the collection
    Field field = fields[0];

    // Get the text of the field
    string s = field.FieldText;

    // Get the index of the field within its owner paragraph
    int index = field.OwnerParagraph.ChildObjects.IndexOf(field);

    // Create a TextRange object with the document and set its text to the field text
    TextRange textRange = new TextRange(document);
    textRange.Text = s;
    
    // Set the font size of the text range
    textRange.CharacterFormat.FontSize = 24f;

    // Insert the text range at the index of the field within its owner paragraph
    field.OwnerParagraph.ChildObjects.Insert(index, textRange);

    // Remove the field from its owner paragraph
    field.OwnerParagraph.ChildObjects.Remove(field);
}
```

---

# spire.doc csharp field conversion
## convert IF fields to text in a Word document
```csharp
// Get the collection of fields in the document
FieldCollection fields = document.Fields;

// Iterate through each field in the collection
for (int i = 0; i < fields.Count; i++)
{
    // Get the current field
    Field field = fields[i];

    // Check if the field is of type FieldIf
    if (field.Type == FieldType.FieldIf)
    {
        // Cast the field as TextRange to access its properties
        TextRange original = field as TextRange;

        // Get the text of the field
        string text = field.FieldText;

        // Create a new TextRange object with the document and set its text to the field text
        TextRange textRange = new TextRange(document);
        textRange.Text = text;

        // Set the font name and size of the new text range to match the original field
        textRange.CharacterFormat.FontName = original.CharacterFormat.FontName;
        textRange.CharacterFormat.FontSize = original.CharacterFormat.FontSize;

        // Get the owner paragraph of the field
        Paragraph par = field.OwnerParagraph;

        // Get the index of the field within its owner paragraph
        int index = par.ChildObjects.IndexOf(field);

        // Remove the field from its owner paragraph
        par.ChildObjects.RemoveAt(index);

        // Insert the new text range at the index of the field within its owner paragraph
        par.ChildObjects.Insert(index, textRange);
    }
}
```

---

# spire.doc csharp cross reference
## create cross reference to bookmark in word document
```csharp
// Create a new Document object
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Add a paragraph to the section and append a bookmark with the specified name
Paragraph paragraph = section.AddParagraph();
paragraph.AppendBookmarkStart("MyBookmark");
paragraph.AppendText("Text inside a bookmark");
paragraph.AppendBookmarkEnd("MyBookmark");

// Add line breaks to the paragraph
for (int i = 0; i < 4; i++)
{
    paragraph.AppendBreak(BreakType.LineBreak);
}

// Create a new Field object for referencing the bookmark
Field field = new Field(document);
field.Type = FieldType.FieldRef;
field.Code = @"REF MyBookmark \p \h";

// Add a new paragraph to the section and append text and the field
paragraph = section.AddParagraph();
paragraph.AppendText("For more information, see ");
paragraph.ChildObjects.Add(field);

// Add a field separator to the paragraph
FieldMark fieldSeparator = new FieldMark(document, FieldMarkType.FieldSeparator);
paragraph.ChildObjects.Add(fieldSeparator);

// Create a TextRange object and set its text
TextRange tr = new TextRange(document);
tr.Text = "above";
paragraph.ChildObjects.Add(tr);

// Add a field end mark to the paragraph
FieldMark fieldEnd = new FieldMark(document, FieldMarkType.FieldEnd);
paragraph.ChildObjects.Add(fieldEnd);
```

---

# spire.doc csharp form fields
## create form fields in word document
```csharp
private void AddForm(Section section)
{
    // Create a paragraph style for description texts
    ParagraphStyle descriptionStyle = new ParagraphStyle(section.Document);
    descriptionStyle.Name = "description";
    descriptionStyle.CharacterFormat.FontSize = 12;
    descriptionStyle.CharacterFormat.FontName = "Arial";
    descriptionStyle.CharacterFormat.TextColor = Color.FromArgb(0x00, 0x45, 0x8e);
    section.Document.Styles.Add(descriptionStyle);

    // Add the first description paragraph
    Paragraph p1 = section.AddParagraph();
    String text1 = "So that we can verify your identity and find your information, "
        + "please provide us with the following information. "
        + "This information will be used to create your online account. "
        + "Your information is not public, shared in any way, or displayed on this site";
    p1.AppendText(text1);
    p1.ApplyStyle(descriptionStyle.Name);

    // Add the second description paragraph
    Paragraph p2 = section.AddParagraph();
    String text2 = "You must provide a real email address to which we will send your password.";
    p2.AppendText(text2);
    p2.ApplyStyle(descriptionStyle.Name);
    p2.Format.AfterSpacing = 8;

    // Create a paragraph style for form field group labels
    ParagraphStyle formFieldGroupLabelStyle = new ParagraphStyle(section.Document);
    formFieldGroupLabelStyle.Name = "formFieldGroupLabel";
    formFieldGroupLabelStyle.ApplyBaseStyle("description");
    formFieldGroupLabelStyle.CharacterFormat.Bold = true;
    formFieldGroupLabelStyle.CharacterFormat.TextColor = Color.White;
    section.Document.Styles.Add(formFieldGroupLabelStyle);

    // Create a paragraph style for form field labels
    ParagraphStyle formFieldLabelStyle = new ParagraphStyle(section.Document);
    formFieldLabelStyle.Name = "formFieldLabel";
    formFieldLabelStyle.ApplyBaseStyle("description");
    formFieldLabelStyle.ParagraphFormat.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
    section.Document.Styles.Add(formFieldLabelStyle);

    // Add a table to the section for the form fields
    Table table = section.AddTable();
    // Set the number of columns
    table.DefaultColumnsNumber = 2; 
    // Set the default row height
    table.DefaultRowHeight = 20; 

    // Read the XML file containing the form structure
    using (Stream stream = File.OpenRead(@"..\..\..\..\..\..\Data\Form.xml"))
    {
        XPathDocument xpathDoc = new XPathDocument(stream);
        XPathNodeIterator sectionNodes = xpathDoc.CreateNavigator().Select("/form/section");

        // Iterate over each section node in the XML file
        foreach (XPathNavigator node in sectionNodes)
        {
            // Add a row for the form field group label
            TableRow row = table.AddRow(false);
            row.Cells[0].CellFormat.Shading.BackgroundPatternColor= Color.FromArgb(0x00, 0x71, 0xb6);
            row.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

            // Add the form field group label text to the cell
            Paragraph cellParagraph = row.Cells[0].AddParagraph();
            cellParagraph.AppendText(node.GetAttribute("name", ""));
            cellParagraph.ApplyStyle(formFieldGroupLabelStyle.Name);

            // Iterate over each field node within the section node
            XPathNodeIterator fieldNodes = node.Select("field");
            foreach (XPathNavigator fieldNode in fieldNodes)
            {
                // Add a row for the form field label and input field
                TableRow fieldRow = table.AddRow(false);

                // Set vertical alignment for the cells in the field row
                fieldRow.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                fieldRow.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                // Add the form field label to the first cell in the row
                Paragraph labelParagraph = fieldRow.Cells[0].AddParagraph();
                labelParagraph.AppendText(fieldNode.GetAttribute("label", ""));
                labelParagraph.ApplyStyle(formFieldLabelStyle.Name);

                // Add the input field paragraph to the second cell in the row
                Paragraph fieldParagraph = fieldRow.Cells[1].AddParagraph();
                String fieldId = fieldNode.GetAttribute("id", "");
                switch (fieldNode.GetAttribute("type", ""))
                {
                    case "text":
                        // Add a text form input field
                        TextFormField field = fieldParagraph.AppendField(fieldId, FieldType.FieldFormTextInput) as TextFormField;
                        field.DefaultText = "";
                        field.Text = "";
                        break;

                    case "list":
                        // Add a dropdown list form field
                        DropDownFormField list
                            = fieldParagraph.AppendField(fieldId, FieldType.FieldFormDropDown) as DropDownFormField;

                        XPathNodeIterator itemNodes = fieldNode.Select("item");
                        foreach (XPathNavigator itemNode in itemNodes)
                        {
                            list.DropDownItems.Add(itemNode.SelectSingleNode("text()").Value);
                        }
                        break;

                    case "checkbox":
                        // Add a checkbox form field
                        fieldParagraph.AppendField(fieldId, FieldType.FieldFormCheckBox);
                        break;
                }
            }

            table.ApplyHorizontalMerge(row.GetRowIndex(), 0, 1);
        }
    }
}
```

---

# spire.doc csharp field
## create IF field in document
```csharp
// Create a new document
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Add a paragraph to the section
Paragraph paragraph = section.AddParagraph();

// Create an IF field and add it to the paragraph
CreateIfField(document, paragraph);

// Method to create an IF field
static void CreateIfField(Document document, Paragraph paragraph)
{
    // Create a new IF field
    IfField ifField = new IfField(document);
    ifField.Type = FieldType.FieldIf;
    ifField.Code = "IF ";

    // Add the IF field to the paragraph
    paragraph.Items.Add(ifField);

    // Add the merge field and condition to the paragraph
    paragraph.AppendField("Count", FieldType.FieldMergeField);
    paragraph.AppendText(" > ");
    paragraph.AppendText("\"100\" ");
    paragraph.AppendText("\"Thanks\" ");
    paragraph.AppendText("\"The minimum order is 100 units\"");

    // Create the end mark of the IF field and add it to the paragraph
    IParagraphBase end = document.CreateParagraphItem(ParagraphItemType.FieldMark);
    (end as FieldMark).Type = FieldMarkType.FieldEnd;
    paragraph.Items.Add(end);

    // Set the end mark of the IF field
    ifField.End = end as FieldMark;
}
```

---

# spire.doc csharp nested fields
## create nested IF fields in a Word document
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Add a paragraph to the section
Paragraph paragraph = section.AddParagraph();

// Create the outer IF field and add it to the paragraph
IfField ifField = new IfField(document);
ifField.Type = FieldType.FieldIf;
ifField.Code = "IF ";
paragraph.Items.Add(ifField);

// Create the inner IF field and add it to the paragraph
IfField ifField2 = new IfField(document);
ifField2.Type = FieldType.FieldIf;
ifField2.Code = "IF ";
paragraph.ChildObjects.Add(ifField2);
paragraph.Items.Add(ifField2);
paragraph.AppendText("\"200\" < \"50\"   \"200\" \"50\" ");

// Create the end mark for the inner IF field and add it to the paragraph
IParagraphBase embeddedEnd = document.CreateParagraphItem(ParagraphItemType.FieldMark);
(embeddedEnd as FieldMark).Type = FieldMarkType.FieldEnd;
paragraph.Items.Add(embeddedEnd);
ifField2.End = embeddedEnd as FieldMark;

// Append additional text and create the end mark for the outer IF field
paragraph.AppendText(" > ");
paragraph.AppendText("\"100\" ");
paragraph.AppendText("\"Thanks\" ");
paragraph.AppendText("\"The minimum order is 100 units\"");
IParagraphBase end = document.CreateParagraphItem(ParagraphItemType.FieldMark);
(end as FieldMark).Type = FieldMarkType.FieldEnd;
paragraph.Items.Add(end);
ifField.End = end as FieldMark;

// Enable field update
document.IsUpdateFields = true;
```

---

# spire.doc csharp form fields
## fill form fields in word document with xml data
```csharp
// Select the "user" node from the XML document
XPathNavigator user = xpathDoc.CreateNavigator().SelectSingleNode("/user");

// Iterate through each form field in the document's first section
foreach (FormField field in document.Sections[0].Body.FormFields)
{
    // Get the XPath to retrieve the value for the current form field
    String path = String.Format("{0}/text()", field.Name);

    // Select the corresponding node from the XML document
    XPathNavigator propertyNode = user.SelectSingleNode(path);

    // If the node exists, set the value of the form field based on its type
    if (propertyNode != null)
    {
        switch (field.Type)
        {
            // Text input field
            case FieldType.FieldFormTextInput:
                field.Text = propertyNode.Value;
                break;

            // Dropdown field
            case FieldType.FieldFormDropDown:
                DropDownFormField combox = field as DropDownFormField;
                for (int i = 0; i < combox.DropDownItems.Count; i++)
                {
                    if (combox.DropDownItems[i].Text == propertyNode.Value)
                    {
                        combox.DropDownSelectedIndex = i;
                        break;
                    }
                    if (field.Name == "country" && combox.DropDownItems[i].Text == "Others")
                    {
                        combox.DropDownSelectedIndex = i;
                    }
                }
                break;

            // Checkbox field
            case FieldType.FieldFormCheckBox:
                if (Convert.ToBoolean(propertyNode.Value))
                {
                    CheckBoxFormField checkBox = field as CheckBoxFormField;
                    checkBox.Checked = true;
                }
                break;
        }
    }
}
```

---

# spire.doc csharp form fields
## modify form field properties in word document
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Get the second form field in the section
FormField formField = section.Body.FormFields[1];

// Check if the form field is a text input field
if (formField.Type == FieldType.FieldFormTextInput)
{
    // Set the text of the form field
    formField.Text = "My name is " + formField.Name;

    // Customize the text formatting of the form field
    formField.CharacterFormat.TextColor = Color.Red;
    formField.CharacterFormat.Italic = true;
}
```

---

# Spire.Doc C# Field Text Extraction
## Extract text from fields in a Word document
```csharp
// Load the document
Document document = new Document("document.docx");

// Get the collection of fields in the document
FieldCollection fields = document.Fields;

// Iterate through each field in the collection
foreach (Field field in fields)
{
    // Get the text of the field
    string fieldText = field.FieldText;
    
    // Process the field text as needed
}

// Dispose the document object
document.Dispose();
```

---

# spire.doc csharp form fields
## get form field by name from document
```csharp
// Load the document from a file
Document document = new Document(@"..\..\..\..\..\..\Data\FillFormField.doc");

// Get the first section of the document
Section section = document.Sections[0];

// Get the form field with the name "email"
FormField formField = section.Body.FormFields["email"];

// Access the name and type of the form field
string fieldName = formField.Name;
string fieldType = formField.FormFieldType.ToString();
```

---

# Spire.Doc C# Form Fields
## Get form fields collection from document
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Get the collection of form fields in the section
FormFieldCollection formFields = section.Body.FormFields;

// Get the count of form fields in the section
int fieldCount = formFields.Count;
```

---

# Spire.Doc C# Get Merge Field Names
## Extract and display merge field names from a Word document
```csharp
// Create a StringBuilder to hold the field information
StringBuilder sb = new StringBuilder();

// Load the document from a file
Document document = new Document(@"..\..\..\..\..\..\Data\MailMerge.doc");

// Get the array of merge field names in the document
string[] fieldNames = document.MailMerge.GetMergeFieldNames();

// Append the count of merge fields in the document to the StringBuilder
sb.Append("The document has " + fieldNames.Length.ToString() + " merge fields.");

// Append a header for the merge field names
sb.Append(" The below is the name of the merge field:" + "\r\n");

// Iterate through each merge field name and append it to the StringBuilder
foreach (string name in fieldNames)
{
    sb.AppendLine(name);
}

// Write the result to a text file
File.WriteAllText("result.txt", sb.ToString());

// Dispose the document object
document.Dispose();
```

---

# spire.doc csharp askfield
## handle AskField events in Word document
```csharp
// Subscribe to the UpdateFields event
doc.UpdateFields += new UpdateFieldsHandler(doc_UpdateFields);

// Enable field update
doc.IsUpdateFields = true;

// Event handler for updating fields
private static void doc_UpdateFields(object sender, IFieldsEventArgs args)
{     
    // Check if the event arguments are of type AskFieldEventArgs
    if (args is AskFieldEventArgs)
    {
        AskFieldEventArgs askArgs = args as AskFieldEventArgs;
        
        // Handle different bookmarks and set response text accordingly
        if (askArgs.BookmarkName == "1")
        {
            askArgs.ResponseText = "Thank you. This is my first time to come to a Chinese restaurant. Could you tell me the different features of Chinese food?";
        }
        
        if (askArgs.BookmarkName == "2")
        {
            askArgs.ResponseText = "No more, thank you. I'm quite full.";
        }
    }
}
```

---

# Spire.Doc C# Address Block Field
## Insert address block field into Word document
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Add a new paragraph to the section
Paragraph par = section.AddParagraph();

// Append a field with type "AddressBlock" to the paragraph
Field field = par.AppendField("ADDRESSBLOCK", FieldType.FieldAddressBlock);

// Set the code for the field, including additional options and formatting
field.Code = "ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\"";
```

---

# Spire.Doc C# Field Operations
## Insert an advance field in a Word document
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Add a paragraph to the section
Paragraph par = section.AddParagraph();

// Append a field with the specified type and text
Field field = par.AppendField("Field", FieldType.FieldAdvance);

// Set the code for the field using the specified parameters
field.Code = "ADVANCE \\d 10 \\l 10 \\r 10 \\u 0 \\x 100 \\y 100 ";

// Enable the automatic update of fields in the document
document.IsUpdateFields = true;
```

---

# Spire.Doc C# Merge Field
## Insert a merge field into a Word document
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Add a paragraph to the section
Paragraph par = section.AddParagraph();

// Append a merge field with the specified name and type
MergeField field = par.AppendField("MyFieldName", FieldType.FieldMergeField) as MergeField;
```

---

# spire.doc csharp field
## insert a none field into word document
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Add a paragraph to the section
Paragraph par = section.AddParagraph();

// Append an empty field with no specific type
Field field = par.AppendField(String.Empty, FieldType.FieldNone);
```

---

# spire.doc csharp page reference field
## insert page reference field into word document
```csharp
// Get the last section of the document
Section section = document.LastSection;

// Add a paragraph to the section
Paragraph par = section.AddParagraph();

// Append a page reference field with the specified name and type
Field field = par.AppendField("pageRef", FieldType.FieldPageRef);

// Set the code for the field with the specified parameters
field.Code = "PAGEREF bookmark1 \\# \"0\" \\* Arabic  \\* MERGEFORMAT";

// Enable the automatic update of fields in the document
document.IsUpdateFields = true;
```

---

# spire.doc csharp fields
## remove custom property fields from word document
```csharp
// Create a new document object
Document document = new Document();

// Get the collection of custom document properties
CustomDocumentProperties cdp = document.CustomDocumentProperties;

// Iterate through the custom document properties and remove them
for (int i = 0; i < cdp.Count; )
{
    cdp.Remove(cdp[i].Name);
}

// Enable the automatic update of fields in the document
document.IsUpdateFields = true;
```

---

# spire.doc csharp field removal
## remove field from word document
```csharp
// Get the first field in the document
Field field = document.Fields[0];

// Get the parent paragraph of the field
Paragraph par = field.OwnerParagraph;

// Get the index of the field within the child objects of the paragraph
int index = par.ChildObjects.IndexOf(field);

// Remove the field from the paragraph
par.ChildObjects.RemoveAt(index);
```

---

# spire.doc csharp text replacement
## replace text with merge field in word document
```csharp
// Find the text "Test" in the document
TextSelection ts = document.FindString("Test", true, true);

// Get the selected text as a single range
TextRange tr = ts.GetAsOneRange();

// Get the paragraph that contains the selected text
Paragraph par = tr.OwnerParagraph;

// Get the index of the selected text within its parent paragraph
int index = par.ChildObjects.IndexOf(tr);

// Create a new merge field
MergeField field = new MergeField(document);
field.FieldName = "MergeField";

// Insert the merge field at the same position as the selected text
par.ChildObjects.Insert(index, field);

// Remove the selected text from the paragraph
par.ChildObjects.Remove(tr);
```

---

# Spire.Doc Field Culture Setting
## Set culture for date fields in Word document
```csharp
// Create a new document
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Add a paragraph to the section
Paragraph paragraph = section.AddParagraph();

// Add text to the paragraph
paragraph.AppendText("Add Date Field: ");

// Append a date field to the paragraph and set its format
Field field1 = paragraph.AppendField("Date1", FieldType.FieldDate) as Field;
field1.Code = @"DATE  \@" + "\"yyyy\\MM\\dd\"";

// Add a new paragraph to the section
Paragraph newParagraph = section.AddParagraph();

// Add text to the new paragraph
newParagraph.AppendText("Add Date Field with setting French Culture: ");

// Append a date field to the new paragraph and set its format
Field field2 = newParagraph.AppendField("\"\\@\"dd MMMM yyyy", FieldType.FieldDate);
field2.CharacterFormat.LocaleIdASCII = 1036;

// Enable automatic update of fields in the document
document.IsUpdateFields = true;
```

---

# Spire.Doc C# Field Locale
## Set locale for document field
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Add a paragraph to the section
Paragraph par = section.AddParagraph();

// Append a date field to the paragraph
Field field = par.AppendField("DocDate", FieldType.FieldDate);

// Set the locale ID to Russian (1049) for the first character range in the field
(field.OwnerParagraph.ChildObjects[0] as TextRange).CharacterFormat.LocaleIdASCII = 1049;

// Set the field text to "2019-10-10"
field.FieldText = "2019-10-10";

// Enable automatic update of fields in the document
document.IsUpdateFields = true;
```

---

# spire.doc csharp field update
## update fields in word document
```csharp
// Load the document from a file
Document document = new Document(@"..\..\..\..\..\..\Data\IfFieldSample.docx");

// Setting the culture source when updating fields
document.FieldOptions.CultureSource = Spire.Doc.Layout.Fields.FieldCultureSource.CurrentThread;

// Enable automatic update of fields in the document
document.IsUpdateFields = true;
```

---

# spire.doc csharp table of contents
## change table of contents style in word document
```csharp
// Create a custom Table of Contents (TOC) style
ParagraphStyle tocStyle = Style.CreateBuiltinStyle(BuiltinStyle.Toc1, doc) as ParagraphStyle;
tocStyle.CharacterFormat.FontName = "Aleo";
tocStyle.CharacterFormat.FontSize = 15f;
tocStyle.CharacterFormat.TextColor = Color.CadetBlue;
doc.Styles.Add(tocStyle);

// Iterate through all sections in the document
foreach (Section section in doc.Sections)
{
    // Iterate through all child objects in the body of each section
    foreach (DocumentObject obj in section.Body.ChildObjects)
    {
        // Check if the object is a StructureDocumentTag (e.g., TOC field)
        if (obj is StructureDocumentTag)
        {
            StructureDocumentTag tag = obj as StructureDocumentTag;
            
            // Iterate through all child objects within the StructureDocumentTag
            foreach (DocumentObject cObj in tag.ChildObjects)
            {
                // Check if the child object is a paragraph
                if (cObj is Paragraph)
                {
                    Paragraph para = cObj as Paragraph;
                    
                    // Check if the paragraph has the style name "TOC1"
                    if (para.StyleName == "TOC1")
                    {
                        // Apply the custom TOC style to the paragraph
                        para.ApplyStyle(tocStyle.Name);
                    }
                }
            }
        }
    }
}
```

---

# Spire.Doc C# TOC Tab Style
## Change Table of Contents tab style in Word document
```csharp
// Iterate through all sections in the document
foreach (Section section in doc.Sections)
{
    // Iterate through all child objects in the body of each section
    foreach (DocumentObject obj in section.Body.ChildObjects)
    {
        // Check if the object is a StructureDocumentTag (e.g., TOC field)
        if (obj is StructureDocumentTag)
        {
            StructureDocumentTag tag = obj as StructureDocumentTag;
            
            // Iterate through all child objects within the StructureDocumentTag
            foreach (DocumentObject cObj in tag.ChildObjects)
            {
                // Check if the child object is a paragraph
                if (cObj is Paragraph)
                {
                    Paragraph para = cObj as Paragraph;
                    
                    // Check if the paragraph has the style name "TOC2"
                    if (para.StyleName == "TOC2")
                    {
                        // Adjust the position and tab leader of each tab in the paragraph's format
                        foreach (Tab tab in para.Format.Tabs)
                        {
                            tab.Position = tab.Position + 20;
                            tab.TabLeader = TabLeader.NoLeader;
                        }
                    }
                }
            }
        }
    }
}
```

---

# Spire.Doc CSharp Table of Contents
## Create a table of contents in a Word document with heading styles
```csharp
// Create a new document
Document doc = new Document();

// Add a section to the document
Section section = doc.AddSection();

// Add a paragraph to the section and append a table of contents (TOC)
Paragraph para = section.AddParagraph();
para.AppendTOC(1, 3);

// Add a paragraph to the section 
Paragraph par = section.AddParagraph();
TextRange tr = par.AppendText("Flowers");
tr.CharacterFormat.FontSize = 30;
par.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

// Add a paragraph to the section 
Paragraph para1 = section.AddParagraph();
para1.AppendText("Ornithogalum");

// Apply the "Heading1" style 
para1.ApplyStyle(BuiltinStyle.Heading1);

// Add a paragraph to the section 
Paragraph para2 = section.AddParagraph();
para2.AppendText("Rosa");

// Apply the "Heading2" style 
para2.ApplyStyle(BuiltinStyle.Heading2);

// Add a paragraph to the section 
Paragraph para3 = section.AddParagraph();
para3.AppendText("Hyacinth");

// Apply the "Heading3" style 
para3.ApplyStyle(BuiltinStyle.Heading3);

// Update the table of contents
doc.UpdateTableOfContents();
```

---

# spire.doc csharp table of contents
## customize table of contents in word document
```csharp
// Create a new document
Document doc = new Document();

// Add a section to the document
Section section = doc.AddSection();

// Create a table of contents and add it to a paragraph in the section
TableOfContent toc = new TableOfContent(doc, "{\\o \"1-3\" \\n 1-1}");
Paragraph para = section.AddParagraph();
para.Items.Add(toc);
para.AppendFieldMark(FieldMarkType.FieldSeparator);
para.AppendText("TOC");
para.AppendFieldMark(FieldMarkType.FieldEnd);
doc.TOC = toc;

// Update the table of contents
doc.UpdateTableOfContents();
```

---

# spire.doc csharp table of content
## remove table of content from word document
```csharp
// Access the body of the first section in the document
Body body = document.Sections[0].Body;

// Define a regular expression pattern to match the style names
Regex regex = new Regex("TOC\\w+");

// Iterate over the paragraphs in the body
for (int i = 0; i < body.Paragraphs.Count; i++)
{
    // Check if the style name matches the regular expression pattern
    if (regex.IsMatch(body.Paragraphs[i].StyleName))
    {
        // Remove the paragraph if it matches the pattern
        body.Paragraphs.RemoveAt(i);
        
        // Decrement the counter to avoid skipping the next paragraph
        i--;
    }
}
```

---

# spire.doc csharp textbox
## delete table from textbox in word document
```csharp
// Create a new Document object
Document doc = new Document();

// Load a Word document from the specified input file
doc.LoadFromFile(input);

// Access the first text box in the document
Spire.Doc.Fields.TextBox textbox = doc.TextBoxes[0];

// Remove the table inside the text box
textbox.Body.Tables.RemoveAt(0);
```

---

# spire.doc csharp textbox
## extract text from textboxes in word document
```csharp
// Check if the document contains any text boxes
if (document.TextBoxes.Count > 0)
{
    // Iterate through the sections in the document
    foreach (Section section in document.Sections)
    {
        // Iterate through the paragraphs in each section
        foreach (Paragraph p in section.Paragraphs)
        {
            // Iterate through the child objects of each paragraph
            foreach (DocumentObject obj in p.ChildObjects)
            {
                // Check if the child object is a text box
                if (obj.DocumentObjectType == DocumentObjectType.TextBox)
                {
                    // Cast the child object to a TextBox
                    Spire.Doc.Fields.TextBox textbox = obj as Spire.Doc.Fields.TextBox;

                    // Iterate through the child objects of the text box
                    foreach (DocumentObject objt in textbox.ChildObjects)
                    {
                        // Check if the child object is a paragraph
                        if (objt.DocumentObjectType == DocumentObjectType.Paragraph)
                        {
                            // Get the text of the paragraph
                            string text = (objt as Paragraph).Text;
                        }

                        // Check if the child object is a table
                        if (objt.DocumentObjectType == DocumentObjectType.Table)
                        {
                            // Cast the child object to a Table
                            Table table = objt as Table;

                            // Extract text from the table
                            ExtractTextFromTables(table);
                        }
                    }
                }
            }
        }
    }
}

// Define a method to extract text from tables
static void ExtractTextFromTables(Table table)
{
    // Iterate through the rows of the table
    for (int i = 0; i < table.Rows.Count; i++)
    {
        // Get the current row
        TableRow row = table.Rows[i];
        
        // Iterate through the cells of the row
        for (int j = 0; j < row.Cells.Count; j++)
        {
            // Get the current cell
            TableCell cell = row.Cells[j];

            // Iterate through the paragraphs in the cell
            foreach (Paragraph paragraph in cell.Paragraphs)
            {
                // Get the text of the paragraph
                string text = paragraph.Text;
            }
        }
    }
}
```

---

# spire.doc csharp textbox
## insert image into textbox
```csharp
// Create a new Document object
Document doc = new Document();

// Add a section to the document
Section section = doc.AddSection();

// Add a paragraph to the section
Paragraph paragraph = section.AddParagraph();

// Append a text box to the paragraph with specified dimensions
Spire.Doc.Fields.TextBox tb = paragraph.AppendTextBox(220, 220);

// Set the horizontal and vertical positioning of the text box
tb.Format.HorizontalOrigin = HorizontalOrigin.Page;
tb.Format.HorizontalPosition = 50;
tb.Format.VerticalOrigin = VerticalOrigin.Page;
tb.Format.VerticalPosition = 50;

// Set the background fill effect of the text box to a picture
tb.Format.FillEfects.Type = BackgroundType.Picture;

// Set the picture for the background fill effect from a file
tb.Format.FillEfects.Picture = Image.FromFile("Spire.Doc.png");
```

---

# spire.doc csharp textbox table
## insert table into textbox in word document
```csharp
// Add a section to the document
Section section = doc.AddSection();

// Add a paragraph to the section
Paragraph paragraph = section.AddParagraph();

// Append a text box to the paragraph with specified dimensions
Spire.Doc.Fields.TextBox textbox = paragraph.AppendTextBox(300, 100);

// Set the horizontal and vertical positioning of the text box
textbox.Format.HorizontalOrigin = HorizontalOrigin.Page;
textbox.Format.HorizontalPosition = 140;
textbox.Format.VerticalOrigin = VerticalOrigin.Page;
textbox.Format.VerticalPosition = 50;

// Add a paragraph to the text box
Paragraph textboxParagraph = textbox.Body.AddParagraph();

// Append text to the paragraph in the text box
TextRange textboxRange = textboxParagraph.AppendText("Table 1");
textboxRange.CharacterFormat.FontName = "Arial";

// Add a table to the body of the text box
Table table = textbox.Body.AddTable(true);

// Reset the number of rows and columns in the table
table.ResetCells(4, 4);

// Define the data for the table
string[,] data = new string[,]
{
    {"Name","Age","Gender","ID" },
    {"John","28","Male","0023" },
    {"Steve","30","Male","0024" },
    {"Lucy","26","female","0025" }
};

// Populate the table with data
for (int i = 0; i < 4; i++)
{
    for (int j = 0; j < 4; j++)
    {
        TextRange tableRange = table[i, j].AddParagraph().AppendText(data[i, j]);
        tableRange.CharacterFormat.FontName = "Arial";
    }
}

// Apply a predefined table style to the table
table.ApplyStyle(DefaultTableStyle.TableColorful2);
```

---

# spire.doc csharp textbox
## lock textbox aspect ratio
```csharp
// Create a document, section, and paragraph
Document document = new Document();
Section section = document.AddSection();
Paragraph paragraph = section.AddParagraph();

// Create a textbox
Spire.Doc.Fields.TextBox textBox1 = paragraph.AppendTextBox(240, 35);

// Configure textbox format
textBox1.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left;
textBox1.Format.LineColor = System.Drawing.Color.Black;
textBox1.Format.LineStyle = TextBoxLineStyle.Simple;

// Lock the aspect ratio of the textbox
textBox1.AspectRatioLocked = true;
```

---

# Spire.Doc C# TextBox Table Extraction
## Extract table data from a textbox in a Word document
```csharp
// Get the first textbox in the document
Spire.Doc.Fields.TextBox textbox = doc.TextBoxes[0];

// Get the first table from the textbox
Table table = textbox.Body.Tables[0] as Table;

// Initialize an empty string to store the table data
string str = null;

// Iterate through each row in the table
foreach (TableRow row in table.Rows)
{
    // Iterate through each cell in the row
    foreach (TableCell cell in row.Cells)
    {
        // Iterate through each paragraph in the cell
        foreach (Paragraph paragraph in cell.Paragraphs)
        {
            // Append the text of each paragraph to the string, separated by a tab
            str += paragraph.Text + "\t";
        }
    }
    // Add a new line after processing each row
    str += "\r\n";
}
```

---

# spire.doc csharp textbox
## remove textbox from word document
```csharp
// Create a new instance of Document
Document doc = new Document();

// Remove the first text box in the document
doc.TextBoxes.RemoveAt(0);

// Clear all the text boxes in the document
//doc.TextBoxes.Clear();
```

---

# spire.doc csharp textbox
## insert and customize textboxes in Word document
```csharp
private void InsertTextbox(Section section)
{
    // Create a paragraph in the specified section
    Paragraph paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();

    // Add three paragraphs to create space between textboxes
    paragraph = section.AddParagraph();
    paragraph = section.AddParagraph();
    paragraph = section.AddParagraph();

    // Create and customize textbox 1
    Spire.Doc.Fields.TextBox textBox1 = paragraph.AppendTextBox(240, 35);
    textBox1.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left;
    textBox1.Format.LineColor = System.Drawing.Color.Gray;
    textBox1.Format.LineStyle = TextBoxLineStyle.Simple;
    textBox1.Format.FillColor = System.Drawing.Color.DarkSeaGreen;
    Paragraph para = textBox1.Body.AddParagraph();
    TextRange txtrg = para.AppendText("Textbox 1 in the document");
    txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
    txtrg.CharacterFormat.FontSize = 14;
    txtrg.CharacterFormat.TextColor = System.Drawing.Color.White;
    para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

    // Add four paragraphs to create space between textboxes
    paragraph = section.AddParagraph();
    paragraph = section.AddParagraph();
    paragraph = section.AddParagraph();
    paragraph = section.AddParagraph();

    // Create and customize textbox 2
    Spire.Doc.Fields.TextBox textBox2 = paragraph.AppendTextBox(240, 35);
    textBox2.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left;
    textBox2.Format.LineColor = System.Drawing.Color.Tomato;
    textBox2.Format.LineStyle = TextBoxLineStyle.ThinThick;
    textBox2.Format.FillColor = System.Drawing.Color.Blue;
    textBox2.Format.LineDashing = LineDashing.Dot;
    para = textBox2.Body.AddParagraph();
    txtrg = para.AppendText("Textbox 2 in the document");
    txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
    txtrg.CharacterFormat.FontSize = 14;
    txtrg.CharacterFormat.TextColor = System.Drawing.Color.Pink;
    para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

    // Add four paragraphs to create space between textboxes
    paragraph = section.AddParagraph();
    paragraph = section.AddParagraph();
    paragraph = section.AddParagraph();
    paragraph = section.AddParagraph();

    // Create and customize textbox 3
    Spire.Doc.Fields.TextBox textBox3 = paragraph.AppendTextBox(240, 35);
    textBox3.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left;
    textBox3.Format.LineColor = System.Drawing.Color.Violet;
    textBox3.Format.LineStyle = TextBoxLineStyle.Triple;
    textBox3.Format.FillColor = System.Drawing.Color.Pink;
    textBox3.Format.LineDashing = LineDashing.DashDotDot;
    para = textBox3.Body.AddParagraph();
    txtrg = para.AppendText("Textbox 3 in the document");
    txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
    txtrg.CharacterFormat.FontSize = 14;
    txtrg.CharacterFormat.TextColor = System.Drawing.Color.Tomato;
    para.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
}
```

---

# spire.doc textbox formatting
## create and format a textbox in a word document
```csharp
// Create a new instance of Document
Document doc = new Document();

// Add a new section to the document
Section sec = doc.AddSection();

// Add a textbox to the first paragraph in the section and get a reference to it
Spire.Doc.Fields.TextBox TB = doc.Sections[0].AddParagraph().AppendTextBox(310, 90);

// Add a paragraph to the body of the textbox and get a reference to it
Paragraph para = TB.Body.AddParagraph();

// Add text to the paragraph
TextRange TR = para.AppendText("Using Spire.Doc, developers will find " +
    "a simple and effective method to endow their applications with rich MS Word features. ");

// Set the font properties for the text
TR.CharacterFormat.FontName = "Cambria";
TR.CharacterFormat.FontSize = 13;

// Configure the position of the textbox
TB.Format.HorizontalOrigin = HorizontalOrigin.Page;
TB.Format.HorizontalPosition = 120;
TB.Format.VerticalOrigin = VerticalOrigin.Page;
TB.Format.VerticalPosition = 100;

// Configure the line style and color of the textbox
TB.Format.LineStyle = TextBoxLineStyle.Double;
TB.Format.LineColor = Color.CornflowerBlue;
TB.Format.LineDashing = LineDashing.Solid;
TB.Format.LineWidth = 5;

// Configure the internal margins of the textbox
TB.Format.InternalMargin.Top = 15;
TB.Format.InternalMargin.Bottom = 10;
TB.Format.InternalMargin.Left = 12;
TB.Format.InternalMargin.Right = 10;
```

---

# spire.doc csharp watermark
## add image watermark to word document
```csharp
private void InsertImageWatermark(Document document) {
    // Create a PictureWatermark object
    PictureWatermark picture = new PictureWatermark();
    // Load the image for the watermark
    picture.Picture = System.Drawing.Image.FromFile(@"..\..\..\..\..\..\Data\ImageWatermark.png");
    // Set the scaling of the watermark
    picture.Scaling = 250;
    // Specify whether the watermark should be washed out
    picture.IsWashout = false;
    // Set the watermark for the document
    document.Watermark = picture;
}
```

---

# Spire.Doc C# Watermark Removal
## Remove image watermark from Word document
```csharp
// Remove the watermark from the document
document.Watermark = null;
```

---

# spire.doc csharp watermark
## remove text watermark from document
```csharp
// Remove the watermark from the document
document.Watermark = null;
```

---

# Spire.Doc C# Text Watermark
## Add text watermark to Word document
```csharp
private void InsertTextWatermark(Section section) {
    // Create a TextWatermark object
    TextWatermark txtWatermark = new TextWatermark();
    // Set the text for the watermark
    txtWatermark.Text = "E-iceblue";
    // Set the font size of the watermark
    txtWatermark.FontSize = 95;
    // Set the color of the watermark
    txtWatermark.Color = Color.Blue;
    // Set the layout of the watermark
    txtWatermark.Layout = WatermarkLayout.Diagonal;
    // Set the watermark for the document section
    section.Document.Watermark = txtWatermark;
}
```

---

# Spire.Doc OLE Extraction
## Extract OLE objects from Word documents and save them as files
```csharp
// Load the document
Document doc = new Document();
doc.LoadFromFile(@"..\..\..\..\..\..\Data\OLEs.docx");

// Iterate through sections and paragraphs to find OLE objects
foreach (Section sec in doc.Sections)
{
    foreach (DocumentObject obj in sec.Body.ChildObjects)
    {
        if (obj is Paragraph)
        {
            Paragraph par = obj as Paragraph;
            foreach (DocumentObject o in par.ChildObjects)
            {
                // Check if the object is an OLE object
                if (o.DocumentObjectType == DocumentObjectType.OleObject)
                {
                    DocOleObject Ole = o as DocOleObject;
                    string oleType = Ole.ObjectType;

                    // Save OLE object based on its type
                    if (oleType == "AcroExch.Document.DC")
                    {
                        File.WriteAllBytes("Result.pdf", Ole.NativeData);
                        FileViewer("Result.pdf");
                    }
                    else if (oleType == "Excel.Sheet.8")
                    {
                        File.WriteAllBytes("ExcelResult.xls", Ole.NativeData);
                        FileViewer("ExcelResult.xls");
                    }
                    else if (oleType == "PowerPoint.Show.12")
                    {
                        File.WriteAllBytes("PPTResult.pptx", Ole.NativeData);
                        FileViewer("PPTResult.pptx");
                    }
                }
            }
        }
    }
}

doc.Dispose();
```

```csharp
private void FileViewer(string fileName)
{
    try
    {
        System.Diagnostics.Process.Start(fileName);
    }
    catch { }
}
```

---

# spire.doc csharp ole
## insert OLE object into Word document
```csharp
// Create a new document object
Document doc = new Document();

// Add a section to the document
Section sec = doc.AddSection();

// Add a paragraph to the section
Paragraph par = sec.AddParagraph();

// Create a DocPicture object and load an image from file
DocPicture picture = new DocPicture(doc);
Image image = Image.FromFile("excel.png");
picture.LoadImage(image);

// Append an OLE object to the paragraph with the specified file, picture, and object type (Excel worksheet)
DocOleObject obj = par.AppendOleObject("example.xlsx", picture, OleObjectType.ExcelWorksheet);
```

---

# spire.doc csharp ole object
## insert OLE object as icon via stream in Word document
```csharp
// Create a new document object
Document doc = new Document();

// Add a section to the document
Section sec = doc.AddSection();

// Add a paragraph to the section
Paragraph par = sec.AddParagraph();

// Open a stream for the OLE object data
Stream stream = File.OpenRead(@"..\..\..\..\..\..\Data\example.zip");

// Create a DocPicture object and load an image
DocPicture picture = new DocPicture(doc);
Image image = Image.FromFile(@"..\..\..\..\..\..\Data\example.png");
picture.LoadImage(image);

// Append an OLE object to the paragraph using the stream, picture, and object type
DocOleObject obj = par.AppendOleObject(stream, picture, "zip");

// Set the OLE object to be displayed as an icon
obj.DisplayAsIcon = true;
```

---

# Spire.Doc C# Checkbox Content Control
## Create and configure checkbox content control in Word document
```csharp
// Create an inline structure document tag (SDT)
StructureDocumentTagInline sdt = new StructureDocumentTagInline(document);

// Set the SDT type to CheckBox
sdt.SDTProperties.SDTType = SdtType.CheckBox;

// Create and configure the checkbox control
SdtCheckBox scb = new SdtCheckBox();
sdt.SDTProperties.ControlProperties = scb;

// Create a TextRange with specific formatting for the checkbox
TextRange tr = new TextRange(document);
tr.CharacterFormat.FontName = "MS Gothic";
tr.CharacterFormat.FontSize = 12;
sdt.ChildObjects.Add(tr);

// Set the CheckBox state
scb.Checked = true;

// Add the SDT to a paragraph
paragraph.ChildObjects.Add(sdt);
```

---

# spire.doc csharp content controls
## add various content controls to a Word document
```csharp
// Add Combo Box Content Control
Paragraph paragraph = section.AddParagraph();
TextRange txtRange = paragraph.AppendText("Add Combo Box Content Control:  ");
txtRange.CharacterFormat.Italic = true;
StructureDocumentTagInline sd = new StructureDocumentTagInline(document);
paragraph.ChildObjects.Add(sd);
sd.SDTProperties.SDTType = SdtType.ComboBox;
SdtComboBox cb = new SdtComboBox();
cb.ListItems.Add(new SdtListItem("Spire.Doc"));
cb.ListItems.Add(new SdtListItem("Spire.XLS"));
cb.ListItems.Add(new SdtListItem("Spire.PDF"));
sd.SDTProperties.ControlProperties = cb;
TextRange rt = new TextRange(document);
rt.Text = cb.ListItems[0].DisplayText;
sd.SDTContent.ChildObjects.Add(rt);

// Add Text Content Control
paragraph = section.AddParagraph();
txtRange = paragraph.AppendText("Add Text Content Control:  ");
txtRange.CharacterFormat.Italic = true;
sd = new StructureDocumentTagInline(document);
paragraph.ChildObjects.Add(sd);
sd.SDTProperties.SDTType = SdtType.Text;
SdtText text = new SdtText(true);
text.IsMultiline = true;
sd.SDTProperties.ControlProperties = text;
rt = new TextRange(document);
rt.Text = "Text";
sd.SDTContent.ChildObjects.Add(rt);

// Add Picture Content Control
paragraph = section.AddParagraph();
txtRange = paragraph.AppendText("Add Picture Content Control:  ");
txtRange.CharacterFormat.Italic = true;
sd = new StructureDocumentTagInline(document);
paragraph.ChildObjects.Add(sd);
sd.SDTProperties.SDTType = SdtType.Picture;
DocPicture pic = new DocPicture(document);
pic.Width = 10;
pic.Height = 10;
pic.LoadImage(Image.FromFile(@"..\..\..\..\..\..\Data\logo.png"));
sd.SDTContent.ChildObjects.Add(pic);

// Add Date Picker Content Control
paragraph = section.AddParagraph();
txtRange = paragraph.AppendText("Add Date Picker Content Control:  ");
txtRange.CharacterFormat.Italic = true;
sd = new StructureDocumentTagInline(document);
paragraph.ChildObjects.Add(sd);
sd.SDTProperties.SDTType = SdtType.DatePicker;
SdtDate date = new SdtDate();
date.CalendarType = CalendarType.Default;
date.DateFormat = "yyyy.MM.dd";
date.FullDate = DateTime.Now;
sd.SDTProperties.ControlProperties = date;
rt = new TextRange(document);
rt.Text = "1990.02.08";
sd.SDTContent.ChildObjects.Add(rt);

// Add Drop-Down List Content Control
paragraph = section.AddParagraph();
txtRange = paragraph.AppendText("Add Drop-Down List Content Control:  ");
txtRange.CharacterFormat.Italic = true;
sd = new StructureDocumentTagInline(document);
paragraph.ChildObjects.Add(sd);
sd.SDTProperties.SDTType = SdtType.DropDownList;
SdtDropDownList sddl = new SdtDropDownList();
sddl.ListItems.Add(new SdtListItem("Harry"));
sddl.ListItems.Add(new SdtListItem("Jerry"));
sd.SDTProperties.ControlProperties = sddl;
rt = new TextRange(document);
rt.Text = sddl.ListItems[0].DisplayText;
sd.SDTContent.ChildObjects.Add(rt);
```

---

# spire.doc c# richtext content control
## Add a Rich Text Content Control to a Word document with formatted text
```csharp
// Create an inline structure document tag (SDT) and add it to the paragraph's child objects
StructureDocumentTagInline sdt = new StructureDocumentTagInline(document);
paragraph.ChildObjects.Add(sdt);

// Set the SDT type to RichText
sdt.SDTProperties.SDTType = SdtType.RichText;

// Create an instance of SdtText, set its multiline property, and assign it as the control properties for the SDT
SdtText text = new SdtText(true);
text.IsMultiline = true;
sdt.SDTProperties.ControlProperties = text;

// Create a TextRange object and set its text and text color, then add it to the SDT's content
TextRange rt = new TextRange(document);
rt.Text = "Welcome to use ";
rt.CharacterFormat.TextColor = Color.Green;
sdt.SDTContent.ChildObjects.Add(rt);

// Create another TextRange object and set its text and text color, then add it to the SDT's content
rt = new TextRange(document);
rt.Text = "Spire.Doc";
rt.CharacterFormat.TextColor = Color.OrangeRed;
sdt.SDTContent.ChildObjects.Add(rt);
```

---

# spire.doc csharp combobox
## modify ComboBox items in Word document
```csharp
// Iterate through each section in the document
foreach (Section section in doc.Sections)
{
    // Iterate through each document object in the section's body
    foreach (DocumentObject bodyObj in section.Body.ChildObjects)
    {
        // Check if the document object is a StructureDocumentTag
        if (bodyObj.DocumentObjectType == DocumentObjectType.StructureDocumentTag)
        {
            // Check if the StructureDocumentTag is of type ComboBox
            if ((bodyObj as StructureDocumentTag).SDTProperties.SDTType == SdtType.ComboBox)
            {
                // Access the ComboBox control properties
                SdtComboBox combo = (bodyObj as StructureDocumentTag).SDTProperties.ControlProperties as SdtComboBox;

                // Remove an item from the ComboBox
                combo.ListItems.RemoveAt(1);

                // Create a new SdtListItem and add it to the ComboBox
                SdtListItem item = new SdtListItem("D", "D");
                combo.ListItems.Add(item);

                // Set the selected value of the ComboBox based on the item value "D"
                foreach (SdtListItem sdtItem in combo.ListItems)
                {
                    if (string.CompareOrdinal(sdtItem.Value, "D") == 0)
                    {
                        combo.ListItems.SelectedValue = sdtItem;
                    }
                }
            }
        }
    }
}
```

---

# spire.doc csharp content control properties
## Extract properties from structured document tags in a Word document
```csharp
// Get all the structure tags in the document
StructureTags structureTags = GetAllTags(doc);

// Initialize variables for storing tag properties
string alias = null;
decimal id = 0;
string tag = null;
string property = "Alias of contentControl" + "\t" + "ID          " + "\t" + "Tag             " + "\t" + "STDType        " + "\r" + "Content        " + "\r\n";
string sdtType = null;
Paragraph paragraph = null;
SdtType sdt = SdtType.RichText;
string content = "";
TextRange textRange = null;

// Retrieve structure document tags and process their properties and content
List<StructureDocumentTag> tags = structureTags.tags;
for (int i = 0; i < tags.Count; i++)
{
    alias = tags[i].SDTProperties.Alias;
    id = tags[i].SDTProperties.Id;
    tag = tags[i].SDTProperties.Tag;
    sdt = tags[i].SDTProperties.SDTType;
    sdtType = sdt.ToString();
    if (sdt == SdtType.RichText || sdt == SdtType.Text)
    {
        if (tags[i].ChildObjects.Count > 0)
        {
            foreach (DocumentObject obj in tags[i].ChildObjects)
            {
                if (obj is Paragraph)
                {
                    paragraph = obj as Paragraph;
                    content += paragraph.Text;
                }
            }
        }
    }
    property += alias + ",\t" + id + ",\t" + tag + ",\t" + sdtType + ",\t" + content + "\r\n";
    content = "";
}

// Retrieve structure document tag inlines and process their properties and content
List<StructureDocumentTagInline> tagInlines = structureTags.tagInlines;
for (int i = 0; i < tagInlines.Count; i++)
{
    alias = tagInlines[i].SDTProperties.Alias;
    id = tagInlines[i].SDTProperties.Id;
    tag = tagInlines[i].SDTProperties.Tag;
    sdt = tagInlines[i].SDTProperties.SDTType;
    sdtType = sdt.ToString();
    if (sdt == SdtType.RichText || sdt == SdtType.Text)
    {
        if (tagInlines[i].ChildObjects.Count > 0)
        {
            foreach (DocumentObject obj in tagInlines[i].ChildObjects)
            {
                if (obj is TextRange)
                {
                    textRange = obj as TextRange;
                    content += textRange.Text;
                }
            }
        }
    }
    property += alias + ",\t" + id + ",\t" + tag + ",\t" + sdtType + ",\t" + content + "\r\n";
    content = "";
}

// Retrieve structure document tag rows and process their properties and content
List<StructureDocumentTagRow> rowTags = structureTags.rowTags;
for (int i = 0; i < rowTags.Count; i++)
{
    alias = rowTags[i].SDTProperties.Alias;
    id = rowTags[i].SDTProperties.Id;
    tag = rowTags[i].SDTProperties.Tag;
    sdt = rowTags[i].SDTProperties.SDTType;
    sdtType = sdt.ToString();
    if (sdt == SdtType.RichText || sdt == SdtType.Text)
    {
        if (rowTags[i].ChildObjects.Count > 0)
        {
            foreach (DocumentObject obj in rowTags[i].ChildObjects)
            {
                if (obj is Paragraph)
                {
                    paragraph = obj as Paragraph;
                    content += paragraph.Text;
                }
            }
        }
    }
    property += alias + ",\t" + id + ",\t" + tag + ",\t" + sdtType + ",\t" + content + "\r\n";
    content = "";
}

// Retrieve structure document tag cells and process their properties and content
List<StructureDocumentTagCell> cellTags = structureTags.cellTags;
for (int i = 0; i < cellTags.Count; i++)
{
    alias = cellTags[i].SDTProperties.Alias;
    id = cellTags[i].SDTProperties.Id;
    tag = cellTags[i].SDTProperties.Tag;
    sdt = cellTags[i].SDTProperties.SDTType;
    sdtType = sdt.ToString();
    if (sdt == SdtType.RichText || sdt == SdtType.Text)
    {
        if (cellTags[i].ChildObjects.Count > 0)
        {
            foreach (DocumentObject obj in cellTags[i].ChildObjects)
            {
                if (obj is Paragraph)
                {
                    paragraph = obj as Paragraph;
                    content += paragraph.Text;
                }
            }
        }
    }
    property += alias + ",\t" + id + ",\t" + tag + ",\t" + sdtType + ",\t" + content + "\r\n";
    content = "";
}

//Get all StructureTags of the Word document
private static StructureTags GetAllTags(Document document)
{
    StructureTags structureTags = new StructureTags();
    foreach (Section section in document.Sections)
    {
        foreach (DocumentObject obj in section.Body.ChildObjects)
        {
            if (obj.DocumentObjectType == DocumentObjectType.StructureDocumentTag)
            {
                structureTags.tags.Add(obj as StructureDocumentTag);

            }

            else if (obj.DocumentObjectType == DocumentObjectType.Paragraph)
            {
                foreach (DocumentObject pobj in (obj as Paragraph).ChildObjects)
                {
                    if (pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline)
                    {
                        structureTags.tagInlines.Add(pobj as StructureDocumentTagInline);
                    }
                }
            }
            else if (obj.DocumentObjectType == DocumentObjectType.Table)
            {
                foreach (TableRow row in (obj as Table).Rows)
                {
                    if (row is StructureDocumentTagRow)
                    {
                        structureTags.rowTags.Add(row as StructureDocumentTagRow);
                    }
                    foreach (TableCell cell in row.Cells)
                    {
                        if (cell is StructureDocumentTagCell)
                        {
                            structureTags.cellTags.Add(cell as StructureDocumentTagCell);
                        }
                        foreach (DocumentObject cellChild in cell.ChildObjects)
                        {
                            if (cellChild.DocumentObjectType == DocumentObjectType.StructureDocumentTag)
                            {
                                structureTags.tags.Add(cellChild as StructureDocumentTag);
                            }
                            else if (cellChild.DocumentObjectType == DocumentObjectType.Paragraph)
                            {
                                foreach (DocumentObject pobj in (cellChild as Paragraph).ChildObjects)
                                {
                                    if (pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline)
                                    {
                                        structureTags.tagInlines.Add(pobj as StructureDocumentTagInline);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    return structureTags;
}

public class StructureTags
{
    List<StructureDocumentTagInline> m_tagInlines;
    public List<StructureDocumentTagInline> tagInlines
    {
        get
        {
            if (m_tagInlines == null)
                m_tagInlines = new List<StructureDocumentTagInline>();
            return m_tagInlines;
        }
        set
        {
            m_tagInlines = value;
        }
    }
    List<StructureDocumentTag> m_tags;
    public List<StructureDocumentTag> tags
    {
        get
        {
            if (m_tags == null)
                m_tags = new List<StructureDocumentTag>();
            return m_tags;
        }
        set
        {
            m_tags = value;
        }
    }
    List<StructureDocumentTagCell> m_celltags;
    public List<StructureDocumentTagCell> cellTags
    {
        get
        {
            if (m_celltags == null)
                m_celltags = new List<StructureDocumentTagCell>();
            return m_celltags;
        }
        set
        {
            m_celltags = value;
        }
    }
    List<StructureDocumentTagRow> m_rowTags;
    public List<StructureDocumentTagRow> rowTags
    {
        get
        {
            if (m_rowTags == null)
                m_rowTags = new List<StructureDocumentTagRow>();
            return m_rowTags;
        }
        set
        {
            m_rowTags = value;
        }
    }
}
```

---

# spire.doc csharp structured document tag
## lock content control content in a document
```csharp
// Create a new document
Document doc = new Document();

// Add a section to the document
Section section = doc.AddSection();

// Add a paragraph to the section
Paragraph paragraph = section.AddParagraph();

// Create a StructureDocumentTag
StructureDocumentTag sdt = new StructureDocumentTag(doc);

// Add a new section to the document
Section section2 = doc.AddSection();

// Add the StructureDocumentTag to the section's body
section2.Body.ChildObjects.Add(sdt);

// Set the type of the StructureDocumentTag to RichText
sdt.SDTProperties.SDTType = SdtType.RichText;

// Iterate through the child objects in the first section's body
foreach (DocumentObject obj in section.Body.ChildObjects)
{
    // Check if the object is a table
    if (obj.DocumentObjectType == DocumentObjectType.Table)
    {
        // Clone and add the table to the StructureDocumentTag's content
        sdt.SDTContent.ChildObjects.Add(obj.Clone());
    }
}

// Lock the content editing settings of the StructureDocumentTag
sdt.SDTProperties.LockSettings = LockSettingsType.ContentLocked;

// Remove the first section from the document
doc.Sections.Remove(section);
```

---

# spire.doc structured document tag color modification
## modify the color of structured document tags in a word document
```csharp
// Iterate through the sections in the document
for (int s = 0; s < doc.Sections.Count; s++)
{
    // Get the current section
    Section section = doc.Sections[s];

    // Iterate through the child objects in the section's body
    for (int i = 0; i < section.Body.ChildObjects.Count; i++)
    {
        // Check if the child object is a Paragraph
        if (section.Body.ChildObjects[i] is Paragraph)
        {
            // Get the paragraph object
            Paragraph para = section.Body.ChildObjects[i] as Paragraph;
            
            // Iterate through the child objects in the paragraph
            for (int j = 0; j < para.ChildObjects.Count; j++)
            {
                // Check if the child object is a StructureDocumentTagInline
                if (para.ChildObjects[j] is StructureDocumentTagInline)
                {
                    // Get the StructureDocumentTagInline object
                    StructureDocumentTagInline sdt = para.ChildObjects[j] as StructureDocumentTagInline;
                    
                    // Get the SDTProperties of the StructureDocumentTagInline
                    SDTProperties sDTProperties = sdt.SDTProperties;

                    // Set the color of the SDTProperties based on the SDTType
                    switch (sDTProperties.SDTType)
                    {
                        case SdtType.RichText:
                            sDTProperties.Color = Color.Orange;
                            break;
                        case SdtType.Text:
                            sDTProperties.Color = Color.Green;
                            break;
                    }
                }
            }
        }

        // Check if the child object is a StructureDocumentTag
        if (section.Body.ChildObjects[i] is StructureDocumentTag)
        {
            // Get the StructureDocumentTag object
            StructureDocumentTag sdt = section.Body.ChildObjects[i] as StructureDocumentTag;
            
            // Get the SDTProperties of the StructureDocumentTag
            SDTProperties sDTProperties = sdt.SDTProperties;

            // Set the color of the SDTProperties based on the SDTType
            switch (sDTProperties.SDTType)
            {
                case SdtType.RichText:
                    sDTProperties.Color = Color.Orange;
                    break;
                case SdtType.Text:
                    sDTProperties.Color = Color.Green;
                    break;
            }
        }
    }
}
```

---

# Spire.Doc C# Remove Content Controls
## Remove structured document tags (content controls) from a Word document
```csharp
// Iterate through the sections in the document
for (int s = 0; s < doc.Sections.Count; s++)
{
    // Get the current section
    Section section = doc.Sections[s];

    // Iterate through the child objects in the section's body
    for (int i = 0; i < section.Body.ChildObjects.Count; i++)
    {
        // Check if the child object is a paragraph
        if (section.Body.ChildObjects[i] is Paragraph)
        {
            // Get the paragraph object
            Paragraph para = section.Body.ChildObjects[i] as Paragraph;
            
            // Iterate through the child objects in the paragraph
            for (int j = 0; j < para.ChildObjects.Count; j++)
            {
                // Check if the child object is a StructureDocumentTagInline
                if (para.ChildObjects[j] is StructureDocumentTagInline)
                {
                    // Get the StructureDocumentTagInline object
                    StructureDocumentTagInline sdt = para.ChildObjects[j] as StructureDocumentTagInline;
                    
                    // Remove the StructureDocumentTagInline from the paragraph
                    para.ChildObjects.Remove(sdt);
                    
                    // Decrement the index to account for the removed object
                    j--;
                }
            }
        }
        
        // Check if the child object is a StructureDocumentTag
        if (section.Body.ChildObjects[i] is StructureDocumentTag)
        {
            // Get the StructureDocumentTag object
            StructureDocumentTag sdt = section.Body.ChildObjects[i] as StructureDocumentTag;
            
            // Remove the StructureDocumentTag from the section's body
            section.Body.ChildObjects.Remove(sdt);
            
            // Decrement the index to account for the removed object
            i--;
        }
    }
}
```

---

# Spire.Doc C# Content Control Appearance
## Set appearance of structured document tags based on their type
```csharp
// Iterate through the sections in the document
foreach (Section section in doc.Sections)
{
    // Iterate through the child objects in the section's body
    foreach (DocumentObject docObj in section.Body.ChildObjects)
    {
        // Check if the current object is a StructureDocumentTag
        if (docObj is StructureDocumentTag)
        {
            // Get the StructureDocumentTag object and its SDTProperties
            StructureDocumentTag stdTag = (StructureDocumentTag)docObj;
            SDTProperties sDTProperties = stdTag.SDTProperties;

            // Set the appearance of the StructureDocumentTag based on its SDTType
            switch (sDTProperties.SDTType)
            {
                case SdtType.Text:
                    sDTProperties.Appearance = SdtAppearance.BoundingBox;
                    break;
                case SdtType.RichText:
                    sDTProperties.Appearance = SdtAppearance.Hidden;
                    break;
                case SdtType.Picture:
                    sDTProperties.Appearance = SdtAppearance.Tags;
                    break;
                case SdtType.CheckBox:
                    sDTProperties.Appearance = SdtAppearance.Default;
                    break;
            }
        }
    }
}
```

---

# spire.doc csharp checkbox
## update checkboxes in structured document tags
```csharp
// Get all the StructureTags from the document
StructureTags structureTags = GetAllTags(document);

// Get the list of StructureDocumentTagInline objects from the StructureTags
List<StructureDocumentTagInline> tagInlines = structureTags.tagInlines;

// Iterate through the list of StructureDocumentTagInline objects
for (int i = 0; i < tagInlines.Count; i++)
{
    // Get the SDTType of the current StructureDocumentTagInline
    string type = tagInlines[i].SDTProperties.SDTType.ToString();

    // Check if the SDTType is "CheckBox"
    if (type == "CheckBox")
    {
        // Get the SdtCheckBox from the ControlProperties of the StructureDocumentTagInline
        SdtCheckBox scb = tagInlines[i].SDTProperties.ControlProperties as SdtCheckBox;

        // Toggle the Checked property of the SdtCheckBox
        if (scb.Checked)
        {
            scb.Checked = false;
        }
        else
        {
            scb.Checked = true;
        }
    }
}

static StructureTags GetAllTags(Document document)
{
    // Create a new StructureTags object to store the StructureDocumentTagInline objects
    StructureTags structureTags = new StructureTags();

    // Iterate through the sections in the document
    foreach (Section section in document.Sections)
    {
        // Iterate through the child objects in the section's body
        foreach (DocumentObject obj in section.Body.ChildObjects)
        {
            // Check if the current object is a Paragraph
            if (obj.DocumentObjectType == DocumentObjectType.Paragraph)
            {
                // Iterate through the child objects in the paragraph
                foreach (DocumentObject pobj in (obj as Paragraph).ChildObjects)
                {
                    // Check if the current object is a StructureDocumentTagInline
                    if (pobj.DocumentObjectType == DocumentObjectType.StructureDocumentTagInline)
                    {
                        // Add the StructureDocumentTagInline to the tagInlines list in the StructureTags object
                        structureTags.tagInlines.Add(pobj as StructureDocumentTagInline);
                    }
                }
            }
        }
    }

    // Return the StructureTags object containing the collected StructureDocumentTagInline objects
    return structureTags;
}

public class StructureTags
{
    List<StructureDocumentTagInline> m_tagInlines;
    public List<StructureDocumentTagInline> tagInlines
    {
        get
        {
            if (m_tagInlines == null)
                m_tagInlines = new List<StructureDocumentTagInline>();
            return m_tagInlines;
        }
        set
        {
            m_tagInlines = value;
        }
    }
}
```

---

# spire.doc csharp math equation
## create math equations from latex and mathml code
```csharp
// Create an OfficeMath object from LaTeX math code
officeMath = new OfficeMath(doc);
officeMath.FromLatexMathCode(latexMathCode[i - 1]);
paragraph.Items.Add(officeMath);

// Convert OfficeMath object to MathML code
paragraph.Text = mathEquations[i - 1].ToMathMLCode();

// Create an OfficeMath object from MathML code
officeMath = new OfficeMath(doc);
officeMath.FromMathMLCode(mathEquations[i - 1].ToMathMLCode());
paragraph.Items.Add(officeMath);
```

---

# Spire.Doc C# Math Equation Extraction
## Extract mathematical equations from Word document and convert to MathML
```csharp
Document doc = new Document();
doc.LoadFromFile("input.docx");
List<OfficeMath> mathEquations = new List<OfficeMath>();
StringBuilder stringBuilder = new StringBuilder();
foreach (Section section in doc.Sections)
{
    foreach (Paragraph paragraph in section.Paragraphs)
    {
        foreach (DocumentObject obj in paragraph.ChildObjects)
        {
            if (obj is OfficeMath)
            {
                stringBuilder.AppendLine((obj as OfficeMath).ToMathMLCode());
                stringBuilder.AppendLine();
                mathEquations.Add(obj as OfficeMath);
            }
        }
    }
}
```

---

# spire.doc csharp math equation conversion
## convert OfficeMath equations to OfficeMathML code
```csharp
// Iterate through sections in the document
foreach (Section section in doc.Sections)
{
    // Iterate through paragraphs in each section
    foreach (Paragraph par in section.Body.Paragraphs)
    {
        // Iterate through child objects in each paragraph
        foreach (DocumentObject obj in par.ChildObjects)
        {
            // Check if the object is an OfficeMath equation
            OfficeMath omath = obj as OfficeMath;
            if (omath == null) continue;
            // Convert OfficeMath equation to MathML code
            string mathml = omath.ToOfficeMathMLCode();
        }
    }
}
```

---

# spire.doc csharp endnote
## insert and format endnote in word document
```csharp
// Get the first section in the document
Section s = doc.Sections[0];

// Get the second paragraph in the section (index 1)
Paragraph p = s.Paragraphs[1];

// Append an endnote to the paragraph
Footnote endnote = p.AppendFootnote(FootnoteType.Endnote);

// Add a paragraph to the endnote's text body and append the reference text
TextRange text = endnote.TextBody.AddParagraph().AppendText("Reference: Wikipedia");

// Set the font name, size, and text color of the reference text
text.CharacterFormat.FontName = "Impact";
text.CharacterFormat.FontSize = 14;
text.CharacterFormat.TextColor = Color.DarkOrange;

// Set the font name, size, and text color of the endnote marker
endnote.MarkerCharacterFormat.FontName = "Calibri";
endnote.MarkerCharacterFormat.FontSize = 25;
endnote.MarkerCharacterFormat.TextColor = Color.DarkBlue;
```

---

# Spire.Doc C# Footnote
## Insert footnote into Word document
```csharp
// Create a new instance of Document
Document document = new Document();

// Find the specified string in the document
TextSelection selection = document.FindString("Spire.Doc", false, true);

// Get the selected text as a single range
TextRange textRange = selection.GetAsOneRange();

// Get the paragraph that contains the selected text
Paragraph paragraph = textRange.OwnerParagraph;

// Get the index of the selected text within the paragraph's child objects
int index = paragraph.ChildObjects.IndexOf(textRange);

// Append a footnote to the paragraph
Footnote footnote = paragraph.AppendFootnote(FootnoteType.Footnote);

// Insert the footnote into the paragraph's child objects at the specified index
paragraph.ChildObjects.Insert(index + 1, footnote);

// Add a paragraph to the footnote's text body and append text to it
textRange = footnote.TextBody.AddParagraph().AppendText("Welcome to evaluate Spire.Doc");

// Set the font name, size, and color for the appended text
textRange.CharacterFormat.FontName = "Arial Black";
textRange.CharacterFormat.FontSize = 10;
textRange.CharacterFormat.TextColor = Color.DarkGray;

// Set the font name, size, style, and color for the footnote marker
footnote.MarkerCharacterFormat.FontName = "Calibri";
footnote.MarkerCharacterFormat.FontSize = 12;
footnote.MarkerCharacterFormat.Bold = true;
footnote.MarkerCharacterFormat.TextColor = Color.DarkGreen;
```

---

# spire.doc csharp footnote
## remove footnotes from document
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Iterate through each paragraph in the section
foreach (Paragraph para in section.Paragraphs)
{
    int index = -1;
    
    // Find the index of the first footnote within the paragraph's child objects
    for (int i = 0, cnt = para.ChildObjects.Count; i < cnt; i++)
    {
        ParagraphBase pBase = para.ChildObjects[i] as ParagraphBase;
        
        if (pBase is Footnote)
        {
            index = i;
            break;
        }
    }

    // If a footnote is found, remove it from the paragraph's child objects
    if (index > -1)
        para.ChildObjects.RemoveAt(index);
}
```

---

# spire.doc csharp footnote
## set footnote position and number format
```csharp
// Get the first section of the document
Section sec = doc.Sections[0];

// Set the footnote options for the section
sec.FootnoteOptions.NumberFormat = FootnoteNumberFormat.UpperCaseLetter;
sec.FootnoteOptions.RestartRule = FootnoteRestartRule.RestartPage;
sec.FootnoteOptions.Position = FootnotePosition.PrintAsEndOfSection;
```

---

# spire.doc print preview
## create and show print preview dialog for word document
```csharp
// Create a new instance of Document
Document doc = new Document();

// Load the Word document from the specified input file
doc.LoadFromFile(input);

// Get the PrintDocument associated with the document
PrintDocument printDoc = doc.PrintDocument;

// Create a new PrintPreviewDialog
PrintPreviewDialog printPreviewDialog = new PrintPreviewDialog();

// Set the PrintDocument for the PrintPreviewDialog
printPreviewDialog.Document = doc.PrintDocument;

// Set the size of the PrintPreviewDialog's client area
printPreviewDialog.ClientSize = new Size(600, 800);

// Show the PrintPreviewDialog
printPreviewDialog.ShowDialog();

// Dispose of the document object when finished using it
doc.Dispose();
```

---

# Spire.Doc C# Custom Paper Size
## Setting custom paper size for document printing
```csharp
// Create a new instance of Document
Document doc = new Document();

// Get the PrintDocument associated with the document
PrintDocument printDoc = doc.PrintDocument;

// Set the paper size of the default page settings to a custom size
printDoc.DefaultPageSettings.PaperSize = new PaperSize("custom", 900, 800);

// Print the document
printDoc.Print();
```

---

# spire.doc csharp print
## print word document using print dialog
```csharp
// Create a new instance of Document
Document document = new Document();

// Load the Word document
document.LoadFromFile(templatePath);

// Create a new PrintDialog
PrintDialog dialog = new PrintDialog();

// Allow printing of the current page
dialog.AllowCurrentPage = true;

// Allow printing of a range of pages
dialog.AllowSomePages = true;

// Use the system's default print dialog for selecting printer settings
dialog.UseEXDialog = true;

// Set the PrintDialog property of the document to the created PrintDialog
document.PrintDialog = dialog;

// Set the PrintDocument property of the PrintDialog to the document's PrintDocument
dialog.Document = document.PrintDocument;

// Print the document using the PrintDialog
dialog.Document.Print();

// Dispose of the document object when finished using it
document.Dispose();
```

---

# Spire.Doc C# Print via XPS
## Print Word document using XPS printing
```csharp
// Create a new MemoryStream for storing the document as XPS
using (MemoryStream ms = new MemoryStream())
{
    // Instantiate a new Document object
    using (Document document = new Document())
    {
        // Load the Word document
        document.LoadFromFile("document.docx");
        
        // Save the document to the MemoryStream as XPS format
        document.SaveToStream(ms, FileFormat.XPS);
    }

    // Reset the position of the MemoryStream to the beginning
    ms.Position = 0;

    // Specify the printer name to be used for printing
    String printerName = "Printer Name";

    // Print the XPS document using the specified printer and job name
    XpsPrint.XpsPrintHelper.Print(ms, printerName, "Printing Job", true);
}
```

---

# Spire.Doc C# Print Multiple Pages to One Sheet
## Demonstrates how to print multiple pages of a Word document onto a single sheet
```csharp
// Create a new instance of Document
Document doc = new Document();

// Create a new PrintDialog from System.Windows.Forms
System.Windows.Forms.PrintDialog printDialog = new System.Windows.Forms.PrintDialog();

// Enable printing to a file
printDialog.PrinterSettings.PrintToFile = true;

// Set the print file name
printDialog.PrinterSettings.PrintFileName = "output.xps";

// Assign the PrintDialog to the document's PrintDialog
doc.PrintDialog = printDialog;

// Print the document with multiple pages condensed into one sheet
doc.PrintMultipageToOneSheet(PagesPerSheet.FourPages, true);

doc.Dispose();
```

---

# spire.doc csharp printing
## print multiple copies of a document
```csharp
// Set the printer name to "Microsoft Print to PDF" for printing
document.PrintDocument.PrinterSettings.PrinterName = "Microsoft Print to PDF";

// Set the number of copies to be printed to 4
document.PrintDocument.PrinterSettings.Copies = 4;

// Print the document
document.PrintDocument.Print();
```

---

# Spire.Doc C# Print Document
## Print document without showing print dialog
```csharp
// Create a new instance of Document
Document doc = new Document();

// Load the Word document
doc.LoadFromFile(input);

// Get the PrintDocument associated with the document
PrintDocument printDoc = doc.PrintDocument;

// Set the print controller to StandardPrintController for silent printing
printDoc.PrintController = new StandardPrintController();

// Print the document
printDoc.Print();

// Dispose of the document object when finished using it
doc.Dispose();
```

---

# spire.doc csharp print settings
## Set margin and duplex options for document printing
```csharp
// Create a new instance of Document
Document doc = new Document();

// Load the Word document from the specified input file
doc.LoadFromFile(input);

// Get the PrintDocument associated with the document
PrintDocument printDoc = doc.PrintDocument;

// Set the OriginAtMargins property to true to align the printable area with the margins
printDoc.OriginAtMargins = true;

// Set the Margins property of the DefaultPageSettings to zero to remove any margins
printDoc.DefaultPageSettings.Margins = new System.Drawing.Printing.Margins(0, 0, 0, 0);

// Set the Duplex property of PrinterSettings to Vertical for double-sided printing
printDoc.PrinterSettings.Duplex = Duplex.Vertical;

// Print the document
printDoc.Print();

// Dispose of the document object when finished using it
doc.Dispose();
```

---

# spire.doc csharp vba macro
## detect and remove VBA macros from Word document
```csharp
// Create a new instance of Document
Document document = new Document();

// Load the Word document
document.LoadFromFile(filePath);

// Check if the document contains VBA macros
if (document.IsContainMacro)
{
    // Clear/remove the VBA macros from the document
    document.ClearMacros();
}

// Save the modified document
document.SaveToFile(resultPath, FileFormat.Docm);

// Dispose of the document object
document.Dispose();
```

---

# spire.doc csharp macros
## load and save macro-enabled word documents
```csharp
// Create a new instance of Document
Document document = new Document();

// Load the Word document from the specified file that may contain VBA macros
document.LoadFromFile("Macros.docm", FileFormat.Docm);

// Save the document to a new file with the specified name and format (Docm for macro-enabled document)
document.SaveToFile("Sample.docm", FileFormat.Docm);

// Dispose of the document object when finished using it
document.Dispose();
```

---

# Spire.Doc C# Picture Caption
## Add captions to pictures in Word document
```csharp
// Create a new instance of Document
Document document = new Document();

// Add a new section to the document
Section section = document.AddSection();

// Add a paragraph to the section
Paragraph par1 = section.AddParagraph();
par1.Format.AfterSpacing = 10;

// Append an image (picture) to the paragraph
DocPicture pic1 = par1.AppendPicture(Image.FromFile("image_path"));
pic1.Height = 100;
pic1.Width = 120;

// Set the caption numbering format to "Number" and add a caption below the picture
CaptionNumberingFormat format = CaptionNumberingFormat.Number;
pic1.AddCaption("Figure", format, CaptionPosition.BelowItem);

// Add another paragraph to the section
Paragraph par2 = section.AddParagraph();

// Append another image (picture) to the paragraph
DocPicture pic2 = par2.AppendPicture(Image.FromFile("image_path"));
pic2.Height = 100;
pic2.Width = 120;

// Add a caption below the second picture
pic2.AddCaption("Figure", format, CaptionPosition.BelowItem);

// Enable field updating in the document
document.IsUpdateFields = true;
```

---

# spire.doc csharp table caption
## add caption to table in word document
```csharp
// Get the body of the first section in the document
Body body = document.Sections[0].Body;

// Get the first table in the body
Table table = body.Tables[0] as Table;

// Add a caption to the table with the "Table" label, numbering format as "Number", and position below the table
table.AddCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.BelowItem);

// Enable field updating in the document
document.IsUpdateFields = true;
```

---

# Spire.Doc C# Picture Caption Cross-Reference
## Create picture captions with cross-references in Word documents
```csharp
// Create a new instance of Document and add a section
Document document = new Document();
Section section = document.AddSection();

// Add a paragraph for the cross-reference
Paragraph firstPara = section.AddParagraph();

// Add a paragraph and picture
Paragraph par1 = section.AddParagraph();
par1.Format.AfterSpacing = 10;
DocPicture pic1 = par1.AppendPicture(null);
pic1.Height = 120;
pic1.Width = 120;

// Set the caption numbering format and add a caption below the picture
CaptionNumberingFormat format = CaptionNumberingFormat.Number;
IParagraph captionParagraph = pic1.AddCaption("Figure", format, CaptionPosition.BelowItem);

// Add a bookmark at the specified location
string bookmarkName = "Figure_2";
Paragraph paragraph = section.AddParagraph();
paragraph.AppendBookmarkStart(bookmarkName);
paragraph.AppendBookmarkEnd(bookmarkName);

// Navigate to the bookmark and replace its content with the caption paragraph
BookmarksNavigator navigator = new BookmarksNavigator(document);
navigator.MoveToBookmark(bookmarkName);
TextBodyPart part = navigator.GetBookmarkContent();
part.BodyItems.Clear();
part.BodyItems.Add(captionParagraph);
navigator.ReplaceBookmarkContent(part);

// Create a cross-reference field for the bookmark
Field field = new Field(document);
field.Type = FieldType.FieldRef;
field.Code = @"REF Figure_2 \p \h";
firstPara.ChildObjects.Add(field);
FieldMark fieldSeparator = new FieldMark(document, FieldMarkType.FieldSeparator);
firstPara.ChildObjects.Add(fieldSeparator);

// Add the reference text
TextRange tr = new TextRange(document);
tr.Text = "Figure 2";
firstPara.ChildObjects.Add(tr);

FieldMark fieldEnd = new FieldMark(document, FieldMarkType.FieldEnd);
firstPara.ChildObjects.Add(fieldEnd);

// Enable field updating in the document
document.IsUpdateFields = true;
```

---

# Spire.Doc C# Caption with Chapter Number
## Set captions with chapter numbers for images in a Word document
```csharp
// Get the first section of the document
Section section = document.Sections[0];

// Specify the base name for the captions
string name = "Caption ";

// Iterate through paragraphs in the body of the section
for (int i = 0; i < section.Body.Paragraphs.Count; i++)
{
    // Iterate through child objects within each paragraph
    for (int j = 0; j < section.Body.Paragraphs[i].ChildObjects.Count; j++)
    {
        // Check if the child object is a picture
        if (section.Body.Paragraphs[i].ChildObjects[j] is DocPicture)
        {
            // Convert the child object to a DocPicture
            DocPicture pic1 = section.Body.Paragraphs[i].ChildObjects[j] as DocPicture;

            // Get the owner paragraph's owner, which should be the Body
            Body body = pic1.OwnerParagraph.Owner as Body;

            if (body != null)
            {
                // Find the index of the owner paragraph within the Body
                int imageIndex = body.ChildObjects.IndexOf(pic1.OwnerParagraph);

                // Create a new paragraph
                Paragraph para = new Paragraph(document);

                // Append the caption name
                para.AppendText(name);

                // Append a field for referencing the chapter number using a style reference
                Field field1 = para.AppendField("test", FieldType.FieldStyleRef);
                field1.Code = " STYLEREF 1 \\s ";

                // Append a separator text
                para.AppendText(" - ");

                // Append a sequence field for the caption number
                SequenceField field2 = (SequenceField)para.AppendField(name, FieldType.FieldSequence);
                field2.CaptionName = name;
                field2.NumberFormat = CaptionNumberingFormat.Number;

                // Insert the new paragraph after the owner paragraph
                body.Paragraphs.Insert(imageIndex + 1, para);
            }
        }
    }
}

// Enable field updating in the document
document.IsUpdateFields = true;
```

---

# Spire.Doc Table Caption Cross-Reference
## Create a table with caption and cross-reference field in a Word document
```csharp
// Add a table to the section
Table table = section.AddTable(true);
table.ResetCells(2, 3);

// Add a caption to the table
IParagraph captionParagraph = table.AddCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.BelowItem);

// Add a bookmark for the caption
string bookmarkName = "Table_1";
Paragraph paragraph = section.AddParagraph();
paragraph.AppendBookmarkStart(bookmarkName);
paragraph.AppendBookmarkEnd(bookmarkName);

// Replace bookmark content with caption
BookmarksNavigator navigator = new BookmarksNavigator(document);
navigator.MoveToBookmark(bookmarkName);
TextBodyPart part = navigator.GetBookmarkContent();
part.BodyItems.Clear();
part.BodyItems.Add(captionParagraph);
navigator.ReplaceBookmarkContent(part);

// Create a cross-reference field
Field field = new Field(document);
field.Type = FieldType.FieldRef;
field.Code = @"REF Table_1 \p \h";

// Add paragraph with cross-reference text
paragraph = section.AddParagraph();
TextRange range = paragraph.AppendText("This is a table caption cross-reference, ");
range.CharacterFormat.FontSize = 14;
paragraph.ChildObjects.Add(field);

// Add field separator and reference text
FieldMark fieldSeparator = new FieldMark(document, FieldMarkType.FieldSeparator);
paragraph.ChildObjects.Add(fieldSeparator);
TextRange tr = new TextRange(document);
tr.Text = "Table 1";
tr.CharacterFormat.FontSize = 14;
tr.CharacterFormat.TextColor = System.Drawing.Color.DeepSkyBlue;
paragraph.ChildObjects.Add(tr);

// Add field end mark
FieldMark fieldEnd = new FieldMark(document, FieldMarkType.FieldEnd);
paragraph.ChildObjects.Add(fieldEnd);

// Enable field updating
document.IsUpdateFields = true;
```

---

# Spire.Doc Fixed Layout Document Processing
## Extract layout information from document pages, lines, and paragraphs
```csharp
// Create a new instance of Document
Document doc = new Document();

// Load the document from the specified file
doc.LoadFromFile(inputFile, FileFormat.Docx);

// Create a FixedLayoutDocument from the loaded document
FixedLayoutDocument layoutDoc = new FixedLayoutDocument(doc);

// Get the first line in the first column of the first page
FixedLayoutLine line = layoutDoc.Pages[0].Columns[0].Lines[0];

// Create a StringBuilder to store the output text
StringBuilder stringBuilder = new StringBuilder();
stringBuilder.AppendLine("Line: " + line.Text);

// Get the paragraph that contains the line and append its text to the StringBuilder
Paragraph para = line.Paragraph;
stringBuilder.AppendLine("Paragraph text: " + para.Text);

// Get the text content of the first page
string pageText = layoutDoc.Pages[0].Text;
stringBuilder.AppendLine(pageText);

// Iterate through each page in the FixedLayoutDocument
foreach (FixedLayoutPage page in layoutDoc.Pages)
{
    // Get all the lines on the current page
    LayoutCollection<LayoutElement> lines = page.GetChildEntities(LayoutElementType.Line, true);

    // Append the page index and number of lines to the StringBuilder
    stringBuilder.AppendLine("Page " + page.PageIndex + " has " + lines.Count + " lines.");
}

// Append the lines of the first paragraph to the StringBuilder
// (except runs and nodes in the header and footer).
stringBuilder.AppendLine("The lines of the first paragraph:");
foreach (FixedLayoutLine paragraphLine in layoutDoc.GetLayoutEntitiesOfNode(((Section)doc.FirstChild).Body.Paragraphs[0]))
{
    stringBuilder.AppendLine(paragraphLine.Text.Trim());
    stringBuilder.AppendLine(paragraphLine.Rectangle.ToString());
}
```

---

# spire.doc csharp chart
## append bar chart to word document
```csharp
// Create a new instance of Document
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Bar chart.");

// Add a new paragraph to the section
Paragraph newPara = section.AddParagraph();

// Append a bar chart shape to the paragraph with specified width and height
ShapeObject chartShape = newPara.AppendChart(ChartType.Bar, 400, 300);
Chart chart = chartShape.Chart;

// Get the title of the chart
ChartTitle title = chart.Title;

// Set the text of the chart title
title.Text = "My Chart";

// Show the chart title
title.Show = true;

// Overlay the chart title on top of the chart
title.Overlay = true;
```

---

# Spire.Doc C# Bubble Chart
## Create and append a bubble chart to a Word document
```csharp
// Create a new instance of Document
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Bubble chart.");

// Add a new paragraph to the section
Paragraph newPara = section.AddParagraph();

// Append a bubble chart shape to the paragraph with specified width and height
ShapeObject shape = newPara.AppendChart(ChartType.Bubble, 500, 300);

// Get the chart object from the shape
Chart chart = shape.Chart;

// Clear any existing series in the chart
chart.Series.Clear();

// Add a new series to the chart with data points for X, Y, and bubble size values
ChartSeries series = chart.Series.Add("Test Series",
    new[] { 2.9, 3.5, 1.1, 4.0, 4.0 },
    new[] { 1.9, 8.5, 2.1, 6.0, 1.5 },
    new[] { 9.0, 4.5, 2.5, 8.0, 5.0 });
```

---

# Spire.Doc C# Chart
## Append Column Chart to Word Document
```csharp
// Create a new instance of Document
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Column chart.");

// Add a new paragraph to the section
Paragraph newPara = section.AddParagraph();

// Append a column chart shape to the paragraph with specified width and height
ShapeObject shape = newPara.AppendChart(ChartType.Column, 500, 300);

// Get the chart object from the shape
Chart chart = shape.Chart;

// Clear any existing series in the chart
chart.Series.Clear();

// Add a new series to the chart with data points for X values (categories) and Y values
chart.Series.Add("Test Series",
    new[] { "Word", "PDF", "Excel", "GoogleDocs", "Office" },
    new double[] { 1900000, 850000, 2100000, 600000, 1500000 });

// Set the number format for the Y-axis labels
chart.AxisY.NumberFormat.FormatCode = "#,##0";
```

---

# spire.doc csharp chart
## create and configure line chart in word document
```csharp
// Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Line chart.");

// Add a new paragraph to the section
Paragraph newPara = section.AddParagraph();

// Append a line chart shape to the paragraph with specified width and height
ShapeObject shape = newPara.AppendChart(ChartType.Line, 500, 300);

// Get the chart object from the shape
Chart chart = shape.Chart;

// Get the title of the chart
ChartTitle title = chart.Title;

// Set the text of the chart title
title.Text = "My Chart";

// Clear any existing series in the chart
ChartSeriesCollection seriesColl = chart.Series;
seriesColl.Clear();

// Define categories (X-axis values)
string[] categories = { "C1", "C2", "C3", "C4", "C5", "C6" };

// Add two series to the chart with specified categories and Y-axis values
seriesColl.Add("AW Series 1", categories, new double[] { 1, 2, 2.5, 4, 5, 6 });
seriesColl.Add("AW Series 2", categories, new double[] { 2, 3, 3.5, 6, 6.5, 7 });
```

---

# spire.doc csharp chart
## create pie chart in word document
```csharp
// Create a new instance of Document
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Pie chart.");

// Add a new paragraph to the section
Paragraph newPara = section.AddParagraph();

// Append a pie chart shape to the paragraph with specified width and height
ShapeObject shape = newPara.AppendChart(ChartType.Pie, 500, 300);
Chart chart = shape.Chart;

// Add a series to the chart with categories (labels) and corresponding data values
ChartSeries series = chart.Series.Add("Test Series",
    new[] { "Word", "PDF", "Excel" },
    new[] { 2.7, 3.2, 0.8 });
```

---

# spire.doc csharp chart
## append scatter chart to word document
```csharp
// Create a new instance of Document
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Scatter chart.");

// Add a new paragraph to the section
Paragraph newPara = section.AddParagraph();

// Append a scatter chart shape to the paragraph with specified width and height
ShapeObject shape = newPara.AppendChart(ChartType.Scatter, 450, 300);
Chart chart = shape.Chart;

// Clear any existing series in the chart
chart.Series.Clear();

// Add a new series to the chart with data points for X and Y values
chart.Series.Add("Scatter chart",
    new[] { 1.0, 2.0, 3.0, 4.0, 5.0 },
    new[] { 1.0, 20.0, 40.0, 80.0, 160.0 });
```

---

# spire.doc csharp chart
## create 3D surface chart in word document
```csharp
// Create a new instance of Document
Document document = new Document();

// Add a section to the document
Section section = document.AddSection();

// Add a paragraph to the section and append text to it
section.AddParagraph().AppendText("Surface3D chart.");

// Add a new paragraph to the section
Paragraph newPara = section.AddParagraph();

// Append a Surface3D chart shape to the paragraph with specified width and height
ShapeObject shape = newPara.AppendChart(ChartType.Surface3D, 500, 300);

// Get the chart object from the shape
Chart chart = shape.Chart;

// Clear any existing series in the chart
chart.Series.Clear();

// Set the title of the chart
chart.Title.Text = "My chart";

// Add multiple series to the chart with categories (X-axis values) and corresponding data values
chart.Series.Add("Series 1",
    new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Office" },
    new double[] { 1900000, 850000, 2100000, 600000, 1500000 });

chart.Series.Add("Series 2",
    new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Office" },
    new double[] { 900000, 50000, 1100000, 400000, 2500000 });

chart.Series.Add("Series 3",
    new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Office" },
    new double[] { 500000, 820000, 1500000, 400000, 100000 });
```

---


