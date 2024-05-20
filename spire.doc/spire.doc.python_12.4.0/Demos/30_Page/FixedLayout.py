from spire.doc.common import *
from spire.doc import *

        
def WriteAllText(fpath:str,content:str):
    with open(fpath,'w',encoding="utf-8") as fp:
        fp.write(content)

# Specify the file path
# inputFile = "./Data/Sample.docx"
inputFile = "HelloWorld.docx"
outputFile = "output.txt"

# Create a new instance of Document
doc = Document()

# Load the document from the specified file
doc.LoadFromFile(inputFile, FileFormat.Docx)

# Create a FixedLayoutDocument object using the loaded document
layoutDoc = FixedLayoutDocument(doc)
result = ''

# Get the first line on the first page
line = layoutDoc.Pages[0].Columns[0].Lines[0]
result += "Line: "
result += line.Text
result += "\n"
# Retrieve the original paragraph associated with the line
para = line.Paragraph
result += "Paragraph text: "
result += para.Text
result += "\n"
# Retrieve all the text that appears on the first page in plain text format (including headers and footers).
pageText = layoutDoc.Pages[0].Text
result += pageText
result += "\n"
# Loop through each page in the document and print how many lines appear on each page.
pages = layoutDoc.Pages
for i in range(pages.Count):
	page = pages[i]
	lines = page.GetChildEntities(LayoutElementType.Line, True)
	result += "Page "
	result += str(page.PageIndex)
	result += " has "
	result += str(lines.Count)
	result += " lines."
	result += "\n"

# Perform a reverse lookup of layout entities for the first paragraph
result += "\n"
result += "The lines of the first paragraph:"
result += "\n"
tempChild = doc.FirstChild
section = Section(tempChild)
para = section.Body.Paragraphs[0]
paragraphLines = layoutDoc.GetLayoutEntitiesOfNode(para)

for i in range(paragraphLines.Count):
	tempLine = paragraphLines[i]
	paragraphLine = FixedLayoutLine(tempLine)
	result += (paragraphLine.Text).strip()
	result += "\n"
	result += paragraphLine.Rectangle.ToString()
	result += "\n"
	result += "\n"
# Write the extracted text to a file
WriteAllText(outputFile, result)

# Dispose of the document resources
doc.Dispose()
			