"""
Document analyser uses the odfpy module: https://pypi.org/project/odfpy/

This script prints:
bookmark count, cell count, changetracking count, character count, 
comment count, draw count, frame count, hyperlink count, 
image count, non-whitespace character count, object count, OLE object count, 
page count, paragraph count, row count, sentence count, 
syllable count, table count, textbox count, word count, and paragraph styles.

"""

import odf
from odf.namespaces import TEXTNS
from odf.element import Element
from odf.opendocument import load
from odf import text,meta,office,draw


filename=input()

doc=load(filename)

print("\nDOCUMENT STATISTICS\n")
for stat in doc.getElementsByType(meta.DocumentStatistic):
	print("Cell count",stat.getAttribute('cellcount'))
	print("Character count:",stat.getAttribute('charactercount'))
	print("Draw count:",stat.getAttribute('drawcount'))
	print("Frame count:",stat.getAttribute('framecount'))	
	print("Image count:",stat.getAttribute('imagecount'))
	print("Non-whitespace character count:",stat.getAttribute('nonwhitespacecharactercount'))
	print("Object count:",stat.getAttribute('objectcount'))
	print("Object linking and embedding (OLE) object count:",stat.getAttribute('oleobjectcount'))
	print("Page count:",stat.getAttribute('pagecount'))
	print("Paragraph count:",stat.getAttribute('paragraphcount'))
	print("Row count:",stat.getAttribute('rowcount'))
	print("Sentence count:",stat.getAttribute('sentencecount'))
	print("Syllable count:",stat.getAttribute('syllablecount'))
	print("Table count:",stat.getAttribute('tablecount'))
	print("Word count:",stat.getAttribute('wordcount'))

#type counter for attributes not covered by odf.meta.DocumentStatistic
def type_counter(doc,type):
	count=0
	for element in doc.getElementsByType(type):
		count+=1
	return count

types={
	'Bookmark':text.Bookmark,
	'Changetracking':text.FormatChange,
	'Comment':office.Annotation,
	'Hyperlink':text.A,
	'Textbox':draw.TextBox
}

for key,value in types.items():
	print(key,'count:',type_counter(doc,value))

def paragraph_style(doc):
	i = 1
	for paragraph in doc.getElementsByType(text.P):
		print('Paragraph',i,'style:',paragraph.getAttribute('stylename'))
		i+=1

paragraph_style(doc)