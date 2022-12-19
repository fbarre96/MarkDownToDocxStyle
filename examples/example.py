from markdowntodocx.markdownconverter import convertMarkdownInFile

res , msg = convertMarkdownInFile("examples/in_document.docx", "examples/out_document.docx", {"Code Car":"CodeStyle"})
