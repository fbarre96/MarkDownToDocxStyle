# MarkDownToDocxStyle
Convert Markdown inside Office Word documents

## Installation

`pip install markdowntodocx`

## Usage



**to convert an existing Docx file:**

see examples/example.py

```
from markdowntodocx.markdownconverter import convertMarkdownInFile

convertMarkdownInFile("/mypath/to/document.docx", "output_path.docx", {"Code Car":"CodeStyle"})
```

**To convert a python-docx Document object:**

```
from markdowntodocx.markdownconverter import convertMarkdownInFile

res, msg = convertMarkdownInFile("examples/in_document.docx", "examples/out_document.docx", {"Code Car":"CodeStyle"})
if res:
    print("Success : output document path is "+msg)
else:
    print("Error in document : "+msg)
```

## Styles and considerations
    You have to define styles in you word document in order to use Markdown **Headers/titles**, **Hyperlinks**, **Code formatting**, **Arrays**, **Unordered List**.

This styles are either standard markdown or come from extended markdown : https://www.markdownguide.org/extended-syntax/
    

* Emphasis (*italic*) `*Text*` or `_Text_`:  converts to word italic
* Strong Emphasis (**Bold**) `**Text**` or `__Text__`:  converts to word bold
* Strike through (~~Strike~~) `~~Strike~~` : converts to word strike through style
* Highlight (==highlight==) `==Highlight==' : converts to word Yellow highlight. 
* Header `# MarkdownHeader1` to `###### MarkdownHeader6`: 
    * Must be in alone in a paragraph. IF NOT, the rest will be erased. 
    * It will use the document style named "Header" by default. 
    * You can specify another style by giving the style dictionnary as last arg for both functions. 
    * E.g : `res, msg = convertMarkdownInFile("examples/in_document.docx", "examples/out_document.docx", {"Header":"style_name"})`
* Inline Code `` `Text` `` (`my code`):
    * It will use the document style named "Code" (Caracter format) by default. 
    * You can specify another style by giving the style dictionnary as last arg for both functions. 
    * E.g : `markdownToWordInFile("/mypath/to/document.docx", "output_path.docx", {"Code Car":"my_inline_code_style"})`
    
* Code Block ``` ` ` `T e x t` ` ` ``` 
```
my code
```

    * It will use the document style named "Code" by default. 
    * You can specify another style by giving the style dictionnary as last arg for both functions. 
    * E.g : `markdownToWordInFile("/mypath/to/document.docx", "output_path.docx", {"Code":"my_block_code_style"})`

* Insert Image ``![Image name](http://link.do.web/myimage.png)``:
    * It will download the image from the hyperlink and insert the picture with a width of 18cm

* Hyperlink `` [google](https://www.google.fr)  `` : Makes it a Word hyperlink [google](https://www.google.fr)
    * Will also attempt to convert any valid http hyperlink to word : `http://www.google.fr` -> http://www.google.fr
    * If the link does not start with http, it will be treated as an internal link to a bookmark

* Bookmark ``this will be bookmared with name bookmark1{#bookmark1}
    * You may hyperlink to it : ``[url text to display](bookmark1)``

* Footnotes (BETA) :
    * Inline foot notes : ``this is a conundrum^[https://fr.wiktionary.org/wiki/conundrum]``
    * External foot notes : ```
    This paragraph will have a footnote[^1]
    And this paragraph will have another[^2]
    [^1]: This is a footnote with markdown as well **bold**
        And it can have many lines if they are indented.
    [^2]: This is the second footnote
    ```

* Array to wordlist: (must be alone in a paragraph otherwise the rest  of the paragraph is deleted)
```
|Column1|column2|Column3|
|-------|-------|-------|
|line|line|line|
```
   --> 
|Column1|column2|Column3|
|-------|-------|-------|
|line|line|line|

    * Cells created will use the document style named "Cell" by default. 
    * You can specify another style by giving the style dictionnary as last arg for both functions. 
    * E.g : `markdownToWordInFile("/mypath/to/document.docx", "output_path.docx", {"Cell":"my_cell_style"})`

* Unordered List : (`- my list` or `* my list` or `+ my list`) : 
    * Must be in alone in a paragraph. IF NOT, the rest of the paragraph will be erased. 
    * It will use the document style named "Header" by default. 
    * You can specify another style by giving the style dictionnary as last arg for both functions. 
    * E.g : `markdownToWordInFile("/mypath/to/document.docx", "output_path.docx", {"BulletList":"my_bullet_style"})`
    
