import re
import io
import requests
import docx
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.constants import CONTENT_TYPE as CT
import random
from docx.shared import Cm
from docx.enum.table import WD_ALIGN_VERTICAL # pylint: disable=no-name-in-module
from docx.enum.text import WD_BREAK, WD_COLOR_INDEX
from docx.oxml.xmlchemy import OxmlElement
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.text.paragraph import Paragraph
from docx.text.run import Run

#### Code rewritten and adapted to handle footnotes from baloo-docx ####
from docx.opc.part import PartFactory
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart

from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.opc.constants import NAMESPACE
from docx.oxml.xmlchemy import (
    BaseOxmlElement, RequiredAttribute, ZeroOrMore, ZeroOrOne
)


class CT_Footnotes(BaseOxmlElement):
    """
    A ``<w:footnotes>`` element, a container for Footnotes properties 
    """

    footnote = ZeroOrMore ('w:footnote', successors=('w:footnotes',))

    @property
    def _next_id(self):
        ids = self.xpath('./w:footnote/@w:id')

        return int(ids[-1]) + 1
    
    def add_footnote(self):
        _next_id = self._next_id
        footnote = CT_Footnote.new(_next_id)
        footnote = self._insert_footnote(footnote)
        return footnote

    def get_footnote_by_id(self, _id):
        namesapce = NAMESPACE().WML_MAIN
        for fn in self.findall('.//w:footnote', {'w':namesapce}):
            if fn._id == _id:
                return fn
        return None
        
class CT_Footnote(BaseOxmlElement):
    """
    A ``<w:footnote>`` element, a container for Footnote properties 
    """
    _id = RequiredAttribute('w:id', ST_DecimalNumber)
    p = ZeroOrOne('w:p', successors=('w:footnote',))

    @classmethod
    def new(cls, _id):
        footnote = OxmlElement('w:footnote')
        footnote._id = _id
        return footnote
    
    def _add_p(self, text):
        _p = OxmlElement('w:p')
        pPr = _p.get_or_add_pPr()
        rstyle = pPr.get_or_add_pStyle()
        rstyle.val = 'FootnoteText'
        
        _r = _p.add_r()
        rPr = _r.get_or_add_rPr()
        rstyle = rPr.get_or_add_rStyle()
        rstyle.val = 'FootnoteReference'
        ref = OxmlElement('w:footnoteRef')
        _r.append(ref)
        _r = _p.add_r()
        ref = OxmlElement('w:footnoteRef')
        _r.append(ref)
        
        run = Run(_r, self)
        run.text = text
        
        self._insert_p(_p)
        return _p

    def _add_p_with_paragraph(self, para):
        _p = para._p
        # paragraph footnote style
        pPr = _p.get_or_add_pPr()
        rstyle = pPr.get_or_add_pStyle()
        rstyle.val = 'FootnoteText'
        # run style (with id of run)
        new_run_element = _p._new_r()
        para.runs[0]._element.addprevious(new_run_element)
        rPr = new_run_element.get_or_add_rPr()
        rstyle = rPr.get_or_add_rStyle()
        rstyle.val = 'FootnoteReference'
        ref = OxmlElement('w:footnoteRef')
        new_run_element.append(ref)
        self._insert_p(_p)
        return _p
    
    @property
    def paragraph(self):
        return Paragraph(self.p, self)
    
class CT_FNR(BaseOxmlElement):
    _id = RequiredAttribute('w:id', ST_DecimalNumber)

    @classmethod
    def new (cls, _id):
        footnoteReference = OxmlElement('w:footnoteReference')
        footnoteReference._id = _id
        return footnoteReference

class CT_FootnoteRef (BaseOxmlElement):

    @classmethod
    def new (cls):
        ref = OxmlElement('w:footnoteRef')
        return ref       
class FootnotesPart(XmlPart):
    """
    Definition of Footnotes Part
    """
    @classmethod
    def default(cls, package):
        partname = PackURI("/word/footnotes.xml")
        content_type = CT.WML_FOOTNOTES
        element = parse_xml(cls._default_footnotes_xml())
        return cls(partname, content_type, element, package)
docx.oxml.register_element_cls('w:footnotes', CT_Footnotes)
docx.oxml.register_element_cls('w:footnote', CT_Footnote)
docx.oxml.register_element_cls('w:footnoteReference', CT_FNR)
docx.oxml.register_element_cls('w:footnoteRef', CT_FootnoteRef)
PartFactory.part_type_for[CT.WML_FOOTNOTES] = FootnotesPart
##### END OF FOOTNOTES CODE ####

footnotes = {}

def convertMarkdownInFile(infile, outfile, styles_names=None):
    default_styles_names = {
        "Hyperlink": "Hyperlink",
        "Code": "Code",
        "Code Car": "Code Car",
        "BulletList": "BulletList",
        "Cell": "Cell",
        "Header": "Header"
    }
    if styles_names:
        for key, val in styles_names.items():
            default_styles_names[key] = val
    document = docx.Document(infile)
    for style_name in default_styles_names.values():
        if style_name not in document.styles:
            return False, "Error in template. There is a style missing : "+str(style_name)
    markdownToWordInDocument(document, default_styles_names)
    document.save(outfile)
    return True, outfile
    
def markdownToWordInDocument(document, styles_names=None):
    default_styles_names = {
        "Hyperlink": "Hyperlink",
        "Code": "Code",
        "Code Car": "Code Car",
        "BulletList": "BulletList",
        "Cell": "Cell"
    }
    if styles_names:
        for key, val in styles_names.items():
            default_styles_names[key] = val
    ps = getParagraphs(document)
    state = "normal"
    for paragraph in ps:
        state = markdownToWordInParagraph(document, paragraph, styles_names, state)
    ps = getParagraphs(document)
    for paragraph in ps:
        state = markdownToWordInParagraphCar(document, paragraph, styles_names, state)
    
def getParagraphs(document):
    """ Retourne un generateur pour tous les paragraphes du document.
        La page d'entête n'étant pas incluse dans documents.paragraphs."""
    body = document._body._body # pylint: disable=protected-access
    ps = body.xpath('//w:p')
    for p in ps:
        yield Paragraph(p, document._body) # pylint: disable=protected-access

def split_run_in_two(paragraph, run, split_index):
    index_in_paragraph = paragraph._p.index(run.element) # pylint: disable=protected-access
    text_before_split = run.text[0:split_index]
    text_after_split = run.text[split_index:]
    run.text = text_before_split
    new_run = paragraph.add_run(text_after_split)
    copy_format_manual(run, new_run)
    paragraph._p[index_in_paragraph+1:index_in_paragraph+1] = [new_run.element] # pylint: disable=protected-access
    return [run, new_run]

def split_run_in_three(paragraph, run, split_start, split_end):
    first_split = split_run_in_two(paragraph, run, split_end)
    second_split = split_run_in_two(paragraph, run, split_start)
    return second_split + [first_split[-1]]

def copy_format_manual(runA, runB):
    fontB = runB.font
    fontA = runA.font
    fontB.bold = fontA.bold
    fontB.italic = fontA.italic
    fontB.underline = fontA.underline
    fontB.strike = fontA.strike
    fontB.subscript = fontA.subscript
    fontB.superscript = fontA.superscript
    fontB.size = fontA.size
    fontB.highlight_color = fontA.highlight_color
    fontB.color.rgb = fontA.color.rgb


def markdownArrayToWordList(document, paragraph, state):
    table_line_regex = re.compile(r"^\|(?:[^\|\n-]*\|)*\s*$", re.MULTILINE)
    matched = re.findall(table_line_regex, paragraph.text)
    if len(matched) == 0:
        return state
    nb_columns = len(matched[0].strip()[1:-1].split("|"))
    array = document.add_table(rows=len(matched), cols=nb_columns)
    for i_row, match in enumerate(matched):
        line = match.strip()
        columns = line[1:-1].split("|") # [1:-1] strip beginning and ending pipe
        if len(columns) != nb_columns:
            raise ValueError("The array with following headers : "+str(matched[0])+" is supposed to have "+str(nb_columns)+ \
                                " columns but the line "+str(line)+" has "+str(len(columns))+" columns")
        for i_column, column in enumerate(columns):
            cell = array.cell(i_row, i_column)
            fill_cell(document, cell, column)
    move_table_after(array, paragraph)
    delete_paragraph(paragraph)
    return state

def markdownUnorderedListToWordList(paragraph, style, state):
    regex = re.compile(r"^\s*[\*|\-|\+]\s([^\n]+)", re.MULTILINE)
    matched = re.findall(regex, paragraph.text)
    if len(matched) > 0:
        start = paragraph.text.index(matched[0])
        end = paragraph.text.index(matched[-1])+len(matched[-1])
        text_end = paragraph.text[end:]
        paragraph.text = paragraph.text[:start-2].strip() # -2 for list marker + space
        for match in matched:
            new_p = insert_paragraph_after(paragraph)
            new_p.style = "BulletList"
            r = new_p.add_run()
            r.add_text(match)
        if text_end.strip() != "":
            insert_paragraph_after(new_p, text_end)
        if paragraph.text.strip() == "":
            delete_paragraph(paragraph)
    return state

def mardownCodeBlockToWordStyle(paragraph, code_style, state):
    if paragraph.text.lstrip().startswith("```") and state != "code_block":
        state = "code_block"
        paragraph.text = paragraph.text.split("```")[0].strip()+"```".join(paragraph.text.split("```")[1:]).strip()
    if state == "code_block":
        paragraph.style = code_style
    if paragraph.text.strip().endswith("```") and state == "code_block":
        state = "normal"
        paragraph.text = "```".join(paragraph.text.split("```")[:-1]).strip()+paragraph.text.split("```")[-1].strip()
    return state

def markdownToWordInParagraph(document, paragraph, styles_names, state):
    
    state = markdownArrayToWordList(document, paragraph, state)
    state = markdownUnorderedListToWordList(paragraph, document.styles[styles_names.get("BulletList","BulletList")], state)
    state = mardownCodeBlockToWordStyle(paragraph, document.styles[styles_names.get("Code","Code")], state)
    return state


def markdownToWordInParagraphCar(document, paragraph, styles_names, state):
    
    header_style = None
    for x in document.styles:
        if x.name == styles_names.get("Header", "Header"):
            header_style = x
    if header_style is None:
        raise KeyError("No style named "+styles_names.get("Header", "Header"))
    code_style = document.styles[styles_names.get("Code Car", "Code Car")]
    hyperlink_style = document.styles[styles_names.get("Hyperlink", "Hyperlink")]
    markdownHeaderToWordStyle(paragraph, header_style)
    transform_marker(paragraph, "==", setHighlight)
    transform_marker(paragraph, "**", setBold)
    transform_marker(paragraph, "__", setBold)
    transform_marker(paragraph, "*", setItalic)
    transform_marker(paragraph, "_", setItalic)
    transform_marker(paragraph, "~~", setStrike)
    if "`" in paragraph.text:
        transform_marker(paragraph, "`", lambda para, run, match: setCode(para, run, match, code_style))
    #bookmarks {#bookmark}
    transform_regex(paragraph, r"({#)([^\}\n]*)(})(?!\w)", (delCar, setBookmark, delCar))
    lambda_hyperlink = lambda para, run, match: setHyperlink(para, run, match, hyperlink_style)
    # markdown hyper link in the format [text to display](link)
    transform_regex(paragraph, r"(?<!\!)(\[)([^\]|^\n]+)(\]\()([^\)|^\n]+)(\))", (delCar, lambda_hyperlink, delCar, delCar, delCar))
    # markdown image hyper link to incorporate in the format ![alt text to display](link)
    transform_regex(paragraph, r"(\!\[)([^\]|^\n]+)(\]\()([^\)|^\n]+\.(?:png|jpg|jpeg|gif))(\))", (delCar, linkImageToImage, delCar, delCar, delCar))
    # just make left hyperlinks clickable
    transform_regex(paragraph, r"(https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b(?:[-a-zA-Z0-9()@:%_\+.~#?&//=]*[-a-zA-Z0-9@:%_\+.~#?&//=]))", (lambda_hyperlink,))
    #footnotes
    #inline footnotes ^[footnote text]
    lambda_foot = lambda para, run, match: setInlineFootnote(document, para, run, match)
    transform_regex(paragraph, r"(\^\[)([^\]\n]*)(\])(?!\w)", (delCar, lambda_foot, delCar))
    # footnotes insertion [^footnote id name]
    lambda_declare_foot = lambda para, run, match: declareFootnote(document, para, run, match)
    transform_regex(paragraph, r"(\[\^)([^\]\n]*)(\])(?!\w|:)", (delCar, lambda_declare_foot, delCar))
    # footnotes text description [^footnote id name]: indented text with possibly many paragraphs
    if state.startswith("inFootnoteDefinition:"):
        if paragraph.text.strip() == "":
            state = "normal"
        else:
            footnote_id = state.split(":")[1]
            footnote = footnotes[footnote_id]
            pPr = paragraph._p.get_or_add_pPr()
            rstyle = pPr.get_or_add_pStyle()
            rstyle.val = 'FootnoteText'
            footnote._insert_p(paragraph._p)
    else:
        state = transform_regex(paragraph, r"^(\[\^)([^\]\n]*)(\]:)(?!\w)(.+(?:\n[ \t]+.+)*)", (delCar, delCar, delCar, defineFootnote))
    return state

def setHyperlink(paragraph, run, match, style=None):
    run.font.underline = True
    if style is not None:
        run.style = style
    try:
        link_text = match.group(2)
        link_url = match.group(4)
    except:
        link_text = match.group(0)
        link_url = match.group(0)

    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    external =  link_url.startswith("http")
    if external:
        part = paragraph.part
        r_id = part.relate_to(link_url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )
    else:
        hyperlink.set(docx.oxml.shared.qn('w:anchor'), link_url)
    index_in_paragraph = paragraph._p.index(run.element)
    run.element.text = link_text
    hyperlink.append(run.element)
    paragraph._p[index_in_paragraph:index_in_paragraph] = [hyperlink]
    # Delete this if using a template that has the hyperlink style in it
    return 0, True, "normal"

def add_footnote(document):
    _footnotes_part = document._part.part_related_by(RT.FOOTNOTES)
    footnotes_part = _footnotes_part.element
    footnote = footnotes_part.add_footnote()
    return footnote

def add_footnote_reference(run, footnote):
    rPr = run._r.get_or_add_rPr()
    rstyle = rPr.get_or_add_rStyle()
    rstyle.val = 'FootnoteReference'
    reference = OxmlElement('w:footnoteReference')
    reference._id = footnote._id
    run._r.append(reference)

def setInlineFootnote(document, paragraph, run, match):
    deletedCars = len(run.text)
    run.text = ""
    #footnotes
    footnote = add_footnote(document)
    footnote._add_p(" " + str(match.group(2)))
    # footnotes reference
    add_footnote_reference(run, footnote)
    
    return deletedCars, False, "normal"

def declareFootnote(document, paragraph, run, match):
    deletedCars = len(run.text)
    run.text = ""
    #footnotes
    footnote = add_footnote(document)
    
    global footnotes
    footnotes[match.group(2)] = footnote
    add_footnote_reference(run, footnote)

    return deletedCars, False, "normal"

def defineFootnote(paragraph, run, match):
    #footnotes
    footnote_id = match.group(2)
    footnote = footnotes[footnote_id]
    _p = footnote._add_p_with_paragraph(paragraph)
    return 0, False, "inFootnoteDefinition:"+str(footnote_id)


def setBookmark(paragraph, run, match):
    """Set the bookmark
    function adapted from https://stackoverflow.com/questions/57586400/how-to-create-bookmarks-in-a-word-document-then-create-internal-hyperlinks-to-t
    """
    tag = run._r
    deleted = len(run.text)
    run.text = ""
    id = random.randint(0, 100000)
    start = docx.oxml.shared.OxmlElement('w:bookmarkStart')
    start.set(docx.oxml.ns.qn('w:id'), str(id))
    start.set(docx.oxml.ns.qn('w:name'), match.group(2))
    tag.append(start)
    end = docx.oxml.shared.OxmlElement('w:bookmarkEnd')
    end.set(docx.oxml.ns.qn('w:id'), str(id))
    end.set(docx.oxml.ns.qn('w:name'), match.group(2))
    tag.append(end)
    return deleted, False, "normal"

def linkImageToImage(para, run, match):
    link_url = match.group(4)
    data = downloadImgData(link_url)
    if data is not None:
        text_len = len(run.text)
        run.text = ""
        run.add_picture(data, width=Cm(17.19))
        return text_len, False, "normal"
    return 0, False, "normal"

def setHighlight(para, run, match):
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return 0, False, "normal"

def setBold(para, run, match):
    run.bold = True
    return 0, False, "normal"

def setItalic(para, run, match):
    run.italic = True
    return 0, False, "normal"

def setStrike(para, run, match):
    run.font.strike = True
    return 0, False, "normal"

def setCode(para, run, match, style):
    run.style = style
    return 0, False, "normal"

def delCar(para, run, match):
    ret = len(run.text)
    run.text = ""
    para._p.remove(run.element)
    return ret, True, "normal"

def transform_marker(paragraph, marker, func,content_regex=None):
    marker = re.escape(marker)
    if content_regex is None:
        content_regex = r"[^"+re.escape(marker[0])+r"\n]*"
    regex = r"(?<!\w)"+"("+marker+")("+content_regex+")("+marker+")"+r"(?!\w)"
    return transform_regex(paragraph, regex, (delCar, func, delCar))


def transform_regex(paragraph, regex, funcs):
    deletedCars = 0
    state = "normal"
    # find every iteration of marker+content+marker in paragraph
    for match in re.finditer(regex, paragraph.text, re.MULTILINE):
        # get starting marker run index and ending marker run index
        positions = [x[0]-deletedCars for x in match.regs[1:]] # Get starting pos of match of each group
        positions.append(match.regs[0][1]-deletedCars)
        runsIndexes = getRunsIndexFromPositions(paragraph, positions)
        # find marker position in run and split
        prev = None
        for indexes in runsIndexes[::-1]:
            if indexes == -1:
                prev = [666,0] # force split
                continue
            # if previous marker is in-between too runs, merge them
            if prev is not None and  (indexes[0] != prev[0] and prev[1] > 0):
                paragraph.runs[indexes[0]].text += paragraph.runs[indexes[0]+1].text
                paragraph.runs[indexes[0]+1].text = ""
                paragraph._p.remove(paragraph.runs[indexes[0]+1]._r)
                prev = [666,0] # force split
            # Split runs if needed, (maybe always ?)
            if prev is not None and (indexes[0] == prev[0] or prev[1] == 0):
                split_run_in_two(paragraph, paragraph.runs[indexes[0]], indexes[1])
            

            prev = indexes
        startRun = runsIndexes[0][0]+1 # first run is previous text, can be empty. +1 to skip it
        # apply transformation func on all runs found
        for i,func in enumerate(funcs):
            if paragraph.runs[startRun+i].text == "":
                startRun +=1
            deleted_count, deleted_run, state = func(paragraph, paragraph.runs[startRun+i], match)
            deletedCars += deleted_count
            if deleted_run:
                startRun -= 1
    return state

def markdownHeaderToWordStyle(paragraph, header_style):
    for match in re.finditer(r"^#{1,6} (.+)$", paragraph.text, re.MULTILINE):
        paragraph.text = re.sub(r"^#{1,6} ", "",paragraph.text)
        paragraph.style = header_style


def getRunsIndexFromPositions(paragraph, positions):
    """returns a list of tuples (runIndex, positionInRun) for each caracter position in the list positions
    example:
        paragraph.text is "This is an *example* of a paragraph"
        runs look like this:
            [
                "This is an *e",
                "xample",
                "* of a pa",
                "ragraph"
            ]
        regex match on paragraph for \*example\* is (11, 20) 
        we want to know which run contains the 11th caracter  and the 20th caracter
        the 11th is in the run 0 at position 11 and the 20th is in the run 2 at position 1 (it's the space)
        so we return [(0,11), (2, 1)],
        2nd example:
        paragraph.text is "**This is bold text**"
        we want to match **.+**
        runs look like this:
        [
            "**This is bold text**"
        ]
        returns [(0, 0), (0, 20)]
        """
    countedLetters = 0
    prevCountedLetters = 0
    ret = [-1] * len(positions)
    for i, run in enumerate(paragraph.runs):
        prevCountedLetters = countedLetters
        countedLetters += len(run.text)
        for j, pos in enumerate(positions):
            if prevCountedLetters <= pos and pos < countedLetters:
                ret[j] = (i,pos-prevCountedLetters)
    return ret

def fill_cell(document, cell, text, font_color=None, bg_color=None, bold=False):
    """
    Fill a table's cell's background with a background color, a text and a font_color for this text
    Also sets the vertical alignement of every cell as centered.
        Args:
            cell: the cell we want to fill
            text: the text to be written inside the cell
        Optional Args:
            font_color: a new font color to use for this text at RGB format (docx rgb). Default is None
            bg_color: A backgroud color to use for the cell at hexa rgb format (FFFFFF is white) default is None.
                        The color is written in xml directly as python-docx does not give a function to do that.
    """
    while len(cell.paragraphs) > 0:
        delete_paragraph(cell.paragraphs[0])
    p = cell.add_paragraph(text)
    p.style = document.styles["Cell"]
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if p.runs:
        p.runs[0].bold = bold
        if font_color is not None:
            p.runs[0].font.color.rgb = font_color
    if bg_color is not None:
        shading_elm_1 = parse_xml((r'<w:shd {} w:fill="'+bg_color+r'"/>').format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shading_elm_1)  # pylint: disable=protected-access


def insertPageBreak(paragraph):
    run = paragraph.add_run()
    run.add_break(WD_BREAK.PAGE)

def move_table_after(table, paragraph):
    """
    Move a given table after a given paragraph
        Args:
            table: the table to move
            paragraph: the paragraph to put the table after.
    """
    tbl, p = table._tbl, paragraph._p  # pylint: disable=protected-access
    p.addnext(tbl)

def delete_paragraph(paragraph):
    """
    Delete a paragraph.
        Args:
            paragraph: the paragraph object to delete.
    """
    p = paragraph._element  # pylint: disable=protected-access
    try:
        p.getparent().remove(p)
    except Exception: # pylint: disable=broad-except
        print("No parent found for element "+str(p))
    p._p = p._element = None  # pylint: disable=protected-access


def downloadImgData(url):
    try:
        data = requests.get(url, timeout=3)
        if data.status_code != 200:
            return None
    except Exception as e:
        return None
    data = data.content
    data = io.BytesIO(data)
    return data

def getParagraphs(document):
    """ Retourne un generateur pour tous les paragraphes du document.
        La page d'entête n'étant pas incluse dans documents.paragraphs."""
    body = document._body._body # pylint: disable=protected-access
    ps = body.xpath('//w:p')
    for p in ps:
        yield Paragraph(p, document._body) # pylint: disable=protected-access


def set_hyperlink(paragraph, run, url, text, style):
    # This gets access to the document.xml.rels file and gets a new relation id value
    run.font.underline = True
    if style is not None:
        run.style = style
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    index_in_paragraph = paragraph._p.index(run.element)
    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )
    run.element.text = text
    hyperlink.append(run.element)
    paragraph._p[index_in_paragraph:index_in_paragraph] = [hyperlink]
    # Delete this if using a template that has the hyperlink style in it
    return hyperlink

def insert_paragraph_after(paragraph, text=None, style=None):
    """
    Insert a new paragraph after the given paragraph.
        Args:
            paragraph: the paragraph object after which the new paragraph will be created
        
        Optional Args:
            text: a string of text to write in the new paragraph, Default = None
            style: a style to be applied on the new paragraph, Default = None

        Returns:
            Returns a paragraph object for the new added paragraph
    """
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)  # pylint: disable=protected-access
    new_para = Paragraph(new_p, paragraph._parent)  # pylint: disable=protected-access
    if text is not None:
        new_para.add_run(text)
    if style is not None:
        new_para.style = style
    return new_para


if __name__ == '__main__':
    res, msg = convertMarkdownInFile("examples/in_document.docx", "examples/out_document.docx", {"Code Car":"CodeStyle"})
    if res:
        print("Success : output document path is "+msg)
    else:
        print("Error in document : "+msg)