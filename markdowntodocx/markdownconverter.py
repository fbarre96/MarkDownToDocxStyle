import re
import io
import requests
import docx
from docx.shared import Cm
from docx.enum.table import WD_ALIGN_VERTICAL # pylint: disable=no-name-in-module
from docx.enum.text import WD_BREAK
from docx.oxml.xmlchemy import OxmlElement
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.text.paragraph import Paragraph


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
    markdownToWordInDocument(document, default_styles_names)
    document.save(outfile)
    return True
    
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
        for run in paragraph.runs:
            markdownToWordInRun(document, paragraph, run, styles_names)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        markdownToWordInRun(document, paragraph, run, styles_names)

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

def markdownHeaderToWordStyle(paragraph, run, style):
    regexHeaders = re.compile(r"^#{1,6} (.+)$", re.MULTILINE)
    matched = re.search(regexHeaders, run.text)
    if matched is not None:
        run.text = matched.group(1)
        paragraph.style = style
    return []

def markdownEmphasisToItalic(paragraph, initialRun):
    """
     de _Man-in-the-Middle_ entre le serveur
     # Absence de support de TLS_FALLBACK_SCSV

    TLS_FALLBACK_SCSV est une option permettant de mitiger les attaques dites de downgrade (type POODLE), afin d’empêcher un attaquant de forcer l’utilisation d’un protocole vulnérable lorsque des protocoles plus récents et sécurisés sont disponibles.

    """
    runs = [initialRun]
    i = 0
    while i < len(runs):
        splitted_runs = splitRunOnMarker(paragraph, runs[i], r"(?<!\w)\*([^\*\n]+)\*(?!\w)", "*")
        if len(splitted_runs) == 3:
            splitted_runs[1].italic = True
            runs.append(splitted_runs[1])
            runs.append(splitted_runs[2])
        else:
            splitted_runs = splitRunOnMarker(paragraph, runs[i], r"(?<!\w)\_([^\_\n]+)\_(?!\w)", "_")
            if len(splitted_runs) == 3:
                splitted_runs[1].italic = True
                runs.append(splitted_runs[1])
                runs.append(splitted_runs[2])
        i+=1
    return runs

def markdownStrongEmphasisToBold(paragraph, initialRun):
    runs = [initialRun]
    i = 0
    while i < len(runs):
        splitted_runs = splitRunOnMarker(paragraph, runs[i], r"(?<!\w)\*\*([^\*\n]+)\*\*(?!\w)", "**")
        if len(splitted_runs) == 3:
            splitted_runs[1].bold = True
            runs.append(splitted_runs[1])
            runs.append(splitted_runs[2])
        else:
            splitted_runs = splitRunOnMarker(paragraph, runs[i], r"(?<!\w)\_\_([^\_\n]+)\_\_(?!\w)", "__")
            if len(splitted_runs) == 3:
                splitted_runs[1].bold = True
                runs.append(splitted_runs[1])
                runs.append(splitted_runs[2])
        i+=1
    return runs

def markdownStrikeThroughToStrike(paragraph, initialRun):
    runs = [initialRun]
    i = 0
    while i < len(runs):
        splitted_runs = splitRunOnMarker(paragraph, runs[i], r"(?<!\w)~~([^~\n]+)~~(?!\w)", "~~")
        if len(splitted_runs) == 3:
            # THE PYDOCX run.strike does not work
            # Use run.font.strike
            splitted_runs[1].font.strike = True
            runs.append(splitted_runs[1])
            runs.append(splitted_runs[2])
        i+=1
    return runs

def markdownCodeToWordStyle(paragraph, initialRun, style):
    runs = [initialRun]
    i = 0
    while i < len(runs):
        splitted_runs = splitRunOnMarker(paragraph, runs[i], r"(?<!\w)\`([^\`\n]+)\`(?!\w)", "`")
        if len(splitted_runs) == 3:
            # THE PYDOCX run.strike does not work
            # Use run.font.strike
            splitted_runs[1].style = style
            runs.append(splitted_runs[1])
            runs.append(splitted_runs[2])
        i+=1
    return runs

def markdownImgToInsertedImage(paragraph, initialRun):
    runs = [initialRun]
    regex_hyperlink = r"\!\[([^\]|^\n]+)\]\(([^\)|^\n]+\.(?:png|jpg|jpeg|gif))\)"
    regex = re.compile(regex_hyperlink)
    i = 0
    while i < len(runs):
        matched = re.search(regex, runs[i].text)
        if matched is not None:
            start = runs[i].text.index(matched.group(0))
            end = start+len(matched.group(0))
            split_runs = split_run_in_three(paragraph, runs[i], start, end)
            data = downloadImgData(matched.group(2))
            if data is not None:
                split_runs[1].text = ""
                split_runs[1].add_picture(data, width=Cm(17.19))
            else:
                split_runs[1].text = split_runs[1].text.replace(matched.group(0), matched.group(1))
            
        i+=1
    return runs

def markdownLinkToHyperlink(paragraph, initialRun, style):
    runs = [initialRun]
    regex_hyperlink = r"(?<!\!)\[([^\]|^\n]+)\]\(([^\)|^\n]+)\)"
    regex = re.compile(regex_hyperlink)
    i = 0
    while i < len(runs):
        matched = re.search(regex, runs[i].text)
        if matched is not None:
            start = runs[i].text.index(matched.group(0))
            end = start+len(matched.group(0))
            split_runs = split_run_in_three(paragraph, runs[i], start, end)
            split_runs[1].text = split_runs[1].text.replace(matched.group(0), matched.group(1))
            set_hyperlink(paragraph, split_runs[1], matched.group(2), matched.group(1), style)
        i+=1
    return runs

def linkToHyperlinkStyle(paragraph, initialRun, style):
    runs = [initialRun]
    regex_hyperlink = r"https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.~#?&//=]*[-a-zA-Z0-9@:%_\+.~#?&//=])"
    regex = re.compile(regex_hyperlink)
    i = 0
    while i < len(runs):
        matched = re.search(regex, runs[i].text)
        if matched is not None:
            start = runs[i].text.index(matched.group(0))
            end = start+len(matched.group(0))
            split_runs = split_run_in_three(paragraph, runs[i], start, end)
            set_hyperlink(paragraph, split_runs[1], matched.group(0), matched.group(0), style)
        i+=1
    return runs


def splitRunOnMarker(paragraph, run, regexToSearch, markerToRemove):
    regex = re.compile(regexToSearch, re.MULTILINE)
    matched = re.findall(regex, run.text)
    for match in matched:
        start = run.text.index(markerToRemove+match+markerToRemove)
        end = start+len(markerToRemove+match+markerToRemove)
        split_runs = split_run_in_three(paragraph, run, start, end)
        split_runs[1].text = split_runs[1].text.replace(markerToRemove+match+markerToRemove, match)
        return split_runs
    return [run]

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

def markdownToWordInRun(document, paragraph, initialRun, styles_names):
    header_style = None
    for x in document.styles:
        if x.name == styles_names.get("Header", "Header"):
            header_style = x
    if header_style is None:
        raise KeyError("No style named "+styles_names.get("Header", "Header"))
    markdownHeaderToWordStyle(paragraph, initialRun, header_style)
    new_runs = set(markdownStrongEmphasisToBold(paragraph, initialRun))
    for run in list(new_runs):
        new_runs |= set(markdownEmphasisToItalic(paragraph, run))
    for run in list(new_runs):
        new_runs |= set(markdownStrikeThroughToStrike(paragraph, run))
    for run in list(new_runs):
        new_runs |= set(markdownCodeToWordStyle(paragraph, run, document.styles[styles_names.get("Code Car", "Code Car")]))
    for run in list(new_runs):
        new_runs |= set(markdownImgToInsertedImage(paragraph, run))
    for run in list(new_runs):
        markdownLinkToHyperlink(paragraph, run, document.styles[styles_names.get("Hyperlink", "Hyperlink")])
    for run in list(new_runs):
        linkToHyperlinkStyle(paragraph, run, document.styles[styles_names.get("Hyperlink", "Hyperlink")])





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
        data = requests.get(url, timeout=5)
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


