import argparse
import sys

from markdowntodocx.markdownconverter import convertMarkdownInFile


def parse_style(value):
    if "=" not in value:
        raise argparse.ArgumentTypeError("style must be in the form KEY=STYLE_NAME, got: " + value)
    key, style_name = value.split("=", 1)
    return key.strip(), style_name.strip()


def read_file(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except OSError as e:
        raise argparse.ArgumentTypeError("cannot read image modifier file " + path + ": " + str(e))


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="markdowntodocx",
        description="Convert markdown inside a Word document (.docx) to docx styles.",
    )
    parser.add_argument("infile", help="input .docx file containing markdown")
    parser.add_argument("outfile", help="output .docx file")
    parser.add_argument(
        "-s", "--style", action="append", type=parse_style, default=[], metavar="KEY=STYLE_NAME",
        help="override a style name (repeatable), e.g. -s \"Code Car=CodeStyle\" -s BulletList=MyBullets. "
             "Keys: Hyperlink, Code, Code Car, BulletList, Cell, Header1-6, Table",
    )
    mermaid = parser.add_mutually_exclusive_group()
    mermaid.add_argument(
        "--mermaid-server", metavar="URL",
        help="mermaid.ink server used to render mermaid graphs (default: https://mermaid.ink/img/)",
    )
    mermaid.add_argument(
        "--mermaid-cli", metavar="PATH",
        help="mermaid CLI binary used to render mermaid graphs locally (typically mmdc)",
    )
    parser.add_argument(
        "--image-modifier", action="append", type=read_file, metavar="XML_FILE",
        help="file containing an XML effect to add to images in paragraphs styled \"ImageModifier\" (repeatable)",
    )
    args = parser.parse_args(argv)

    res, msg = convertMarkdownInFile(
        args.infile,
        args.outfile,
        dict(args.style) or None,
        mermaid_server_link=args.mermaid_server,
        mermaid_cli=args.mermaid_cli,
        image_modifier=args.image_modifier,
    )
    if not res:
        print(msg, file=sys.stderr)
        return 1
    print("Converted document saved to " + str(msg))
    return 0


if __name__ == "__main__":
    sys.exit(main())
