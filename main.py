"""Tests to get text from a pdf."""

file = "/home/kobaya/Téléchargements/chargemap_20261002-07936931.pdf"


def test_pdfminer():
    from html.parser import HTMLParser
    from io import BytesIO
    from pdfminer.high_level import extract_text_to_fp, extract_text
    from pdfminer.layout import LAParams

    text = extract_text(file)
    print("-" * 50)
    print("pdfminer")
    print(text)

    output_bytes = BytesIO()
    with open(file, 'rb') as fin:
        extract_text_to_fp(fin, output_bytes, laparams=LAParams(),
                           output_type='html')
    html = output_bytes.getvalue().decode('utf-8')

    with open("output.html", "w", encoding="utf-8") as f:
        f.write(html)


    class MyHTMLParser(HTMLParser):
        def handle_starttag(self, tag, attrs):
            print("Encountered a start tag:", tag)

        def handle_endtag(self, tag):
            print("Encountered an end tag :", tag)

        def handle_data(self, data):
            print("Encountered some data  :", data)

    parser = MyHTMLParser()
    parser.feed(html)


def test_pikepdf():
    import pikepdf

    print("-" * 50)
    print("pikepdf")
    with pikepdf.open(file) as pdf:
        page1 = pdf.pages[0]
        print(page1)


if __name__ == "__main__":
    test_pdfminer()
    # test_pikepdf()
