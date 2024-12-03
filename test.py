import zipfile
import xml.dom.minidom


def extract_xml_from_docx(docx_path):
    with zipfile.ZipFile(docx_path, "r") as docx:
        xml_content = docx.read("word/document.xml")
        return xml_content.decode("utf-8")


def pretty_print_xml(xml_content):
    dom = xml.dom.minidom.parseString(xml_content)
    return dom.toprettyxml()


# Extract XML content from the .docx file
xml_content = extract_xml_from_docx("CC03.docx")

# Pretty print the XML content
pretty_xml_content = pretty_print_xml(xml_content)

# Write the pretty-printed XML content to an .xml file
with open("CC03.xml", "w") as xml_file:
    xml_file.write(pretty_xml_content)
