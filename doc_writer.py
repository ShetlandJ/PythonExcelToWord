from doc_template import DocTemplate
import os


def write_doc(constituency, excel_book, template):
    data_written = False
    constituency_data = excel_book.get_constituency_data(constituency)

    d = DocTemplate()
    d.load(template.path)

    for year in constituency_data:
        year_data = constituency_data[year]

        for column_header in year_data:
            data_written = True
            value = year_data[column_header]
            if value:
                d.write_data(year, column_header, value.formatted)
            else:
                d.write_data(year, column_header, "-")

    if data_written:
        if not os.path.exists("Constituencies"):
            os.mkdir("Constituencies")
        d.set_title(constituency)
        d.save(os.path.join("Constituencies", constituency + ".docx"))