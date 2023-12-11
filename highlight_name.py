import fitz  # PyMuPDF


def highlight_keywords(pdf_path, keyword, output_path):
    # 打開PDF文件
    pdf_document = fitz.open(pdf_path)

    # 遍歷每一頁
    for page_number in range(pdf_document.page_count):
        page = pdf_document[page_number]

        # 在頁面上搜索關鍵字
        keyword_instances = page.search_for(keyword)

        # 遍歷每個關鍵字的實例並添加高亮
        for inst in keyword_instances:
            highlight = page.add_highlight_annot(inst)

            # 設置高亮顏色（這裡設置為黃色）
            highlight.set_colors({"stroke": (1, 1, 0), "fill": (1, 1, 0)})

    # 保存修改後的PDF
    pdf_document.save(output_path)
    pdf_document.close()


# 示例使用
pdf_path = "陳颺.pdf"
output_path = "陳颺_highlighted.pdf"
name = "陳颺"

highlight_keywords(pdf_path, name, output_path)
