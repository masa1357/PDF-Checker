import pdfplumber
import fitz  # PyMuPDF
import pathlib

def search_files(path, extension):
    files = []
    for file in path.iterdir():
        if file.is_dir():
            files.extend(search_files(file, extension))
        elif file.suffix == extension:
            files.append(file)
    return files

def highlight_issues(pdf_path, issues):
    doc = fitz.open(pdf_path)
    for page_num, lines in issues.items():
        page = doc[page_num]
        for line in lines:
            # ハイライトする領域の座標を設定
            rect = fitz.Rect(line['x0'], line['top'], line['x1'], line['bottom'])
            highlight = page.add_highlight_annot(rect)
            highlight.update()
    new_filename = pathlib.Path(pdf_path).stem + "_highlighted.pdf"
    doc.save(new_filename)
    doc.close()
    return new_filename

def check(file, expected_indentation=72):
    issues = {}
    with pdfplumber.open(file) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                continue
            words = page.extract_words()
            problem_lines = []
            for word in words:
                if float(word['x0']) > expected_indentation:
                    problem_lines.append(word)
            if problem_lines:
                issues[i] = problem_lines
    return issues

def main(path):
    pdf_files = search_files(path, ".pdf")
    for file in pdf_files:
        print(f"Processing File: {file}")
        issues = check(file)
        if issues:
            highlighted_pdf = highlight_issues(file, issues)
            print(f"Highlighted PDF saved as: {highlighted_pdf}")
        else:
            print("No indentation issues found.")

if __name__ == "__main__":
    path = pathlib.Path.cwd()
    main(path)
