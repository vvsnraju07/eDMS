import os
import fitz  # PyMuPDF
from flask import Flask, request, render_template, send_from_directory, jsonify
from fastapi import FastAPI, Query
app = Flask(__name__)
app = FastAPI()
def load_pdfs(folder_path):
    documents = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf') and not filename.endswith('_highlighted.pdf'):
            file_path = os.path.join(folder_path, filename)
            try:
                doc = fitz.open(file_path)
                full_text = ""
                for page_num in range(len(doc)):
                    page = doc.load_page(page_num)
                    text = page.get_text("text")  # Extract plain text
                    full_text += text
                documents.append({"filename": filename, "filepath": file_path, "content": full_text})
                doc.close()
            except Exception as e:
                print(f'Error processing {filename}: {str(e)}')
    return documents

def count_keyword_occurrences(documents, keyword):
    keyword_counts = {}
    keyword_contexts = {}
    for document in documents:
        text = document["content"]
        count = 0
        contexts = []
        start = 0
        while (index := text.lower().find(keyword.lower(), start)) != -1:
            count += 1
            start = index + len(keyword)
            # Extract context
            tokens = text.split()
            word_index = text[:index].count(' ')
            start_context = max(word_index - 5, 0)
            end_context = min(word_index + 11, len(tokens))
            context = ' '.join(tokens[start_context:end_context])
            contexts.append(context)
        if count > 0:
            keyword_counts[document["filename"]] = count
            keyword_contexts[document["filename"]] = contexts
    return keyword_counts, keyword_contexts

def sort_files_by_frequency(keyword_counts):
    sorted_files = sorted(keyword_counts.items(), key=lambda item: item[1], reverse=True)
    return sorted_files

def highlight_keyword_in_pdf(file_path, keyword):
    base, ext = os.path.splitext(file_path)
    highlighted_file_path = f"{base}_{keyword}_highlighted{ext}"
    
    if os.path.exists(highlighted_file_path):
        return highlighted_file_path
    
    doc = fitz.open(file_path)
    keyword_lower = keyword.lower()
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text_instances = page.search_for(keyword, quads=False)  # `quads=False` ensures rectangles, not quads
        for inst in text_instances:
            highlight = page.add_highlight_annot(inst)
            highlight.update()
    
    doc.save(highlighted_file_path, garbage=4, deflate=True)
    doc.close()
    return highlighted_file_path

@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')

@app.route('/search', methods=['POST'])
def search():
    keyword = request.form['keyword']
    folder_name = request.form['folder_name']
    folder_path = os.path.join('C:\\Users\\25865\\Desktop\\search', folder_name)
    documents = load_pdfs(folder_path)
    keyword_counts, keyword_contexts = count_keyword_occurrences(documents, keyword)
    sorted_files = sort_files_by_frequency(keyword_counts)

    result = []
    for filename, count in sorted_files:
        for doc in documents:
            if doc["filename"] == filename:
                highlighted_file_path = highlight_keyword_in_pdf(doc["filepath"], keyword)
                result.append({
                    'filename': filename,
                    'count': count,
                    'highlighted_path': os.path.relpath(highlighted_file_path, start='C:\\Users\\25865\\Desktop\\search'),
                    'contexts': keyword_contexts[filename]
                })
    return jsonify(result)

@app.route('/highlighted_pdfs/<path:filename>')
def download_file(filename):
    folder_path = 'C:\\Users\\25865\\Desktop\\search'
    return send_from_directory(folder_path, filename)

if __name__ == "__main__":
    app.run(debug=True)
