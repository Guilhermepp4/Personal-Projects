from flask import Flask, render_template, request, send_file
import os
import tempfile
import zipfile
from docx import Document
from win32com import client
import pythoncom

app = Flask(__name__)
FILES_DIR = './files/'

@app.route('/')
def home():
    # List files in the directory
    files = [f for f in os.listdir(FILES_DIR) if os.path.isfile(os.path.join(FILES_DIR, f))]
    return render_template('index.html', files=files)

@app.route('/replacementText', methods=['POST'])
def replacementForm():
    # Get the list of selected filenames
    selected_files = request.form.getlist('selected')

    if selected_files.__len__() == 0:
        return "You have to select at least one file", 400

    # Access replacement texts
    replaceText1 = request.form.get('replaceText1')
    replaceText2 = request.form.get('replaceText2')
    replaceText3 = request.form.get('replaceText3')
    footer_text = request.form.get('footerOwner')
    village_Text = request.form.get('footerLocal')

    processed_files = []

    for file_name in selected_files:
        file_path = os.path.join(FILES_DIR, file_name)
        
        if not os.path.isfile(file_path):
            return f"File not found: {file_name}", 404

        try:
            # Load and process the document
            document = Document(file_path)

            start_phrase = 'execução relativo à “'
            end_phrase = '” e cujo Dono de Obra '
            start_phrase2 = "Dono de Obra é "
            end_phrase2 = ", observa as normas legais "
            start_phrase3 = "Braga, "
            end_phrase3 = " 202"

            for paragraph in document.paragraphs:
                text = paragraph.text

                # Primeira substituição
                start_index = text.find(start_phrase)
                if start_index != -1:
                    end_index = text.find(end_phrase, start_index + len(start_phrase))
                    if end_index != -1:
                        end_index += len(end_phrase)
                        new_text = text[:start_index] + start_phrase + replaceText1 + end_phrase + text[end_index:]
                        paragraph.text = new_text
                    else:
                        paragraph.text = text[:start_index] + replaceText1 + text[start_index + len(start_phrase):]

                # Atualizar o texto do parágrafo após a primeira substituição
                text = paragraph.text

                # Segunda substituição
                start_index = text.find(start_phrase2)
                if start_index != -1:
                    end_index = text.find(end_phrase2, start_index + len(start_phrase2))
                    if end_index != -1:
                        end_index += len(end_phrase2)
                        new_text = text[:start_index] + start_phrase2 + replaceText2 + end_phrase2 + text[end_index:]
                        paragraph.text = new_text
                    else:
                        paragraph.text = text[:start_index] + replaceText2 + text[start_index + len(start_phrase2):]
                
                # Atualizar o texto do parágrafo após a segunda substituição
                text = paragraph.text

                # Terceira substituição
                start_index = text.find(start_phrase3)
                if start_index != -1:
                    end_index = text.find(end_phrase3, start_index + len(start_phrase3))
                    if end_index != -1:
                        end_index += len(end_phrase3)
                        new_text = text[:start_index] + start_phrase3 + replaceText3 + end_phrase3 + text[end_index:]
                        paragraph.text = new_text
                    else:
                        paragraph.text = text[:start_index] + replaceText3 + text[start_index + len(start_phrase3):]
                
                # Quarta substituição
                if "AQUI" in text:
                    paragraph.text = text.replace("AQUI", "")
            
            # Atualizar o rodapé em cada seção
            run_Footer = False
            run_Village = False

            for section in document.sections:
                footer = section.footer
                for table in footer.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                if paragraph.text.strip():  # Verificar se a célula já contém texto
                                    if not run_Footer:
                                        paragraph.clear()  # Limpar o texto existente
                                        paragraph.add_run(footer_text)
                                        run_Footer = True
                                    elif not run_Village:
                                        paragraph.clear()
                                        paragraph.add_run(village_Text)
                                        run_Village = True
                                if run_Footer and run_Village:
                                    break
                            if run_Footer and run_Village:
                                break
                        if run_Footer and run_Village:
                            break
                    if run_Footer and run_Village:
                        break
                if run_Footer and run_Village:
                    break

            # Save the document as a temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_docx:
                document.save(temp_docx.name)
                temp_docx_path = temp_docx.name

            pdf_path = temp_docx_path.replace(".docx", ".pdf")
            
            try:
                pythoncom.CoInitialize()
                
                word = client.Dispatch("Word.Application")
                doc = word.Documents.Open(temp_docx_path)
                doc.SaveAs(pdf_path, FileFormat=17)  # 17 is the PDF format
                doc.Close()
                word.Quit()
            except Exception as e:
                return f"Error converting file to PDF: {e}", 500
            finally:
                pythoncom.CoUninitialize()

            processed_files.append(pdf_path)

        except Exception as e:
            return f"Error processing file {file_name}: {e}", 500

    if processed_files:
        # Create a ZIP file containing all processed PDF files
        zip_path = os.path.join(tempfile.gettempdir(), "processed_files.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for pdf_path in processed_files:
                zipf.write(pdf_path, os.path.basename(pdf_path))
        
        # Send the ZIP file
        return send_file(zip_path, as_attachment=True, download_name="processed_files.zip", mimetype="application/zip")

    return "No files were processed.", 400

if __name__ == '__main__':
    app.run(debug=True)
