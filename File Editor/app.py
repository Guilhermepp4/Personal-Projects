from docx import Document
from flask import Flask, render_template, request, send_file
import io
import tempfile
from win32com import client
import pythoncom

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/replacementText', methods=['POST'])
def replacementForm():
    # Acessar o arquivo enviado
    file = request.files.get('fileInput')
    
    if not file:
        return "You have to upload a file"

    # Acessar os textos de substituição
    replaceText1 = request.form.get('replaceText1')
    replaceText2 = request.form.get('replaceText2')
    replaceText3 = request.form.get('replaceText3')
    footer_text = request.form.get('footerOwner')
    village_Text = request.form.get('footerLocal')

    try:
        # Ler o arquivo diretamente
        document = Document(io.BytesIO(file.read()))
    except Exception as e:
        return f"Erro ao processar o arquivo: {e}", 500

    start_phrase = 'execução relativo à “'
    end_phrase = '” e cujo Dono de Obra '
    start_phrase2 = "Dono de Obra é "
    end_phrase2 = ", observa as normas legais "
    start_phrase3 = "Braga, "
    end_phrase3 = " 202"

    # Iterar pelos parágrafos para encontrar e substituir o texto
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
            


    # Salvar o documento modificado em um arquivo temporário
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_docx:
        document.save(temp_docx.name)
        temp_docx_path = temp_docx.name

    # Converter o arquivo docx para pdf usando win32com
    pdf_path = temp_docx_path.replace(".docx", ".pdf")
    
    try:
        # Inicializar o COM
        pythoncom.CoInitialize()
        
        # Iniciar o Word e converter o documento
        word = client.Dispatch("Word.Application")
        doc = word.Documents.Open(temp_docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 é o formato PDF
        doc.Close()
        word.Quit()
    except Exception as e:
        return f"Erro ao converter o arquivo para PDF: {e}", 500
    finally:
        pythoncom.CoUninitialize()

    # Enviar o arquivo PDF para o cliente
    return send_file(pdf_path, as_attachment=True, download_name="fileInput.pdf", mimetype="application/pdf")

if __name__ == '__main__':
    app.run(debug=True)
