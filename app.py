from flask import Flask, request, render_template, send_file
import pandas as pd
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    if file:
        # Salvar o arquivo enviado
        filename = 'uploaded_file.xlsx'
        file.save(filename)

        # Processar o arquivo
        resumo_filename = process_file(filename)

        # Retornar o arquivo processado
        return send_file(resumo_filename, as_attachment=True)

def process_file(filename):
    # Carregar o arquivo Excel
    xls = pd.ExcelFile(filename)

    # Ler as abas relevantes
    dados = pd.read_excel(xls, "Dados")
    pgs_atuais = pd.read_excel(xls, "pgs atuais")

    # Verificar se as colunas esperadas existem
    required_columns = ["Nome do entregador", "Total NF", "Valor Ifood", "Tipo de Chave Pix", "Chave Pix", "CNPJ"]
    missing_columns = [col for col in required_columns if col not in pgs_atuais.columns]
    if missing_columns:
        raise KeyError(f"As seguintes colunas estão faltando na aba 'pgs atuais': {', '.join(missing_columns)}")

    # Selecionar colunas necessárias para o resumo
    resumo = pgs_atuais[["Nome do entregador", "Total NF", "Valor Ifood", "Tipo de Chave Pix", "Chave Pix", "CNPJ"]]

    # Salvar em um novo arquivo Excel
    resumo_filename = "resumo.xlsx"
    resumo.to_excel(resumo_filename, sheet_name="Resumo", index=False)

    return resumo_filename

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)