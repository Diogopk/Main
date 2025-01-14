from flask import Flask, request, render_template, send_file, flash, redirect, url_for
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # Necessário para usar o flash


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('Nenhum arquivo selecionado')
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        flash('Nenhum arquivo selecionado')
        return redirect(url_for('index'))

    if file:
        # Salvar o arquivo enviado
        filename = 'uploaded_file.xlsx'
        file.save(filename)

        try:
            # Processar o arquivo
            resumo_filename = process_file(filename)
        except KeyError as e:
            flash(str(e))
            return redirect(url_for('index'))

        # Retornar o arquivo processado
        return send_file(resumo_filename, as_attachment=True)


def process_file(filename):
    # Carregar o arquivo Excel
    xls = pd.ExcelFile(filename)

    # Ler as abas relevantes
    dados = pd.read_excel(xls, "Dados")
    pgs_atuais = pd.read_excel(xls, "pgs atuais")

    # Verificar se as colunas esperadas existem
    required_columns_pgs_atuais = ["Nome do entregador", "Valor NF", "Valor Ifood"]
    missing_columns_pgs_atuais = [col for col in required_columns_pgs_atuais if col not in pgs_atuais.columns]
    if missing_columns_pgs_atuais:
        raise KeyError(
            f"As seguintes colunas estão faltando na aba 'pgs atuais': {', '.join(missing_columns_pgs_atuais)}")

    required_columns_dados = ["Nome do entregador", "Tipo de Chave Pix", "Chave Pix", "CPF", "CNPJ"]
    missing_columns_dados = [col for col in required_columns_dados if col not in dados.columns]
    if missing_columns_dados:
        raise KeyError(f"As seguintes colunas estão faltando na aba 'Dados': {', '.join(missing_columns_dados)}")

    # Fazer a junção (merge) dos dados
    merged_data = pd.merge(pgs_atuais, dados, on="Nome do entregador", how="left")

    # Selecionar colunas necessárias para o resumo
    resumo = merged_data[
        ["Nome do entregador", "Valor NF", "Valor Ifood", "Tipo de Chave Pix", "Chave Pix", "CPF", "CNPJ"]]

    # Salvar em um novo arquivo Excel
    resumo_filename = "resumo.xlsx"
    resumo.to_excel(resumo_filename, sheet_name="Resumo", index=False)

    return resumo_filename


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
