import pdfplumber
import io
from googleapiclient.http import MediaIoBaseDownload
import os
import pytesseract
from pdf2image import convert_from_path
import re
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tkinter as tk
from tkinter import filedialog, messagebox
from googleapiclient.discovery import build

#definindo variaveis
diretorio = r"C:\Users\mateus.costa\Desktop\processamentoRobo"
janela = tk.Tk()
janela.title('Robô Extrator de Dados para Agendamento')
janela.geometry("200x200")
    
def executar_robo():
        # --- COMEÇO DA CONEXÃO COM O GOOGLE ---

        #Define as permissões
    escopo = [
        "https://spreadsheets.google.com/feeds", 
        "https://www.googleapis.com/auth/drive"
        ]

        #Pega o crachá de acesso 
    credenciais = ServiceAccountCredentials.from_json_keyfile_name('credenciais.json', escopo)
    cliente = gspread.authorize(credenciais)
    servico_drive = build('drive', 'v3', credentials= credenciais)
    opcao_agendamento = [{'aba':'GO', 'id_raiz': '1SsLxR5MPNh7Di9XS0F5aU5GzDOC7ujls', 'id_proce': '1hC1CLWie33jkrKZpxVenME7LfyhZV8cz'}, {'aba':'Pediatra', 'id_raiz': '1FbBJYzdBkdJ17Juv2cz78Wwo9QGSiOW0', 'id_proce': '1gH6C3BJBLHsQroMmYmqL96ECv6sAL1wd'}]
    caminho_tesseract = r"C:\Users\mateus.costa\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
    caminho_poppler = r"C:\\poppler\\poppler-25.12.0\\Library\\bin"
    pytesseract.pytesseract.tesseract_cmd = caminho_tesseract
    
    total_pacientes_lidos = 0
    
    #para decidir qual pasta entrar
    for setor in opcao_agendamento:
        dados_lista = []
        resultados_drive = servico_drive.files().list(q=f"'{setor['id_raiz']}' in parents").execute()
        arquivos = resultados_drive.get('files', [])

        #pega o caminho completo do arquivo para extrair informaçoes
        for arquivo in arquivos:
            nome_do_pdf = arquivo.get('name')
            id_do_pdf = arquivo.get('id')
            print(f"Encontrei no Drive: {nome_do_pdf} (ID: {id_do_pdf})")
            if nome_do_pdf.endswith('.pdf'):
                caminho_completo = os.path.join(diretorio, nome_do_pdf)
                pedido = servico_drive.files().get_media(fileId=id_do_pdf)
                arquivo_fisico = io.FileIO(caminho_completo, 'wb')
                baixar = MediaIoBaseDownload(arquivo_fisico, pedido)
                concluido = False
                while not concluido:
                    status, concluido = baixar.next_chunk()
                    arquivo_fisico.close() 
                    print(f"Download concluído: {nome_do_pdf}")
                
                    # Tenta extrair texto digital
                    with pdfplumber.open(caminho_completo) as pdf:
                        pagina = pdf.pages[0]
                        texto_bruto = pagina.extract_text()
                        
                    # Se não houver texto digital, ativa o OCR
                    if not texto_bruto or not texto_bruto.strip():
                        print("Ativando OCR")
                        try:
                            lista_paginas = convert_from_path(caminho_completo, dpi=200, poppler_path=caminho_poppler)
                            texto_bruto = pytesseract.image_to_string(lista_paginas[0])
                        except Exception as e:
                            print(f"Falha no OCR: {e}")
                    else:
                        print(texto_bruto)

                    #A busca do prontuário dentro do bloco do PDF
                    var_prontuario = re.search(r'Prontuario.*?\b(\d{5,7})\b', texto_bruto, re.DOTALL)
                    if var_prontuario:
                        prontuario = var_prontuario.group(1)
                    
                    #Busca do nome com filtro anti-lixo
                    var_nome = re.search(r'Nome do cidadao.*?[\n\r]+(.*?)\s+\d{10,}', texto_bruto, re.DOTALL)    
                    if var_nome:
                        nome_sujo = var_nome.group(1)
                        linhas_nome = nome_sujo.split('\n')
                        nome = "" 
                        
                        for linha in linhas_nome:
                            linha_limpa = linha.strip()
                            # Pula a linha se for lixo de cabeçalho, CNS ou muito curta
                            if "CNS" in linha_limpa or "Classifica" in linha_limpa or "Idade" in linha_limpa or len(linha_limpa) <= 3:
                                continue
                            
                            # Se passou pelo filtro, é o nome! Limpamos números e barras (ex: "5 |")
                            nome = re.sub(r'[0-9\|/]', '', linha_limpa).strip()
                            if nome:
                                break 

                    #Busca do CNES paciente
                    var_cnes = re.search(r'Nome do cidadao.*?[\n\r]+.*?\s(\d{15})', texto_bruto, re.DOTALL)        
                    if var_cnes:
                        cnes = var_cnes.group(1)
                                
                    #Busca do nome da unidade
                    var_unidade = re.search(r'Nome:\s*(.*?)\s*CNES', texto_bruto, re.DOTALL)        
                    if var_unidade:
                        nome_unidade = var_unidade.group(1)

                    #Busca da classificação
                    var_classificacao = re.search(r'Classificagao de risco.*?[\n\r]+.*?\s(\w+)$', texto_bruto, re.MULTILINE)        
                    if var_classificacao:
                        classificacao = var_classificacao.group(1)
                        
                    dados_paciente = {'Focus': prontuario, 'Nome': nome, 'CPF': "", 'CNES': cnes, 'Unidade': nome_unidade, 'Status':"Pendente", 'Classificação Risco': classificacao}
                    dados_lista.append(dados_paciente)
                    print(dados_lista)
                    
                    os.remove(caminho_completo)
                    servico_drive.files().update(fileId=id_do_pdf, addParents=setor['id_proce'], removeParents=setor['id_raiz']).execute()
                    print(f"Arquivo {nome_do_pdf} movido para a pasta Processados com sucesso!")

        if len(dados_lista) > 0:
            total_pacientes_lidos += len(dados_lista)
            df = pd.DataFrame(dados_lista)
            print('Montando a tabela e conectando ao Google Drive')

                    # Abre a planilha e escolhe a aba 
            planilha = cliente.open_by_url('https://docs.google.com/spreadsheets/d/1BnE-95hdiTExrgFktQJnuY1u7fTnRTHEIVev_5hL1Xs/edit?gid=1377247306#gid=1377247306')
            aba = planilha.worksheet(setor['aba']) 

                    # Transforma o DataFrame em uma lista de linhas, trocando os vazios por texto em branco
            dados_para_enviar = df.fillna("").values.tolist()

                    # 5. Escreve na última linha vazia da planilha
            aba.append_rows(dados_para_enviar, value_input_option='USER_ENTERED')

            print(f"Sucesso! {len(dados_para_enviar)} pacientes enviados para a planilha 'Lista de espera', aba {setor['aba']}.")
        else:
            print(f"Nenhum arquivo novo na pasta {setor['aba']}")
        
    messagebox.showinfo("Concluído", f"{total_pacientes_lidos} pacientes processados e enviados com sucesso!")
    
    # Encerra a janela gráfica e o programa
    janela.destroy()

botao_iniciar = tk.Button(janela, text='Iniciar leitura dos arquivos', command= executar_robo)
botao_iniciar.pack(pady=50)

janela.mainloop()