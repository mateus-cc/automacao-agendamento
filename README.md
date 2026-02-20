# 🤖 Robô Extrator de Dados - Automação de Agendamentos

## 📌 Sobre o Projeto
Um sistema de automação desenvolvido em Python para otimizar o fluxo de agendamentos médicos (Ginecologia e Pediatria) na rede pública municipal. O robô substitui a digitação manual de guias médicas em PDF, extraindo dados críticos e alimentando uma base de dados centralizada no Google Sheets de forma 100% autônoma.

## 🚀 O Problema Resolvido
O processamento manual de guias de encaminhamento gerava gargalos de tempo e risco de erros de digitação. O sistema resolve isso baixando os PDFs diretamente da nuvem, lendo os dados (mesmo em documentos escaneados) e organizando tudo para a equipe de agendamento.

## ⚙️ Funcionalidades
* **Integração Google Drive:** Download automático de novos PDFs das pastas raízes e movimentação para a pasta "Processados" após a leitura.
* **Leitura Híbrida de PDFs:** Utiliza `pdfplumber` para PDFs digitais nativos e `PyTesseract` (OCR) para extrair texto de guias físicas escaneadas.
* **Filtros Avançados (Regex):** Extração precisa de dados não estruturados, lidando com ruídos de leitura para capturar: Nome, Prontuário, CNES, Unidade Solicitante e Classificação de Risco.
* **Integração Google Sheets:** Envio dos dados estruturados via `pandas` diretamente para a aba correta da planilha da fila de espera.
* **Interface Gráfica:** Desenvolvida com `Tkinter` para permitir o uso por usuários não técnicos através de um executável `.exe`.

## 🛠️ Tecnologias Utilizadas
* **Python 3** (Pandas, Re, OS, Shutil)
* **PyTesseract & pdf2image** (Motor de OCR)
* **pdfplumber** (Leitura de PDF)
* **Google APIs** (Google Drive API, gspread, oauth2client)
* **Tkinter** (GUI)
* **PyInstaller** (Geração do executável standalone)

## 🏗️ Arquitetura
O projeto foi refatorado utilizando o **Princípio DRY (Don't Repeat Yourself)**, operando através de um motor de repetição iterável que permite escalar a leitura para dezenas de novos departamentos médicos no futuro alterando apenas um dicionário de variáveis.
