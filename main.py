from pdfminer.high_level import extract_text
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill

# Caminho para o arquivo PDF
pdf_path = "Relatório de Diagnóstico Mind The Bizz 2022.1.pdf"

# Extrai o texto de todas as páginas do arquivo PDF
text = extract_text(pdf_path)

# Divide o texto em páginas
pages = text.split('\x0c')[1:]  # Ignora a primeira página

# Cria um novo arquivo Excel
wb = Workbook()
ws = wb.active

# Adicionar rótulos das colunas (somente na primeira página)
col_labels = [
    "Nome do projeto/negócio mentorado",
    "Setor/Segmento de atuação",
    "Responsável 1 | Profissão | Telefone | E-mail",
    "Responsável 2 | Profissão | Telefone | E-mail",
    "Local de Origem",
    "Nível de maturidade do projeto",
    "Resumo Diagnóstico",
    "Status de Desenvolvimento de Entregas",
    "Nome do mentor responsável pelo projeto/negócio"
]

for col_num, label in enumerate(col_labels, start=1):
    cell = ws.cell(row=1, column=col_num, value=label)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")  # Cor cinza

# Percorre as páginas e preenche as células correspondentes
row_num = 2  # Inicia na segunda linha (após os rótulos das colunas)

for page in pages:
    # Separa o texto da página em linhas
    lines = page.split('\n')
    
    # Remove as três primeiras linhas da página
    lines = lines[3:]
    
    for line in lines:
        line = line.strip()  # Remove espaços em branco no início e no fim da linha
        
        if line and line not in col_labels:  # Verifica se a linha não está vazia e não é uma label repetida
            cell = ws.cell(row=row_num, column=col_num, value=line)
            col_num += 1  # Avança para a próxima coluna
            
            if col_num > len(col_labels):  # Verifica se chegou à última coluna
                row_num += 1  # Avança para a próxima linha
                col_num = 1  # Reinicia a contagem das colunas

# Salvar o arquivo Excel
wb.save("teste.xlsx")
