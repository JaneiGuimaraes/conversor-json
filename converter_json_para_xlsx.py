import json
import openpyxl
from openpyxl.styles import Alignment
from tkinter import Tk, filedialog, messagebox
import sys
import os

def selecionar_arquivo():
    root = Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="Selecione o arquivo JSON",
        filetypes=[("JSON", "*.json")]
    )

def formatar_descricoes(descricoes):
    """Consolida TODAS as descrições em uma única lista formatada"""
    textos = []
    for desc in descricoes if isinstance(descricoes, list) else [descricoes]:
        if isinstance(desc, dict):
            texto = desc.get('value', '')
            # Remove títulos e formata consistentemente
            linhas = [linha.strip() for linha in texto.split('\n') if linha.strip()]
            for linha in linhas:
                if linha.startswith('- '):
                    textos.append(f"• {linha[2:]}")
                elif ':' in linha:  # Para linhas tipo "Nº de Linhas: 3"
                    textos.append(f"• {linha}")
                else:  # Remove títulos como "Características Gerais..."
                    if not any(palavra in linha.lower() for palavra in ['características', 'especificações', 'detalhes']):
                        textos.append(f"• {linha}")
    return '\n'.join(textos)

def formatar_opcionais(opcionais):
    """Extrai TODOS os opcionais de forma organizada"""
    if not opcionais:
        return ""
    
    textos = []
    
    # Opções básicas
    if isinstance(opcionais, dict):
        if 'budgetPage' in opcionais:
            textos.append(f"• budgetPage: {opcionais['budgetPage']}")
        if 'productPage' in opcionais:
            textos.append(f"• productPage: {opcionais['productPage']}")
    
    # Opções complexas (quando é uma lista)
    if isinstance(opcionais, list):
        for grupo in opcionais:
            if isinstance(grupo, dict):
                nome_grupo = grupo.get('name', 'Opção')
                opcoes = grupo.get('optionals', [])
                for opcao in opcoes:
                    if isinstance(opcao, dict):
                        nome_opcao = opcao.get('name', '')
                        if nome_opcao:
                            textos.append(f"• {nome_grupo}: {nome_opcao}")
    
    return '\n'.join(textos)

def converter_json(json_path):
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            dados = json.load(f)
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Produtos"
        
        # Cabeçalhos
        ws.append(["PRODUTO", "DESCRIÇÃO COMPLETA", "OPCIONAIS"])
        
        # Estilo
        estilo = Alignment(wrap_text=True, vertical='top')
        
        # Processar cada produto
        for produto in dados:
            ws.append([
                produto.get('name', ''),
                formatar_descricoes(produto.get('descriptions', [])),
                formatar_opcionais(produto.get('optionals', []))
            ])
        
        # Ajuste de layout
        ws.column_dimensions['A'].width = 35  # Nome
        ws.column_dimensions['B'].width = 70  # Descrição
        ws.column_dimensions['C'].width = 40  # Opcionais
        
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = estilo
            # Altura dinâmica baseada no conteúdo
            num_linhas = max(
                len(str(row[1].value).split('\n')),
                len(str(row[2].value).split('\n'))
            )
            ws.row_dimensions[row[0].row].height = max(20, num_linhas * 15)
        
        # Salvar
        output_path = os.path.splitext(json_path)[0] + "_CONSOLIDADO.xlsx"
        wb.save(output_path)
        
        messagebox.showinfo(
            "Pronto!",
            f"Planilha gerada com:\n"
            f"- Todas descrições consolidadas\n"
            f"- Todos opcionais extraídos\n\n"
            f"Salvo em: {output_path}"
        )
    
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao processar:\n{str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    arquivo_json = selecionar_arquivo()
    if arquivo_json:
        converter_json(arquivo_json)