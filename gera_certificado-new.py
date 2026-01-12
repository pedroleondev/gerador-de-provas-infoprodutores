import pandas as pd
import os
from PIL import Image, ImageFont, ImageDraw
from time import sleep

# Lê o arquivo Excel
#df = pd.read_excel("ruivas-2024.xlsx", sheet_name="APROVADOS")
df = pd.read_excel("arquivos\\WORKSHOP RUIVO EM 2H.xlsx", sheet_name="APROVADAS")

# Obtém os nomes da coluna "NOME"
nomes = df["NOME"].astype(str)

# Obtém o diretório do script atual para criar o caminho relativo
script_dir = os.path.dirname(os.path.abspath(__file__))
certificado_nome = "certificado\\certificado.png"
caminho_certificado = os.path.join(script_dir, certificado_nome)

# Cria a pasta "alunos" se ela não existira
pasta_alunos = os.path.join(script_dir, "alunos")
if not os.path.exists(pasta_alunos):
    os.mkdir(pasta_alunos)

# Abre o certificado
img = Image.open(caminho_certificado)
fonte = ImageFont.truetype("Montserrat-SemiBold.ttf", 100)
draw = ImageDraw.Draw(img)
x, y = 290, 1145  # coordenadas da posição inicial para escrever o nome

# Inicia contador
contador = 0

# Itera pelos nomes
for nome in nomes:
    # Converte o nome para string e remove espaços extras
    nome = str(nome).strip()

    # Ignora linhas vazias ou com 'nan'
    if nome.lower() in ["nan", "none", ""]:
        continue
    # Cria o caminho completo para o arquivo
    caminho_arquivo = os.path.join(pasta_alunos, f"{nome}.png")
    
    # Cria uma cópia da imagem original
    img_nova = img.copy()
    
    # Cria um objeto para desenhar na imagem
    draw = ImageDraw.Draw(img_nova)
    
    # Adiciona o nome à imagem
    draw.text((x, y), nome, font=fonte, fill=(0, 0, 0))
    
    # Salva a imagem no caminho especificado
    img_nova.save(caminho_arquivo)
    
    # V2 - Contador de certificados gerados
    contador += 1 # incrementar o contador
    print(f"{contador} → {nome}")
    sleep(0.3)


# Exibe resumo final
print("="*50)
print(f"✨ Total de certificados gerados: {contador} ✨")
print("="*50)

print("Finalizado.")
