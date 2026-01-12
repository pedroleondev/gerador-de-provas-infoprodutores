import pandas as pd
from time import sleep
import os
from PIL import Image, ImageFont, ImageDraw

# Carregar os dados do Excel
df = pd.read_excel("arquivos\\2025_PROVA.xlsx", sheet_name="APROVADAS")
nomes = df["NOME"]

# Obtém o diretório do script atual
script_dir = os.path.dirname(os.path.abspath(__file__))
certificadoNome = "certificado\\certificado.png"
caminhoCertificado = os.path.join(script_dir, certificadoNome)

# Caminho para a pasta "aprovados"
pasta_aprovados = os.path.join(script_dir, "aprovados")
os.makedirs(pasta_aprovados, exist_ok=True)

# Carregar a imagem base uma única vez
img_original = Image.open(caminhoCertificado).convert("RGB")

# Carregar a fonte uma única vez
fonte = ImageFont.truetype("Montserrat-SemiBold.ttf", 85)

# Posição do nome
x, y = 310, 1155

for nome in nomes:
    # Criar uma cópia da imagem base
    img_nova = img_original.copy()
    draw = ImageDraw.Draw(img_nova)
    
    # Garantir que o nome não contenha caracteres inválidos
    nome_limpo = "".join(c if c.isalnum() or c in " _-" else "_" for c in nome)

    # Adicionar texto
    draw.text((x, y), nome, font=fonte, fill=(0, 0, 0))

    # Salvar a imagem na pasta "aprovados"
    caminho_salvar = os.path.join(pasta_aprovados, f"{nome_limpo}.png")
    img_nova.save(caminho_salvar)

    # Fechar a imagem para liberar memória
    img_nova.close()
    
    print(f"Certificado gerado para: {nome}")

    sleep(0.5)

# Fechar a imagem original para liberar memória
img_original.close()

print("Finalizado.")
