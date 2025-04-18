from PIL import Image
import os

def create_icon():
    # Criar uma imagem 256x256 com fundo transparente
    img = Image.new('RGBA', (256, 256), (255, 255, 255, 0))
    
    # Desenhar um cadeado simples
    from PIL import ImageDraw
    
    draw = ImageDraw.Draw(img)
    
    # Corpo do cadeado (retângulo arredondado)
    draw.rectangle([64, 128, 192, 256], fill=(70, 70, 70, 255))
    
    # Arco do cadeado
    draw.arc([48, 32, 208, 160], 0, 180, fill=(70, 70, 70, 255), width=32)
    
    # Salvar nos tamanhos necessários para o arquivo .ico
    sizes = [(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)]
    
    icon_sizes = []
    for size in sizes:
        resized_img = img.resize(size, Image.Resampling.LANCZOS)
        icon_sizes.append(resized_img)
    
    # Salvar como .ico
    img.save('icon.ico', format='ICO', sizes=[(x.width, x.height) for x in icon_sizes])
    
    # Salvar também como .png para o ícone da janela
    img.save('icon.png', format='PNG')

if __name__ == '__main__':
    create_icon()