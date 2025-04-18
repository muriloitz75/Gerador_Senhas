import PyInstaller.__main__
import os

def find_resources():
    """Procura pelos arquivos de recursos na pasta do projeto."""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    resources = []
    
    # Procura pelo ícone
    icon_path = os.path.join(current_dir, 'ico1.ico')
    if os.path.exists(icon_path):
        resources.append(('ico1.ico', '.'))
    
    # Procura pela imagem
    img_path = os.path.join(current_dir, 'img.png')
    if os.path.exists(img_path):
        resources.append(('img.png', '.'))
    
    return resources

# Obter o caminho absoluto do diretório atual
current_dir = os.path.dirname(os.path.abspath(__file__))

# Encontrar o ícone
icon_path = os.path.join(current_dir, 'ico1.ico')

# Preparar os recursos
resources = find_resources()
resource_args = []
for src, dst in resources:
    resource_args.extend(['--add-data', f'{src}{os.pathsep}{dst}'])

# Preparar os argumentos do PyInstaller
pyinstaller_args = [
    'password_generator.py',
    '--onefile',
    '--windowed',
    '--name', 'Gerador de Senhas',
    '--clean',
    '--noconfirm'
]

# Adicionar o ícone se existir
if os.path.exists(icon_path):
    pyinstaller_args.extend(['--icon', icon_path])

# Adicionar os recursos
pyinstaller_args.extend(resource_args)

# Executar o PyInstaller
PyInstaller.__main__.run(pyinstaller_args)

