#!/usr/bin/env python3
"""
Script para criar executável do PlanYazo
Autor: Hacktown 2025
"""

import os
import subprocess
import sys
import shutil

def main():
    print("🚀 Criando executável do PlanYazo...")
    
    # Verificar se PyInstaller está instalado
    try:
        import PyInstaller
        print("✅ PyInstaller encontrado")
    except ImportError:
        print("❌ PyInstaller não encontrado. Instalando...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # Usar ht.ico como ícone, se existir
    icon_file = "ht.ico"
    if os.path.exists(icon_file):
        cmd_icon = f"--icon={icon_file}"
    else:
        cmd_icon = None

    # Configurações do PyInstaller
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                    # Criar um único arquivo executável
        "--windowed",                   # Não mostrar console (aplicação GUI)
        "--name=PlanYazo",              # Nome do executável
        "--add-data=Plan.xlsx;.",       # Incluir Plan.xlsx
        "--add-data=Yazo.xlsx;.",       # Incluir Yazo.xlsx
        "--hidden-import=pandas",       # Incluir pandas explicitamente
        "--hidden-import=openpyxl",     # Incluir openpyxl explicitamente
        "--hidden-import=tkinter",      # Incluir tkinter explicitamente
        "--clean",                      # Limpar cache
        "Front.py"                      # Arquivo principal
    ]
    if cmd_icon:
        cmd.insert(6, cmd_icon)  # inserir logo após --name
    
    # Remover arquivos se não existirem
    if not os.path.exists("Plan.xlsx"):
        cmd = [arg for arg in cmd if not arg.startswith("--add-data=Plan.xlsx")]
        print("⚠️  Plan.xlsx não encontrado - não será incluído no executável")
    
    if not os.path.exists("Yazo.xlsx"):
        cmd = [arg for arg in cmd if not arg.startswith("--add-data=Yazo.xlsx")]
        print("⚠️  Yazo.xlsx não encontrado - não será incluído no executável")
    
    print("📦 Executando PyInstaller...")
    print(f"Comando: {' '.join(cmd)}")
    
    # Executar PyInstaller
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode == 0:
        print("✅ Executável criado com sucesso!")
        
        # Mover executável para pasta raiz
        exe_path = os.path.join("dist", "PlanYazo.exe")
        if os.path.exists(exe_path):
            # Criar pasta para o executável final
            final_dir = "PlanYazo_Executavel"
            if not os.path.exists(final_dir):
                os.makedirs(final_dir)
            
            # Copiar executável
            final_exe = os.path.join(final_dir, "PlanYazo.exe")
            shutil.copy2(exe_path, final_exe)
            
            # Copiar arquivos Excel se existirem
            if os.path.exists("Plan.xlsx"):
                shutil.copy2("Plan.xlsx", os.path.join(final_dir, "Plan.xlsx"))
            
            if os.path.exists("Yazo.xlsx"):
                shutil.copy2("Yazo.xlsx", os.path.join(final_dir, "Yazo.xlsx"))
            
            # Criar README para o executável
            create_readme(final_dir)
            
            print(f"📁 Executável disponível em: {final_dir}/")
            print(f"📄 Arquivos incluídos:")
            print(f"   - PlanYazo.exe")
            if os.path.exists("Plan.xlsx"):
                print(f"   - Plan.xlsx")
            if os.path.exists("Yazo.xlsx"):
                print(f"   - Yazo.xlsx")
            print(f"   - README_Executavel.txt")
            
        else:
            print("❌ Executável não foi criado corretamente")
            
    else:
        print("❌ Erro ao criar executável:")
        print(result.stderr)
    
    # Limpar arquivos temporários
    print("🧹 Limpando arquivos temporários...")
    if os.path.exists("build"):
        shutil.rmtree("build")
    if os.path.exists("dist"):
        shutil.rmtree("dist")
    if os.path.exists("PlanYazo.spec"):
        os.remove("PlanYazo.spec")

def create_readme(final_dir):
    """Criar README para o executável"""
    readme_content = """PlanYazo - Executável

INSTRUÇÕES DE USO:

1. Certifique-se que os arquivos Plan.xlsx e Yazo.xlsx estão na mesma pasta do executável
2. Dê duplo clique em PlanYazo.exe para executar
3. A interface gráfica será aberta automaticamente

FUNCIONALIDADES:
- 🎤 Palestra: Transfere dados da aba PALESTRA
- 🔧 Workshop: Transfere dados da aba WORKSHOP  
- 👥 Painel: Transfere dados da aba PAINEL
- 🗑️ Apagar Dados: Remove todos os registros

IMPORTANTE:
- Feche os arquivos Excel antes de executar
- Apenas linhas com fundo verde na coluna A serão processadas
- Confirme as operações quando solicitado

Desenvolvido para Hacktown 2025
"""
    
    readme_path = os.path.join(final_dir, "README_Executavel.txt")
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(readme_content)

if __name__ == "__main__":
    main() 