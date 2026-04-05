"""
storage.py — Etapa 4: Gestão de pastas de saída

Regras:
  - cada template tem pasta própria
  - se a pasta já existir, apaga o conteúdo (sem timestamp)
  - garante isolamento entre templates
"""

import os
import shutil

# Pasta raiz de saída (pode ser alterada)
OUTPUT_ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")


def prepare_output_folder(template_code: str, root: str = None) -> str:
    """
    Cria (ou limpa) a pasta de saída para o template.

    Se a pasta existir, apaga todo o conteúdo.
    Se não existir, cria.

    Retorna: caminho absoluto da pasta criada.
    """
    if root is None:
        root = OUTPUT_ROOT

    folder = os.path.join(root, template_code)

    if os.path.exists(folder):
        shutil.rmtree(folder)

    os.makedirs(folder, exist_ok=True)

    return os.path.abspath(folder)


def list_output_folders(root: str = None) -> list:
    """Lista todas as pastas de saída existentes."""
    if root is None:
        root = OUTPUT_ROOT

    if not os.path.exists(root):
        return []

    return sorted([
        d for d in os.listdir(root)
        if os.path.isdir(os.path.join(root, d))
    ])
