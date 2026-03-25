"""
Funções utilitárias para manipulação de arquivos.
"""

import pandas as pd
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


def carregar_excel(caminho: Path, **kwargs) -> pd.DataFrame:
    """Carrega arquivo Excel com tratamento de erro."""
    try:
        if not caminho.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

        df = pd.read_excel(caminho, **kwargs)
        logger.info(f"Arquivo carregado com sucesso: {caminho.name} ({len(df)} linhas)")
        return df
    except Exception as e:
        logger.error(f"Erro ao carregar {caminho.name}: {e}")
        raise


def salvar_excel(df: pd.DataFrame, caminho: Path, **kwargs):
    """Salva DataFrame em Excel com tratamento de erro."""
    try:
        # Garantir que o diretório existe
        caminho.parent.mkdir(parents=True, exist_ok=True)

        df.to_excel(caminho, index=False, **kwargs)
        logger.info(f"Arquivo salvo com sucesso: {caminho.name}")
    except Exception as e:
        logger.error(f"Erro ao salvar {caminho.name}: {e}")
        raise