# excel_functions.py
"""
Arquivo com funções Python para o add-in do Excel
Este arquivo contém as funções que serão executadas pelo add-in
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os

def process_excel_data(data_range, output_path=None):
    """
    Processa dados do Excel e salva como TXT
    
    Args:
        data_range: Lista de listas com os dados do Excel
        output_path: Caminho para salvar o arquivo TXT (opcional)
    
    Returns:
        str: Caminho do arquivo salvo ou dados processados
    """
    try:
        # Converter dados para DataFrame
        df = pd.DataFrame(data_range)
        
        # Usar primeira linha como cabeçalho se apropriado
        if len(df) > 1:
            df.columns = df.iloc[0]
            df = df.drop(df.index[0]).reset_index(drop=True)
        
        # Processamento básico - remover valores nulos
        df = df.dropna()
        
        # Definir caminho de saída se não fornecido
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"excel_data_{timestamp}.txt"
        
        # Salvar como TXT
        df.to_csv(output_path, sep='\t', index=False)
        
        return f"Dados salvos em: {output_path}"
        
    except Exception as e:
        return f"Erro ao processar dados: {str(e)}"

def analyze_data(data_range):
    """
    Análise estatística básica dos dados
    """
    try:
        df = pd.DataFrame(data_range)
        
        # Tentar converter para numérico onde possível
        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='ignore')
        
        # Estatísticas descritivas
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        
        if len(numeric_cols) > 0:
            stats = df[numeric_cols].describe()
            return stats.to_dict()
        else:
            return "Nenhuma coluna numérica encontrada"
            
    except Exception as e:
        return f"Erro na análise: {str(e)}"

def filter_data(data_range, column_name, filter_value):
    """
    Filtra dados baseado em uma coluna específica
    """
    try:
        df = pd.DataFrame(data_range)
        
        if len(df) > 1:
            df.columns = df.iloc[0]
            df = df.drop(df.index[0]).reset_index(drop=True)
        
        # Aplicar filtro
        filtered_df = df[df[column_name] == filter_value]
        
        # Retornar como lista de listas
        result = [filtered_df.columns.tolist()] + filtered_df.values.tolist()
        return result
        
    except Exception as e:
        return f"Erro no filtro: {str(e)}"
