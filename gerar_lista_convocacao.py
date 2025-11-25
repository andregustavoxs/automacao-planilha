"""
Sistema de Geração de Lista de Convocação - TCE-MA
Autor: André Santos
Descrição: Automatiza a geração da lista de convocação de estagiários
seguindo as regras de cotas PCD, NEGRO e GERAL para ENSINO SUPERIOR e NÍVEL TÉCNICO.
"""

import pandas as pd
from pathlib import Path
from typing import Dict, List, Tuple


class GeradorListaConvocacao:
    """Classe responsável por gerar a lista de convocação seguindo as regras do edital."""

    def __init__(self, arquivo_origem: str, aba_geral: str, aba_negro: str, aba_pcd: str, nivel: str):
        """
        Inicializa o gerador com o arquivo de origem.

        Args:
            arquivo_origem: Caminho para o arquivo Excel com as classificações
            aba_geral: Nome da aba GERAL (AMPLA)
            aba_negro: Nome da aba NEGRO
            aba_pcd: Nome da aba PCD
            nivel: Nível dos candidatos (ex: 'ENSINO SUPERIOR' ou 'NÍVEL TÉCNICO')
        """
        self.arquivo_origem = arquivo_origem
        self.aba_geral = aba_geral
        self.aba_negro = aba_negro
        self.aba_pcd = aba_pcd
        self.nivel = nivel
        self.df_geral = None
        self.df_negro = None
        self.df_pcd = None

    def carregar_planilhas(self):
        """Carrega as três planilhas de classificação (GERAL, NEGRO, PCD)."""
        print(f"\nCarregando planilhas de {self.nivel}...")

        # Ler planilha GERAL (AMPLA)
        self.df_geral = pd.read_excel(
            self.arquivo_origem,
            sheet_name=self.aba_geral,
            skiprows=3
        )

        # Ler planilha NEGRO
        self.df_negro = pd.read_excel(
            self.arquivo_origem,
            sheet_name=self.aba_negro,
            skiprows=3
        )

        # Ler planilha PCD
        self.df_pcd = pd.read_excel(
            self.arquivo_origem,
            sheet_name=self.aba_pcd,
            skiprows=3
        )

        # Limpar espaços nos nomes das colunas e nos valores de CURSO
        for df in [self.df_geral, self.df_negro, self.df_pcd]:
            df.columns = df.columns.str.strip()
            df['CURSO'] = df['CURSO'].str.strip()
            df['NOME'] = df['NOME'].str.strip()

        print(f"  [OK] GERAL: {len(self.df_geral)} candidatos")
        print(f"  [OK] NEGRO: {len(self.df_negro)} candidatos")
        print(f"  [OK] PCD: {len(self.df_pcd)} candidatos")

    def determinar_tipo_posicao(self, posicao: int) -> str:
        """
        Determina o tipo de candidato que deve ocupar determinada posição.

        Args:
            posicao: Número da posição (1, 2, 3, ...)

        Returns:
            'PCD', 'NEGRO' ou 'GERAL'
        """
        ultimo_digito = posicao % 10

        # Posições terminadas em 1: PCD (1, 11, 21, 31, ...)
        if ultimo_digito == 1:
            return 'PCD'

        # Posições terminadas em 3, 6, 9: NEGRO (3, 6, 9, 13, 16, 19, ...)
        if ultimo_digito in [3, 6, 9]:
            return 'NEGRO'

        # Demais posições: GERAL
        return 'GERAL'

    def gerar_lista_por_curso(self, curso: str) -> List[Tuple[int, str, str, str, str]]:
        """
        Gera a lista de convocação para um curso específico.

        Args:
            curso: Nome do curso

        Returns:
            Lista de tuplas (classificação, tipo, candidato, curso, nível)
        """
        # Obter candidatos de cada categoria para o curso
        candidatos_geral = self.df_geral[self.df_geral['CURSO'] == curso]['NOME'].tolist()
        candidatos_negro = self.df_negro[self.df_negro['CURSO'] == curso]['NOME'].tolist()
        candidatos_pcd = self.df_pcd[self.df_pcd['CURSO'] == curso]['NOME'].tolist()

        # Criar índices para controlar a posição em cada lista
        idx_geral = 0
        idx_negro = 0
        idx_pcd = 0

        # Conjunto para rastrear candidatos já convocados (evitar duplicações)
        convocados = set()

        # Lista resultado
        resultado = []

        # Calcular total de candidatos únicos para o curso
        todos_candidatos = set(candidatos_geral + candidatos_negro + candidatos_pcd)
        total_candidatos = len(todos_candidatos)

        posicao = 1

        while len(convocados) < total_candidatos:
            tipo_esperado = self.determinar_tipo_posicao(posicao)
            candidato_selecionado = None
            tipo_selecionado = None

            # Tentar preencher com o tipo esperado
            if tipo_esperado == 'PCD':
                # Procurar próximo PCD não convocado
                while idx_pcd < len(candidatos_pcd):
                    if candidatos_pcd[idx_pcd] not in convocados:
                        candidato_selecionado = candidatos_pcd[idx_pcd]
                        tipo_selecionado = 'PCD'
                        idx_pcd += 1
                        break
                    idx_pcd += 1

                # Se não houver PCD, pegar do GERAL
                if candidato_selecionado is None:
                    while idx_geral < len(candidatos_geral):
                        if candidatos_geral[idx_geral] not in convocados:
                            candidato_selecionado = candidatos_geral[idx_geral]
                            tipo_selecionado = 'GERAL'
                            idx_geral += 1
                            break
                        idx_geral += 1

            elif tipo_esperado == 'NEGRO':
                # Procurar próximo NEGRO não convocado
                while idx_negro < len(candidatos_negro):
                    if candidatos_negro[idx_negro] not in convocados:
                        candidato_selecionado = candidatos_negro[idx_negro]
                        tipo_selecionado = 'NEGRO'
                        idx_negro += 1
                        break
                    idx_negro += 1

                # Se não houver NEGRO, pegar do GERAL
                if candidato_selecionado is None:
                    while idx_geral < len(candidatos_geral):
                        if candidatos_geral[idx_geral] not in convocados:
                            candidato_selecionado = candidatos_geral[idx_geral]
                            tipo_selecionado = 'GERAL'
                            idx_geral += 1
                            break
                        idx_geral += 1

            else:  # GERAL
                # Procurar próximo GERAL não convocado
                while idx_geral < len(candidatos_geral):
                    if candidatos_geral[idx_geral] not in convocados:
                        candidato_selecionado = candidatos_geral[idx_geral]
                        tipo_selecionado = 'GERAL'
                        idx_geral += 1
                        break
                    idx_geral += 1

            # Se encontrou um candidato, adicionar à lista
            if candidato_selecionado:
                convocados.add(candidato_selecionado)
                resultado.append((
                    posicao,
                    tipo_selecionado,
                    candidato_selecionado,
                    curso,
                    self.nivel
                ))
                posicao += 1
            else:
                # Não há mais candidatos disponíveis
                break

        return resultado

    def gerar_lista_completa(self) -> pd.DataFrame:
        """
        Gera a lista de convocação completa para todos os cursos.

        Returns:
            DataFrame com a lista de convocação
        """
        print(f"\nGerando lista de convocacao para {self.nivel}...")

        # Obter lista de todos os cursos únicos
        cursos = sorted(self.df_geral['CURSO'].unique())

        # Lista para armazenar todos os resultados
        resultado_final = []

        # Processar cada curso
        for curso in cursos:
            print(f"  Processando: {curso}")
            lista_curso = self.gerar_lista_por_curso(curso)
            resultado_final.extend(lista_curso)

        # Criar DataFrame
        df_resultado = pd.DataFrame(
            resultado_final,
            columns=['CLASSIFICAÇÃO', 'TIPO', 'CANDIDATO', 'CURSO', 'NÍVEL']
        )

        print(f"  [OK] {self.nivel}: {len(df_resultado)} convocacoes")

        return df_resultado


def salvar_resultado(df: pd.DataFrame, arquivo_saida: str):
    """
    Salva o resultado em um arquivo Excel.

    Args:
        df: DataFrame com os resultados
        arquivo_saida: Caminho do arquivo de saída
    """
    print(f"\nSalvando resultado em: {arquivo_saida}")
    df.to_excel(arquivo_saida, index=False, sheet_name='Lista de Convocacao')
    print("[OK] Arquivo salvo com sucesso!")


def main():
    """Função principal do script."""
    # Configurar caminhos
    arquivo_origem = 'planilha_original/CLASSIFICAÇÃO DEFINITIVA - TCE MA 01.2025.xlsx'
    arquivo_saida = 'resultado/lista_convocacao_gerada.xlsx'

    # Criar diretório de resultado se não existir
    Path('resultado').mkdir(exist_ok=True)

    print("="*70)
    print(" SISTEMA DE GERACAO DE LISTA DE CONVOCACAO - TCE-MA/SUDEC")
    print("="*70)

    # ==================== PROCESSAR ENSINO SUPERIOR ====================
    gerador_superior = GeradorListaConvocacao(
        arquivo_origem=arquivo_origem,
        aba_geral='SUPERIOR - AMPLA',
        aba_negro='SUPERIOR - NEGROS',
        aba_pcd='SUPERIOR - PCD',
        nivel='ENSINO SUPERIOR'
    )
    gerador_superior.carregar_planilhas()
    df_superior = gerador_superior.gerar_lista_completa()

    # ==================== PROCESSAR NÍVEL TÉCNICO ====================
    gerador_tecnico = GeradorListaConvocacao(
        arquivo_origem=arquivo_origem,
        aba_geral='TECNICO - AMPLA',
        aba_negro=' TECNICO - NEGROS',
        aba_pcd='TECNICO - PCD',
        nivel='NÍVEL TÉCNICO'
    )
    gerador_tecnico.carregar_planilhas()
    df_tecnico = gerador_tecnico.gerar_lista_completa()

    # ==================== COMBINAR RESULTADOS ====================
    print("\n" + "="*70)
    print("Combinando resultados...")
    df_resultado_final = pd.concat([df_superior, df_tecnico], ignore_index=True)
    print(f"[OK] Total geral: {len(df_resultado_final)} convocacoes")

    # ==================== SALVAR RESULTADO ====================
    salvar_resultado(df_resultado_final, arquivo_saida)

    # ==================== EXIBIR ESTATÍSTICAS ====================
    print("\n" + "="*70)
    print("ESTATISTICAS DA LISTA GERADA")
    print("="*70)
    print(f"Total de convocacoes: {len(df_resultado_final)}")

    print(f"\nPor nivel:")
    print(df_resultado_final['NÍVEL'].value_counts())

    print(f"\nPor tipo:")
    print(df_resultado_final['TIPO'].value_counts())

    print(f"\nPor nivel e tipo:")
    print(df_resultado_final.groupby(['NÍVEL', 'TIPO']).size())

    print(f"\nPor curso (top 10):")
    print(df_resultado_final['CURSO'].value_counts().head(10))

    print("="*70)


if __name__ == '__main__':
    main()
