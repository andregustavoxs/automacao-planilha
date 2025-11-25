"""
Painel Interativo - Sistema de Geração de Lista de Convocação TCE-MA
Autor: André Santos
Descrição: Interface web com Streamlit para automatização da lista de convocação
"""

import streamlit as st
import pandas as pd
from pathlib import Path
import io
import time
import gc
from gerar_lista_convocacao import GeradorListaConvocacao
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet


def gerar_pdf(df):
    """Gera PDF a partir do DataFrame"""
    buffer = io.BytesIO()

    # Criar documento PDF em paisagem
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
    elements = []

    # Título
    styles = getSampleStyleSheet()
    title = Paragraph("<b>Lista de Convocação - TCE MA</b>", styles['Title'])
    elements.append(title)
    elements.append(Paragraph("<br/>", styles['Normal']))

    # Preparar dados da tabela
    data = [df.columns.tolist()] + df.values.tolist()

    # Criar tabela
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    elements.append(table)
    doc.build(elements)

    buffer.seek(0)
    return buffer


def main():
    # Configuração da página
    st.set_page_config(
        page_title="Gerador de Lista de Convocação - SUDEC",
        page_icon="📋",
        layout="centered"
    )

    # Cabeçalho minimalista
    st.title("Gerador de Lista de Convocação")


    # Expander com informações sobre o sistema
    with st.expander("ℹ️ Sobre o Sistema"):
        st.write("""
        Este sistema automatiza a geração de listas de convocação para a **SUDEC** (Supervisão de Desenvolvimento e Carreira).

        **Funcionalidade:**
        - Processa planilhas de classificação definitiva do seletivo de estagiários.
        - Gera lista ordenada considerando ampla concorrência e cotas (negros e PCD)
        - Produz planilha formatada para inserção no SGE (Sistema de Gestão de Estagiários)
        """)

    st.markdown("")
    st.markdown("")

    # Upload do arquivo
    uploaded_file = st.file_uploader(
        "Envie a planilha de classificação",
        type=['xlsx'],
        label_visibility="collapsed"
    )

    if uploaded_file is not None:
        # Salvar arquivo temporariamente
        temp_path = Path("temp_upload.xlsx")
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Configuração automática das abas
        aba_superior_ampla = 'SUPERIOR - AMPLA'
        aba_superior_negros = 'SUPERIOR - NEGROS'
        aba_superior_pcd = 'SUPERIOR - PCD'
        aba_tecnico_ampla = 'TECNICO - AMPLA'
        aba_tecnico_negros = ' TECNICO - NEGROS'
        aba_tecnico_pcd = 'TECNICO - PCD'

        # Processar automaticamente
        if 'df_resultado' not in st.session_state or st.session_state.get('last_file') != uploaded_file.name:
            # Placeholders para feedback
            status_container = st.empty()
            progress_container = st.empty()

            try:
                # Barra de progresso
                progress_bar = progress_container.progress(0)

                # Processar Ensino Superior
                status_container.info("Processando Ensino Superior...")
                progress_bar.progress(10)

                gerador_superior = GeradorListaConvocacao(
                    arquivo_origem=str(temp_path),
                    aba_geral=aba_superior_ampla,
                    aba_negro=aba_superior_negros,
                    aba_pcd=aba_superior_pcd,
                    nivel='ENSINO SUPERIOR'
                )
                gerador_superior.carregar_planilhas()
                progress_bar.progress(30)

                df_superior = gerador_superior.gerar_lista_completa()
                progress_bar.progress(50)

                # Processar Nível Técnico
                status_container.info("Processando Nível Técnico...")
                gerador_tecnico = GeradorListaConvocacao(
                    arquivo_origem=str(temp_path),
                    aba_geral=aba_tecnico_ampla,
                    aba_negro=aba_tecnico_negros,
                    aba_pcd=aba_tecnico_pcd,
                    nivel='NÍVEL TÉCNICO'
                )
                gerador_tecnico.carregar_planilhas()
                progress_bar.progress(70)

                df_tecnico = gerador_tecnico.gerar_lista_completa()
                progress_bar.progress(85)

                # Combinar resultados
                status_container.info("Finalizando...")
                df_resultado_final = pd.concat([df_superior, df_tecnico], ignore_index=True)
                progress_bar.progress(100)

                # Armazenar no session_state
                st.session_state['df_resultado'] = df_resultado_final
                st.session_state['last_file'] = uploaded_file.name

                # Limpar arquivo temporário
                try:
                    gc.collect()
                    time.sleep(0.1)
                    if temp_path.exists():
                        temp_path.unlink()
                except Exception:
                    pass

                # Limpar containers de progresso
                status_container.empty()
                progress_container.empty()

                # Força reload para mostrar resultados
                st.rerun()

            except Exception as e:
                status_container.empty()
                progress_container.empty()
                st.error(f"Erro ao processar: {str(e)}")
                try:
                    gc.collect()
                    time.sleep(0.1)
                    if temp_path.exists():
                        temp_path.unlink()
                except Exception:
                    pass
                return

    # Exibir resultados se já processado
    if 'df_resultado' in st.session_state:
        df_resultado = st.session_state['df_resultado']

        st.markdown("")
        st.markdown("")

        # Mostrar prévia dos dados
        st.subheader("Prévia dos Dados Processados")
        st.info(f"Total de registros: {len(df_resultado)}")

        # Exibir dataframe com scroll
        st.dataframe(
            df_resultado,
            use_container_width=True,
            height=400
        )

        st.markdown("")
        st.markdown("")

        # Opções de download
        col1, col2, col3 = st.columns(3)

        # Converter para Excel em memória
        output_xlsx = io.BytesIO()
        with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer:
            df_resultado.to_excel(writer, index=False, sheet_name='Lista de Convocacao')
        output_xlsx.seek(0)

        # Converter para CSV em memória
        output_csv = io.StringIO()
        df_resultado.to_csv(output_csv, index=False, encoding='utf-8-sig')
        output_csv_bytes = output_csv.getvalue().encode('utf-8-sig')

        # Gerar PDF
        output_pdf = gerar_pdf(df_resultado)

        # Botão de download Excel
        with col1:
            st.download_button(
                label="📥 Baixar Excel",
                data=output_xlsx,
                file_name="lista_convocacao_gerada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        # Botão de download CSV
        with col2:
            st.download_button(
                label="📥 Baixar CSV",
                data=output_csv_bytes,
                file_name="lista_convocacao_gerada.csv",
                mime="text/csv",
                type="secondary",
                use_container_width=True
            )

        # Botão de download PDF
        with col3:
            st.download_button(
                label="📥 Baixar PDF",
                data=output_pdf,
                file_name="lista_convocacao_gerada.pdf",
                mime="application/pdf",
                type="secondary",
                use_container_width=True
            )


if __name__ == '__main__':
    main()
