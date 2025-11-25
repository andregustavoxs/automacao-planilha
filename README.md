# Sistema de Geração de Lista de Convocação - TCE-MA

Automação desenvolvida para o SUDEC (Setor de Gestão de Estagiários) do Tribunal de Contas do Estado do Maranhão para gerar listas de convocação de estagiários seguindo as regras de cotas estabelecidas no edital.

## Descrição

Este sistema processa a planilha de classificação definitiva dos candidatos aprovados no processo seletivo e gera automaticamente a lista de convocação para **ENSINO SUPERIOR** e **NÍVEL TÉCNICO**, respeitando as cotas para:
- PCD (Pessoas com Deficiência)
- NEGRO (Candidatos negros)
- GERAL (Ampla concorrência)

## Regras de Classificação (Ensino Superior e Nível Técnico)

A lista de convocação segue as seguintes regras para ambos os níveis:

1. **Posições terminadas em 1** (1, 11, 21, 31...): reservadas para candidatos **PCD**
   - Se não houver candidatos PCD disponíveis para o curso, convoca-se da categoria GERAL

2. **Posições terminadas em 3, 6 ou 9** (3, 6, 9, 13, 16, 19...): reservadas para candidatos **NEGRO**
   - Se não houver candidatos NEGRO disponíveis para o curso, convoca-se da categoria GERAL

3. **Demais posições**: reservadas para candidatos **GERAL**

4. **Evita duplicações**: candidatos já convocados não são chamados novamente

5. **Classificação por curso**: a numeração sequencial reinicia para cada curso

## Estrutura do Projeto

```
automacao-planilha/
├── planilha_original/
│   └── CLASSIFICAÇÃO DEFINITIVA - TCE MA 01.2025.xlsx
├── planilha_referencia/
│   └── lista_convocacao_oficial.xlsx
├── resultado/
│   └── lista_convocacao_gerada.xlsx
├── gerar_lista_convocacao.py
├── requirements.txt
└── README.md
```

## Instalação

### 1. Pré-requisitos

- Python 3.8 ou superior
- Ambiente virtual Python (.venv)

### 2. Instalar dependências

```bash
# Ativar o ambiente virtual
.venv\Scripts\activate

# Instalar as bibliotecas necessárias
pip install -r requirements.txt
```

## Como Usar

### Opção 1: Painel Interativo (Recomendado)

1. **Execute o arquivo batch**:
   ```bash
   executar_app.bat
   ```

2. **Ou execute manualmente**:
   ```bash
   # Ativar o ambiente virtual
   .venv\Scripts\activate

   # Iniciar o painel Streamlit
   streamlit run app.py
   ```

3. **Use a interface web**:
   - Faça upload da planilha de classificação
   - Configure as abas se necessário
   - Clique em "Processar"
   - Baixe o resultado

### Opção 2: Script em Linha de Comando

1. **Preparar os arquivos**: Certifique-se de que a planilha de classificação definitiva está na pasta `planilha_original/` com o nome correto.

2. **Executar o script**:
   ```bash
   # No terminal, dentro do diretório do projeto
   python gerar_lista_convocacao.py
   ```

3. **Resultado**: O arquivo gerado estará em: `resultado/lista_convocacao_gerada.xlsx`

### Estrutura da planilha gerada

A planilha de saída contém as seguintes colunas:

| Coluna | Descrição |
|--------|-----------|
| CLASSIFICAÇÃO | Número sequencial da convocação |
| TIPO | PCD, NEGRO ou GERAL |
| CANDIDATO | Nome completo do candidato |
| CURSO | Curso do candidato |
| NÍVEL | ENSINO SUPERIOR ou NÍVEL TÉCNICO |

## Exemplo de Saída

A planilha gerada contém candidatos de ambos os níveis:

```
CLASSIFICAÇÃO | TIPO  | CANDIDATO                          | CURSO                        | NÍVEL
1             | PCD   | SUZIENY SOUZA DOS SANTOS           | ADMINISTRAÇÃO                | ENSINO SUPERIOR
2             | GERAL | LAURA NECY MARIA FROZ DE OLIVEIRA  | ADMINISTRAÇÃO                | ENSINO SUPERIOR
3             | NEGRO | NAYANE REIS SANTOS                 | ADMINISTRAÇÃO                | ENSINO SUPERIOR
4             | GERAL | THALISON COSTA RAMOS               | ADMINISTRAÇÃO                | ENSINO SUPERIOR
5             | GERAL | PEDRO PAULO DO NASCIMENTO SARAIVA  | ADMINISTRAÇÃO                | ENSINO SUPERIOR
6             | NEGRO | VANESSA CAMPOS ARAÚJO              | ADMINISTRAÇÃO                | ENSINO SUPERIOR
...
[após todos os cursos de ENSINO SUPERIOR]
...
1             | PCD   | ARTHUR DOS SANTOS DIAS             | TÉCNICO EM ADMINISTRAÇÃO     | NÍVEL TÉCNICO
2             | GERAL | YASMIN LIMA SOUSA FROZ             | TÉCNICO EM ADMINISTRAÇÃO     | NÍVEL TÉCNICO
3             | NEGRO | ARTHUR JOAQUIM REIS FERREIRA       | TÉCNICO EM ADMINISTRAÇÃO     | NÍVEL TÉCNICO
...
```

## Estatísticas

Ao final da execução, o sistema exibe estatísticas completas da lista gerada:
- Total de convocações
- Quantidade por nível (ENSINO SUPERIOR, NÍVEL TÉCNICO)
- Quantidade por tipo (PCD, NEGRO, GERAL)
- Quantidade por nível e tipo combinados
- Quantidade por curso (top 10)

## Funcionalidades

- ✅ Processamento automático seguindo regras de cotas
- ✅ Interface web interativa com Streamlit
- ✅ Upload de arquivos via navegador
- ✅ Visualização de estatísticas em tempo real
- ✅ Download direto do resultado processado
- ✅ Suporte para múltiplos níveis (Superior e Técnico)

## Próximas Versões

- Integração com o SGE (Sistema de Gestão de Estagiários)
- Exportação em múltiplos formatos (PDF, CSV)
- Validação automática de dados

## Autor

André Santos - TCE-MA/SUDEC

## Licença

Uso interno TCE-MA
