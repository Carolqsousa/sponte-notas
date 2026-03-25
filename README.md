# 📊 Sponte — Mapa de Notas

App web para extração automática de notas dos alunos via API Sponte, com geração de planilhas Excel organizadas por professor.

---

## 🚀 O que o app faz

- Conecta à API do Sponte Web de cada unidade
- Extrai notas de todos os alunos das turmas abertas do semestre
- Calcula médias ponderadas por componente e por prova
- Gera planilhas Excel com uma aba por professor
- Destaca notas abaixo de 7 (vermelho) e notas não lançadas (amarelo)
- Permite atualizar uma planilha existente mantendo comentários já preenchidos
- Suporta download de todas as unidades de uma vez (ZIP)

---

## 📋 Estrutura das planilhas

Cada arquivo gerado tem **uma aba por professor**. Professores com turmas nos dois formatos de boletim ganham abas separadas (`Nome - A` e `Nome - B`).

### Formato A — Progress Check / Mid-term / Final
Usado por: Elementary, Intermediate, Pre-Intermediate, Teen, Upper Intermediate 1 e 2, Pre-Teen, Stars, Young, etc.

| Coluna | Descrição |
|--------|-----------|
| Turma | Nome da turma no Sponte |
| Professor | Nome do professor *(oculto por padrão)* |
| student_id | ID do aluno no Sponte *(oculto por padrão)* |
| Nome Aluno | Nome completo do aluno |
| Situação | Aprovado / Reprovado *(oculto por padrão)* |
| Média Geral | Média ponderada de todas as provas lançadas |
| Action Plan | Campo livre para plano de ação *(preenchido manualmente)* |
| Coordinator's Comment | Campo livre para comentário da coordenação *(preenchido manualmente)* |
| ⚠️ Alerta | Indica alunos com nota < 7 ou sem nota nas provas lançadas *(oculto, usado no filtro)* |
| PC - Speaking/Oral | Nota do Progress Check |
| PC - Gramática & Vocab | Nota do Progress Check |
| PC - Homework | Nota do Progress Check |
| Média PC | Média ponderada do Progress Check |
| Mid - Homework | Nota do Mid-term |
| Mid - Reading | Nota do Mid-term |
| ... | ... |
| Média Mid | Média ponderada do Mid-term |
| Final - ... | Notas do Final |
| Média Final | Média ponderada do Final |

### Formato B — Avaliação 1 / 2 / 3 / 4
Usado por: Advanced, Mastery, Vantage, Upper Intermediate 3.

Cada avaliação tem um único componente:
- **Av1** → Reading And Use of English
- **Av2** → Writing
- **Av3** → Listening
- **Av4** → Homework

---

## 🎨 Formatação visual

| Cor | Significado |
|-----|-------------|
| 🔴 Vermelho claro | Nota abaixo de 7 |
| 🟡 Amarelo claro | Nota não lançada |
| ⬜ Cinza | Aluno que não consta mais na turma |
| 🟦 Azul escuro (cabeçalho) | Cabeçalho das colunas |

**Filtro automático:** por padrão a planilha exibe apenas alunos com `⚠️` na coluna Alerta — ou seja, quem tem alguma nota abaixo de 7 ou não lançada nas provas já aplicadas. Para ver todos os alunos basta limpar o filtro.

**Seletor de prova:** no app, antes de gerar a planilha, o usuário informa qual foi a última prova lançada (Progress Check, Mid-term ou Final). Isso define quais colunas são consideradas no cálculo do alerta.

---

## 🔄 Atualizar planilha existente

O app permite fazer upload de uma planilha gerada anteriormente. Ao atualizar:

- Notas são atualizadas com os dados mais recentes da API
- **Action Plan** e **Coordinator's Comment** são preservados por aluno (cruzamento por `student_id`)
- Alunos que saíram da turma ficam na planilha com status `⚠️ Não consta mais na turma` em cinza

---

## ⚙️ Configuração

### Pré-requisitos
- Python 3.10+
- Conta no [Streamlit Cloud](https://share.streamlit.io) (gratuito)
- Tokens de API do Sponte para cada unidade

