# Painel de Produtividade — Ativos e MP

Painel executivo para acompanhar a produtividade da equipe de **Ativos e Reagendamentos**:
quantas e **quais** ordens cada pessoa **cria**, **reagenda** e **cancela** — por mês e por dia.
Mesma identidade visual dos painéis anteriores (fonte Inter, paleta navy/azul, cards de KPI,
logo na barra lateral, gráficos com fundo transparente).

Abre já no **mês atual** e é pensado para uso direto pela diretoria.

---

## Arquivos

| Arquivo | O que é |
|---|---|
| `painel_produtividade.py` | O painel (Streamlit). |
| `logica.py` | Regras de negócio (o que conta como criar/reagendar/cancelar e a atribuição por pessoa). **É aqui que se ajusta a lógica.** |
| `extrair_produtividade.py` | Extrai as WOLI do Salesforce e gera `dados_produtividade.xlsx`. |
| `dados_produtividade.xlsx` | Base que o painel lê (vem com **dados de demonstração** até você rodar a extração). |
| `logo.png` | **Coloque aqui a sua logo** (a barra lateral usa este nome). |
| `INICIAR_PAINEL.bat` | Abre o painel (Windows, duplo-clique). |
| `ATUALIZAR_DADOS.bat` | Roda a extração do Salesforce (Windows). |
| `requirements.txt` | Dependências. |

---

## Como usar (Windows, mais simples)

1. Coloque sua **`logo.png`** nesta pasta.
2. **Ver na hora (demo):** duplo-clique em **`INICIAR_PAINEL.bat`**. O painel abre no navegador
   com dados de demonstração (aparece um aviso "Dados de demonstração").
3. **Dados reais:** configure as credenciais (abaixo) e duplo-clique em **`ATUALIZAR_DADOS.bat`**.
   Isso gera o `dados_produtividade.xlsx` de verdade e remove o aviso de demo.
4. Volte ao painel e recarregue a página (F5).

No dia a dia: rode `ATUALIZAR_DADOS.bat` quando quiser números atualizados e recarregue o painel.

## Como usar (linha de comando)

```bash
pip install -r requirements.txt
python extrair_produtividade.py       # gera dados_produtividade.xlsx
streamlit run painel_produtividade.py # abre o painel
```

---

## Segurança das credenciais (importante)

O extrator lê as credenciais de **variáveis de ambiente** primeiro:

```
# Windows (PowerShell)
setx SF_USERNAME "ext-...@culligan.com"
setx SF_PASSWORD "sua_senha"
setx SF_TOKEN    "seu_security_token"
```

Se preferir, dá para preencher direto as constantes no topo do `extrair_produtividade.py`,
mas evite versionar/enviar o arquivo com a senha dentro. Recomendo **trocar a senha e
regenerar o security token** que já foram usados em texto puro em scripts anteriores.

---

## ⚠️ Pontos para validar com a equipe (a lógica que vai para a diretoria)

Tudo isso está no topo do `logica.py`, fácil de ajustar. Vale confirmar com quem opera:

1. **Quem é a "pessoa" de cada ação.** Hoje o painel atribui:
   - *Criação* → quem **abriu o caso** (`Case_CreatedBy`);
   - *Reagendamento* → quem **criou a nova WOLI** (`WOLI_CreatedBy`);
   - *Cancelamento* → quem fez a **última alteração** da WOLI cancelada (`WOLI_LastModifiedBy`).

   Se, no seu Salesforce, esses registros forem criados/alterados por um **usuário de
   integração/automação** (e não pela pessoa), me avise: trocamos por outro campo
   (ex.: um "dono" ou "agente" da WOLI). É só mudar o nome do campo em `logica.py`.

2. **Reagendamento = nova WOLI.** O painel entende que cada nova WOLI de uma ordem já
   existente é um reagendamento (a WOLI anterior vira `Reagendado`). Se o seu fluxo
   remarca **sem criar nova WOLI**, a contagem muda — me diga que ajusto.

3. **Cancelamento pela pessoa.** Se você notar cancelamentos atribuídos a um usuário de
   sistema, dá para trocar a fonte para o **histórico de status** da WOLI (quem mudou o
   status para `Cancelado`). Deixei isso mapeado; é uma evolução rápida.

4. **Instalação (Vendas) × Diversos.** "Reagendamento de Vendas" é detectado quando o
   **Tipo de Serviço** da OS contém "Instalação". Se o nome for outro, ajusto o critério.

5. **Nomes das pessoas.** A equipe está cadastrada em `logica.py` (dicionário `EQUIPE`).
   O casamento com o Salesforce ignora maiúsculas/acentos, mas se algum nome de usuário
   estiver diferente lá, é só corrigir no dicionário.

---

## O que o painel mostra

- **Visão Geral:** KPIs (criadas, reagendadas, canceladas, total, média diária),
  evolução diária no mês, ações por operação, ranking por pessoa e o corte
  Instalação (Vendas) × Diversos, além do quadro-resumo por pessoa.
- **Por Pessoa:** o detalhe de uma pessoa, com opção de **filtrar um dia específico**
  (abre no dia de hoje) e a lista de **quais ordens** ela tocou.
- **Ordens (detalhe):** tabela completa e filtrável de todas as ações, com **exportação
  para Excel**.
