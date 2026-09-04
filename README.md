# Concilia+ — App Multi-Cliente (Streamlit)

Um único app, um único link. Cada lojista entra com **usuário + senha**; o
app carrega automaticamente a configuração daquele cliente (nome, planilhas,
colunas, estratégia de matching) e nunca mostra nada de outro lojista na
mesma sessão.

## Estrutura

```
concilia-multicliente/
├── app.py                        # login + roteamento + UI (Streamlit)
├── clients_config.py             # config de cada lojista — cresce a cada cliente novo
├── reconciliation.py             # motor de conciliação genérico
├── report_builder.py             # gerador do Excel de saída (2 abas)
├── requirements.txt
├── .streamlit/
│   └── secrets.toml.example      # modelo do que vai no painel do Streamlit Cloud
├── .gitignore
└── README.md
```

**Separação importante:** as *senhas* de cada cliente **não ficam no
GitHub** — ficam nos Secrets do próprio Streamlit Cloud. Suspender/reativar
cliente nunca precisa de git.

---

## Passo a passo para colocar no ar (do zero)

### 1. Criar o repositório no GitHub

```bash
cd concilia-multicliente
git init
git add .
git commit -m "Concilia+ multi-cliente v1"
git branch -M main
git remote add origin https://github.com/SEU-USUARIO/concilia-multicliente.git
git push -u origin main
```

(Repositório pode ser público — não tem senha nem dado de cliente nenhum
aqui, só nomes de colunas e a lógica de conciliação.)

### 2. Conectar no Streamlit Community Cloud

1. Acesse [share.streamlit.io](https://share.streamlit.io) e faça login com
   sua conta do GitHub
2. **New app** → selecione o repositório `concilia-multicliente`, branch
   `main`, arquivo principal `app.py`
3. Clique em **Deploy** — a build leva 1-2 minutos na primeira vez

### 3. Configurar os Secrets (senhas)

1. No painel do app recém-criado, vá em **⋮ (menu) → Settings → Secrets**
2. Cole o conteúdo abaixo (copiado e adaptado de
   `.streamlit/secrets.toml.example`):

   ```toml
   [credentials]
   hoh-clinica-integrada = "escolha-uma-senha-forte-aqui"
   ```

3. **Save** — o app reinicia sozinho com os secrets carregados

### 4. Testar com o cliente de exemplo

O cliente `hoh-clinica-integrada` já vem configurado e **testado com dados
reais** (é o mesmo Sistema.xlsx/Maquina.xlsx que validamos nesta conversa:
R$ 22.995,19 de entradas, R$ 619,41 em taxas, 1 divergência real, 6
pagamentos em lote resolvidos automaticamente).

1. Abra a URL do app (algo como `concilia-multicliente.streamlit.app`)
2. Usuário: `hoh-clinica-integrada` / Senha: a que você cadastrou no passo 3
3. Suba os dois arquivos de teste, clique em **Conciliar vendas**
4. Confira se os 3 cards batem com os números acima, e baixe o relatório

Se bateu, o pipeline está redondo de ponta a ponta — pode passar a
cadastrar clientes reais.

---

## Operação do dia a dia

### Cliente novo, formato de relatório já conhecido

Só cadastrar credencial — sem tocar em código:

1. **Settings → Secrets** → adicione uma linha, ex:
   `clinica-sorriso = "senha-dela"`
2. Em `clients_config.py`, adicione o bloco dela (se for igual a um formato
   que já existe, é copiar/colar um bloco existente e trocar `display_name`)
3. Manda pro cliente: link do app + usuário (`clinica-sorriso`) + senha

### Cliente novo, formato de relatório nunca visto

1. Peça ao cliente os dois relatórios (Sistema + Máquina) de um período já
   fechado
2. Me mande os arquivos aqui no Claude — eu identifico as colunas, confirmo
   se o matching é por `valor_data` ou `id_direto`, crio o bloco novo em
   `clients_config.py` e testo com os arquivos reais dele antes de te
   devolver
3. No GitHub, abra `clients_config.py` no navegador (ícone de lápis), cole o
   bloco novo dentro do dicionário `CLIENTS`, **Commit changes**
4. O Streamlit Cloud detecta o commit e atualiza o app sozinho em 1-2 minutos
5. Cadastre a senha dele em Secrets (igual ao caso anterior)

### Cliente atrasou ou cancelou

**Settings → Secrets** → apague a linha dele. A senha para de funcionar na
hora. (Alternativa: deixar a senha cadastrada mas mudar `"ativo": False` no
bloco dele em `clients_config.py` — nesse caso o app mostra "Acesso
suspenso" em vez de "usuário ou senha inválidos", mais claro se for algo
temporário.)

### Cliente voltou a pagar

Recadastre a senha em Secrets (ou volte `"ativo": True`).

### O formato de um cliente existente mudou

Me mande o relatório novo dele — ajusto o bloco dele em `clients_config.py`
e devolvo o trecho atualizado pra colar no GitHub.

---

## Resumo — quando você precisa de mim

| Situação | Precisa de mim? | Precisa mexer em algo técnico? |
|---|---|---|
| Cliente novo, formato conhecido | Não | Colar 2 coisas: bloco em `clients_config.py` (copiar de outro já existente) + senha em Secrets |
| Cliente novo, formato novo | Sim | Colar o bloco que eu te der no GitHub (navegador) + senha em Secrets |
| Suspender/reativar cliente | Não | Só o painel de Secrets (ou `"ativo"` no config) |
| Formato de cliente existente mudou | Sim | Colar o trecho atualizado no GitHub (navegador) |
| Uso normal, mês a mês | Não | Nada — o lojista usa sozinho |

## Sobre as duas estratégias de matching

- **`valor_data`** (padrão, mais comum): usa quando os IDs do Sistema e da
  Máquina não se correspondem — descoberta validada com dados reais, onde os
  identificadores vivem em faixas numéricas completamente diferentes.
  Inclui resolução automática de pagamentos agrupados em lote.
- **`id_direto`**: só use se você confirmar, nos arquivos reais do cliente,
  que o PDV e a adquirente exportam o mesmo identificador de transação —
  testado com dados sintéticos nesta entrega, mas ainda sem validação com um
  cliente real desse perfil. Quando aparecer o primeiro caso assim, vale
  conferir com atenção antes de liberar pro cliente.

## Limites a saber (tier gratuito do Streamlit Community Cloud)

- 1 GB de memória por app — confortável para o volume de linhas típico de
  reconciliação de pequeno/médio lojista
- O app "dorme" depois de ~12h sem nenhum acesso — a primeira pessoa a abrir
  depois disso espera alguns segundos a mais
- Sem domínio próprio no plano grátis (fica em `algumacoisa.streamlit.app`)
- Arquivos até 10 MB por upload (limite aplicado em `app.py`)
