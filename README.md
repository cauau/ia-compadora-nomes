# 🔍 Comparador de Nomes com TF-IDF

Script em Python para comparar nomes entre duas colunas de uma planilha Excel usando técnicas de NLP (TF-IDF + similaridade de cosseno).

---

## 🚀 Funcionalidades

* Comparação automática entre nomes
* Identificação do melhor match
* Cálculo de score de similaridade (%)
* Classificação:

  * ✅ CONFERE
  * ❌ DIFERENTE
* Threshold automático baseado em meta
* Normalização de texto (acentos, símbolos, etc)

---

## 📂 Estrutura da planilha

### Entrada

| Coluna | Descrição            |
| ------ | -------------------- |
| B      | Nome base            |
| C      | Nome para comparação |

### Saída (gerada automaticamente)

| Coluna | Conteúdo       |
| ------ | -------------- |
| D      | RESULTADO      |
| E      | SCORE (%)      |
| F      | MELHOR_NOME_C  |
| G      | LINHA_MELHOR_C |

---

## ⚙️ Como usar

### 1. Clonar o repositório

```bash
git clone [https://github.com/cauau/ia-compadora-nomes]
cd SEU-REPO
```

### 2. Instalar dependências

```bash
pip install -r requirements.txt
```

### 3. Executar o script

```bash
python main.py sua_planilha.xlsx
```

---

## 🧠 Como funciona

O projeto utiliza técnicas de NLP para comparar textos:

* **TF-IDF** → transforma nomes em vetores numéricos
* **Cosine Similarity** → mede o grau de similaridade
* **N-grams de caracteres (2 a 4)** → melhora a comparação mesmo com erros de digitação

Também inclui:

* Remoção de acentos
* Padronização de texto
* Remoção de palavras irrelevantes (DA, DE, DO, etc)

---

## 📊 Exemplo

| Nome (B)    | Nome (C)      | Resultado |
| ----------- | ------------- | --------- |
| João Silva  | JOAO DA SILVA | CONFERE   |
| Maria Souza | MARIA S SOUZA | CONFERE   |
| Pedro Lima  | CARLOS LIMA   | DIFERENTE |

---

## 🛠 Tecnologias

* Python
* scikit-learn
* openpyxl

---

## 📄 Licença

MIT
