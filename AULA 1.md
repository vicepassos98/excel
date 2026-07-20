# Operações Matemáticas

| Símbolo | Operação | Exemplo | Uso |
|:-------:|----------|---------|-----|
| + | Adição | `=67+42` | Realizar operações entre dois ou mais números ou duas ou mais células. |
| - | Subtração | `=390-247` | Realizar operações entre dois ou mais números ou duas ou mais células. |
| * | Multiplicação | `=7*6` | Realizar operações entre dois ou mais números ou duas ou mais células. |
| / | Divisão | `=335/5` | Realizar operações entre dois ou mais números ou duas ou mais células. |
| ^ | Exponenciação | `=2^8` | Realizar operações entre dois ou mais números ou duas ou mais células. |

> **Importante:** Para iniciar qualquer operação no Excel, coloque o sinal `=` no início da fórmula.

# Funções

| Função | Fórmula | Exemplo | Uso |
|---------|---------|---------|-----|
| Soma | `=SOMA(Intervalo)` | `=SOMA(C3:C7)` | Somar vários números dentro de um ou mais intervalos. |

# [Planilha de Vendas](https://docs.google.com/spreadsheets/d/1Yj38W8jobApl3GWp2ng0ZbEMOxp1zOxa/export?format=xlsx)

|   | **A** | **B** | **C** | **D** | **E** | **F** | **G** | **H** | **I** |
|---|-------|----------------------|-----------|-----------|--------|-----------|---------------|------------|------------------|
| **1** | CÓDIGO | PRODUTO | PREÇO | CUSTO | VENDAS | MARGEM | FATURAMENTO | LUCRO | MARGEM DE LUCRO |
| **2** | 101 | Teclado Multilaser | R$ 89,00 | R$ 53,00 | 35 | `=C2-D2` | `=C2*E2` | `=F2*E2` | `=H2/G2` |
| **3** | 102 | Teclado Logitech | R$ 120,00 | R$ 102,00 | 102 |  |  |  |  |
| **4** | 103 | Teclado Mecânico | R$ 300,00 | R$ 215,00 | 67 |  |  |  |  |
| **5** | 104 | Mouse Multilaser | R$ 20,00 | R$ 12,00 | 193 |  |  |  |  |
| **6** | 105 | Mouse Sem Fio | R$ 55,00 | R$ 42,00 | 89 |  |  |  |  |
| **7** | 106 | Mouse Gamer Logi | R$ 150,00 | R$ 95,00 | 35 |  |  |  |  |
| **8** |  |  |  |  |  |  |  |  |  |
| **9** |  | **Resumo Vendas** | **Valores** |  |  |  |  |  |  |
| **10** |  | Vendas | `=E2+E3+E4+E5+E6+E7` |  |  |  |  |  |  |
| **11** |  | Faturamento | `=SOMA(G2:G7)` |  |  |  |  |  |  |
| **12** |  | Lucro | `=SOMA(H2:H7)` |  |  |  |  |  |  |
| **13** |  | Margem de Lucro | `=C12/C11` |  |  |  |  |  |  |

---

# Conceitos Financeiros

## Margem

A **margem** representa o ganho obtido em cada unidade vendida. Ela é calculada subtraindo o custo do preço de venda.

**Fórmula:**

```excel
=Preço - Custo
```

**Exemplo:**

```excel
=C2-D2
```

Neste exemplo:

- Preço: R$ 89,00
- Custo: R$ 53,00

Margem = **R$ 36,00**

---

## Faturamento

O **faturamento** é o valor total das vendas realizadas, sem descontar os custos.

**Fórmula:**

```excel
=Preço × Quantidade Vendida
```

**Exemplo:**

```excel
=C2*E2
```

Neste exemplo:

- Preço: R$ 89,00
- Vendas: 35

Faturamento = **R$ 3.115,00**

---

## Lucro

O **lucro** é o dinheiro ganho após descontar o custo dos produtos vendidos.

**Fórmula:**

```excel
=Margem × Quantidade Vendida
```

**Exemplo:**

```excel
=F2*E2
```

Neste exemplo:

- Margem: R$ 36,00
- Vendas: 35

Lucro = **R$ 1.260,00**

---

## Margem de Lucro

A **margem de lucro** indica qual porcentagem do faturamento realmente se transforma em lucro.

**Fórmula:**

```excel
=Lucro ÷ Faturamento
```

**Exemplo:**

```excel
=H2/G2
```

Se o resultado for **0,40**, basta formatar a célula como **Porcentagem (%)** para visualizar **40%**.

**Interpretação:**

- **10%** → A empresa lucra R$ 10 para cada R$ 100 vendidos.
- **25%** → A empresa lucra R$ 25 para cada R$ 100 vendidos.
- **40%** → A empresa lucra R$ 40 para cada R$ 100 vendidos.
- Quanto maior a margem de lucro, maior é a rentabilidade das vendas.
