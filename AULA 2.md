# Funções Estatísticas

| Função | Fórmula | Exemplo | Uso |
|---------|----------|----------|-----|
| Média | `=MÉDIA(Intervalo)` | `=MÉDIA(C3:C7)` | Definir a média aritmética (soma dos valores dividida pela quantidade de elementos) de um intervalo de células. |
| Mediana | `=MEDIANA(Intervalo)` | `=MEDIANA(C3:C7)` | Calcular a mediana (valor central) de um intervalo de células. |
| Máximo | `=MÁXIMO(Intervalo)` | `=MÁXIMO(C3:C7)` | Determinar o maior valor de um intervalo de dados. |
| Mínimo | `=MÍNIMO(Intervalo)` | `=MÍNIMO(C3:C7)` | Determinar o menor valor de um intervalo de dados. |

---

# Média

A **média aritmética** representa o valor médio de um conjunto de dados. Ela é calculada somando todos os valores e dividindo o resultado pela quantidade de elementos.

### Fórmula matemática

```text
Média = Soma dos valores ÷ Quantidade de valores
```

### Exemplo

Considere as notas abaixo:

```
7 | 8 | 9 | 6 | 10
```

Soma das notas:

```
7 + 8 + 9 + 6 + 10 = 40
```

Quantidade de notas:

```
5
```

Resultado:

```
40 ÷ 5 = 8
```

Logo, a **média é 8**.

### No Excel

```excel
=MÉDIA(C3:C7)
```

O Excel soma automaticamente os valores do intervalo e divide pela quantidade de células preenchidas.

---

# Mediana

A **mediana** é o valor que ocupa a posição central de um conjunto de dados quando eles são organizados em ordem crescente.

Ela é muito utilizada quando existem valores muito altos ou muito baixos, pois não sofre tanta influência desses valores extremos quanto a média.

### Exemplo 1 (quantidade ímpar)

Valores:

```
4 | 6 | 8 | 9 | 15
```

Valor central:

```
8
```

Logo, a **mediana é 8**.

### Exemplo 2 (quantidade par)

Valores:

```
2 | 4 | 8 | 10
```

Existem dois valores centrais (4 e 8). Nesse caso, calcula-se a média entre eles:

```
(4 + 8) ÷ 2 = 6
```

Logo, a **mediana é 6**.

### No Excel

```excel
=MEDIANA(C3:C7)
```

O Excel organiza automaticamente os valores e retorna o valor central (ou a média dos dois valores centrais, quando houver uma quantidade par de elementos).

---

## Diferença entre Média e Mediana

| Média | Mediana |
|--------|----------|
| Soma todos os valores e divide pela quantidade de elementos. | Encontra o valor central dos dados ordenados. |
| Pode ser influenciada por valores muito altos ou muito baixos. | Sofre pouca influência de valores extremos. |
| Indicada para conjuntos de dados equilibrados. | Indicada quando existem valores extremos (muito altos ou muito baixos). |

### Exemplo

Salários de cinco funcionários:

```
R$ 2.000
R$ 2.100
R$ 2.200
R$ 2.300
R$ 30.000
```

- **Média:** R$ 7.720
- **Mediana:** R$ 2.200

Observe que o salário de R$ 30.000 aumenta bastante a média, enquanto a mediana continua representando melhor o valor central do grupo.



# [Prática – Funções Estatísticas](

|   | **A** | **B** | **C** | **D** | **E** | **F** | **G** | **H** | **I** | **J** | **K** | **L** |
|---|--------|--------|--------|--------|--------|--------|--------|--------|--------|--------|--------|--------|
| **1** | Vendedores | 01/jun | 02/jun | 03/jun | 04/jun | 05/jun | 06/jun | Total | Média | Mediana | Máximo | Mínimo |
| **2** | Agamenon | R$ 250,00 | R$ 255,00 | R$ 210,00 | R$ 290,00 | R$ 300,00 | R$ 500,00 | `=SOMA(B2:G2)` | `=MÉDIA(B2:G2)` | `=MEDIANA(B2:G2)` | `=MÁXIMO(B2:G2)` | `=MÍNIMO(B2:G2)` |
| **3** | Astrogildo | R$ 160,00 | R$ 190,00 | R$ 200,00 | R$ 170,00 | R$ 220,00 | R$ 340,00 | | | | | |
| **4** | Balbina | R$ 120,00 | R$ 170,00 | R$ 230,00 | R$ 180,00 | R$ 210,00 | R$ 410,00 | | | | | |
| **5** | Carlota | R$ 290,00 | R$ 220,00 | R$ 250,00 | R$ 320,00 | R$ 325,00 | R$ 240,00 | | | | | |
| **6** | Ernesto | R$ 200,00 | R$ 270,00 | R$ 220,00 | R$ 210,00 | R$ 330,00 | R$ 260,00 | | | | | |
| **7** | Melquíades | R$ 220,00 | R$ 230,00 | R$ 300,00 | R$ 250,00 | R$ 400,00 | R$ 490,00 | | | | | |
| **8** | Osvalda | R$ 230,00 | R$ 240,00 | R$ 190,00 | R$ 200,00 | R$ 420,00 | R$ 400,00 | | | | | |
| **9** | Penelope | R$ 210,00 | R$ 180,00 | R$ 260,00 | R$ 270,00 | R$ 520,00 | R$ 215,00 | | | | | |
| **10** | Valmir | R$ 240,00 | R$ 200,00 | R$ 270,00 | R$ 300,00 | R$ 430,00 | R$ 330,00 | | | | | |
| **11** | Zulmira | R$ 150,00 | R$ 260,00 | R$ 240,00 | R$ 340,00 | R$ 345,00 | R$ 290,00 | | | | | |
| **12** | Total | `=SOMA(B2:B11)` | | | | | | | | | | |
| **13** | Média | `=MÉDIA(B2:B11)` | | | | | | | | | | |
| **14** | Mediana | `=MEDIANA(B2:B11)` | | | | | | | | | | |
| **15** | Máximo | `=MÁXIMO(B2:B11)` | | | | | | | | | | |
| **16** | Mínimo | `=MÍNIMO(B2:B11)` | | | | | | | | | | |

> **Objetivo da atividade:** Complete as fórmulas da linha 2 para os demais vendedores (linhas 3 a 11) utilizando a alça de preenchimento. Em seguida, calcule os indicadores gerais das vendas do dia **01/jun** (coluna B) nas linhas 12 a 16.
