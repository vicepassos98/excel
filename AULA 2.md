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
