# Funções CONT.NÚM, CONT.VALORES, CONT.SE e SOMASE

## 1. Principais funções

| Função           | Fórmula                                      | Exemplo de uso             | O que faz                                                                                 |
| ---------------- | -------------------------------------------- | -------------------------- | ----------------------------------------------------------------------------------------- |
| **CONT.NÚM**     | `=CONT.NÚM(Intervalo)`                       | `=CONT.NÚM(C3:C7)`         | Conta quantas células possuem **números** dentro de um intervalo.                         |
| **CONT.VALORES** | `=CONT.VALORES(Intervalo)`                   | `=CONT.VALORES(A1:A28)`    | Conta quantas células **não estão vazias**, sejam elas preenchidas com números ou textos. |
| **CONT.SE**      | `=CONT.SE(Intervalo;Critério)`               | `=CONT.SE(A1:A28;"Arroz")` | Conta quantas vezes um determinado **critério** aparece em um intervalo.                  |
| **SOMASE**       | `=SOMASE(Intervalo;Critério;Intervalo_Soma)` | `=SOMASE(D:D;"Arroz";F:F)` | Soma valores de acordo com um **critério** encontrado em outra coluna.                    |

---

## 2. Exemplos práticos

### CONT.NÚM

Imagine uma lista de notas:

| Aluno  | Nota |
| ------ | ---: |
| João   |    8 |
| Maria  |    7 |
| Pedro  |      |
| Ana    |    9 |
| Carlos |    6 |

Para descobrir **quantas notas foram preenchidas**, podemos usar:

```excel
=CONT.NÚM(B2:B6)
```

**Resultado:** `4`

A função conta apenas as células que possuem **números**.

---

### CONT.VALORES

Imagine uma lista de produtos:

| Produto  |
| -------- |
| Arroz    |
| Feijão   |
| Macarrão |
| Café     |
| Açúcar   |

Para descobrir **quantos produtos foram cadastrados**, usamos:

```excel
=CONT.VALORES(A2:A6)
```

**Resultado:** `5`

A função conta todas as células que estão **preenchidas**, sejam elas números ou textos.

---

### CONT.SE

Imagine uma lista de produtos vendidos:

| Produto |
| ------- |
| Arroz   |
| Feijão  |
| Arroz   |
| Café    |
| Arroz   |

Para descobrir **quantas vezes o produto "Arroz" aparece**, usamos:

```excel
=CONT.SE(A2:A6;"Arroz")
```

**Resultado:** `3`

A função procura o critério **"Arroz"** e conta quantas vezes ele aparece.

---

### SOMASE

Imagine uma tabela de vendas:

| Produto | Vendedor |     Valor |
| ------- | -------- | --------: |
| Arroz   | João     | R$ 100,00 |
| Feijão  | Maria    |  R$ 80,00 |
| Arroz   | Pedro    | R$ 150,00 |
| Café    | João     |  R$ 50,00 |
| Arroz   | Maria    | R$ 200,00 |

Para descobrir **quanto foi vendido somente de Arroz**, usamos:

```excel
=SOMASE(A2:A6;"Arroz";C2:C6)
```

**Resultado:** `R$ 450,00`

A função encontra as linhas em que o produto é **Arroz** e soma os valores correspondentes:

**R$ 100,00 + R$ 150,00 + R$ 200,00 = R$ 450,00**

---

# Exercício – Análise de Vendas de Veículos

## 3. [Tabela de Vendas](https://docs.google.com/spreadsheets/d/1rEwuNrEWPxj2_SNKB0jG229wotBfbtlW/export?format=xlsx&utm_source=chatgpt.com)

Considere a tabela abaixo. As letras representam as **colunas do Excel** e os números representam as **linhas**.

| **A**        | **B**               | **C**               | **D**           |
| ------------ | ------------------- | ------------------- | --------------- |
| **Data**     | **Modelo**          | **Valor**           | **Vendedor**    |
| A2 – 01/jun  | B2 – Yaris Cross    | C2 – R$ 180.000,00  | D2 – Liduvina   |
| A3 – 02/jun  | B3 – Corolla        | C3 – R$ 190.000,00  | D3 – Crispim    |
| A4 – 03/jun  | B4 – Corolla Cross  | C4 – R$ 220.000,00  | D4 – Otácilio   |
| A5 – 04/jun  | B5 – Corolla        | C5 – R$ 190.000,00  | D5 – Filomena   |
| A6 – 05/jun  | B6 – Yaris Cross    | C6 – R$ 180.000,00  | D6 – Gaudência  |
| A7 – 06/jun  | B7 – Hilux          | C7 – R$ 390.000,00  | D7 – Crispim    |
| A8 – 07/jun  | B8 – Corolla        | C8 – R$ 190.000,00  | D8 – Crispim    |
| A9 – 08/jun  | B9 – Corolla Cross  | C9 – R$ 220.000,00  | D9 – Liduvina   |
| A10 – 09/jun | B10 – Yaris Cross   | C10 – R$ 180.000,00 | D10 – Otácilio  |
| A11 – 10/jun | B11 – Hilux         | C11 – R$ 390.000,00 | D11 – Gaudência |
| A12 – 11/jun | B12 – Corolla Cross | C12 – R$ 220.000,00 | D12 – Filomena  |
| A13 – 12/jun | B13 – Yaris Cross   | C13 – R$ 180.000,00 | D13 – Otácilio  |
| A14 – 13/jun | B14 – Corolla       | C14 – R$ 190.000,00 | D14 – Filomena  |

---

# 4. Resumo por Vendedor

Crie uma tabela para descobrir quantas **unidades cada vendedor vendeu** e qual foi o **valor total de suas vendas**.

| **F**          | **G**              | **H**                 |
| -------------- | ------------------ | --------------------- |
| **Vendedor**   | **Unidades**       | **Valor**             |
| F2 – Crispim   | `=CONT.SE(D:D;F2)` | `=SOMASE(D:D;F2;C:C)` |
| F3 – Filomena  | `=CONT.SE(D:D;F3)` | `=SOMASE(D:D;F3;C:C)` |
| F4 – Gaudência | `=CONT.SE(D:D;F4)` | `=SOMASE(D:D;F4;C:C)` |
| F5 – Liduvina  | `=CONT.SE(D:D;F5)` | `=SOMASE(D:D;F5;C:C)` |
| F6 – Otácilio  | `=CONT.SE(D:D;F6)` | `=SOMASE(D:D;F6;C:C)` |
| F7 – Total     |                    |                       |

### Como resolver

#### Coluna G – Unidades

Utilize a função `CONT.SE` para contar quantas vezes o vendedor aparece na coluna **D**.

Exemplo:

```excel
=CONT.SE(D:D;F2)
```

* `D:D` → coluna onde estão os vendedores.
* `F2` → nome do vendedor que será procurado.

Depois, arraste a fórmula para as demais linhas.

#### Coluna H – Valor

Utilize a função `SOMASE` para somar os valores das vendas realizadas por cada vendedor.

Exemplo:

```excel
=SOMASE(D:D;F2;C:C)
```

* `D:D` → coluna onde estão os vendedores.
* `F2` → vendedor utilizado como critério.
* `C:C` → coluna onde estão os valores das vendas.

Depois, arraste a fórmula para os demais vendedores.

---

# 5. Resumo por Modelo

Crie uma tabela para descobrir quantas **unidades de cada modelo foram vendidas** e qual foi o **valor total vendido por modelo**.

| **F**               | **G**               | **H**                  |
| ------------------- | ------------------- | ---------------------- |
| **Modelo**          | **Unidades**        | **Valor**              |
| F10 – Yaris Cross   | `=CONT.SE(B:B;F10)` | `=SOMASE(B:B;F10;C:C)` |
| F11 – Corolla       | `=CONT.SE(B:B;F11)` | `=SOMASE(B:B;F11;C:C)` |
| F12 – Corolla Cross | `=CONT.SE(B:B;F12)` | `=SOMASE(B:B;F12;C:C)` |
| F13 – Hilux         | `=CONT.SE(B:B;F13)` | `=SOMASE(B:B;F13;C:C)` |

### Como resolver

#### Coluna G – Unidades

Use a função `CONT.SE` para contar quantas vezes cada modelo aparece na coluna **B**.

Exemplo:

```excel
=CONT.SE(B:B;F10)
```

* `B:B` → coluna onde estão os modelos.
* `F10` → modelo que será procurado.

Depois, arraste a fórmula para os demais modelos.

#### Coluna H – Valor

Use a função `SOMASE` para somar os valores de todas as vendas de cada modelo.

Exemplo:

```excel
=SOMASE(B:B;F10;C:C)
```

* `B:B` → coluna onde estão os modelos.
* `F10` → modelo utilizado como critério.
* `C:C` → coluna onde estão os valores das vendas.

Depois, arraste a fórmula para os demais modelos.

---

# 6. Resumo das Funções

| Função         | Objetivo                                                     |
| -------------- | ------------------------------------------------------------ |
| `CONT.NÚM`     | Conta quantas células possuem **números**.                   |
| `CONT.VALORES` | Conta quantas células estão **preenchidas**.                 |
| `CONT.SE`      | Conta quantas células atendem a um determinado **critério**. |
| `SOMASE`       | Soma valores que atendem a um determinado **critério**.      |

## Dica

Ao trabalhar com funções no Excel, observe sempre:

* **Qual coluna contém os dados que serão analisados?**
* **Qual é o critério que deve ser procurado?**
* **Qual coluna contém os valores que serão somados?**

Nas funções `CONT.SE` e `SOMASE`, a referência de célula, como `F2` ou `F10`, permite criar uma fórmula que pode ser **arrastada para outras linhas**, facilitando o preenchimento da tabela.
