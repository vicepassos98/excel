# Aula – Análise de Crédito no Excel

Nesta aula, vamos utilizar funções **lógicas, financeiras e de busca** para criar uma análise de crédito de clientes.

As funções utilizadas serão:

* `E`
* `SE`
* `PGTO`
* `CONT.SE`
* `SOMASE`
* `PROCV`

O objetivo é verificar automaticamente se um cliente pode ter seu empréstimo aprovado, calcular suas parcelas, analisar os empréstimos liberados e criar um mecanismo de busca para consultar os dados de cada cliente.

---

# 1. Funções utilizadas

| Função      | Fórmula                                                | Exemplo                             | Uso                                                                  |
| ----------- | ------------------------------------------------------ | ----------------------------------- | -------------------------------------------------------------------- |
| **E**       | `=E(Teste_Lógico_1;Teste_Lógico_2;...)`                | `=E(5>4;10>8)`                      | Verifica se **todas** as condições são verdadeiras.                  |
| **SE**      | `=SE(Teste_Lógico;Valor_Se_Verdadeiro;Valor_Se_Falso)` | `=SE(A1>=5;"APROVADO";"REPROVADO")` | Retorna um resultado diferente dependendo de uma condição.           |
| **PGTO**    | `=PGTO(TAXA;NP;VP;VF;TIPO)`                            | `=PGTO(1,5%;12;-25000;0;0)`         | Calcula o valor das parcelas de um financiamento ou empréstimo.      |
| **CONT.SE** | `=CONT.SE(Intervalo;Critério)`                         | `=CONT.SE(L3:L29;"APROVADO")`       | Conta quantas vezes determinado critério aparece em um intervalo.    |
| **SOMASE**  | `=SOMASE(Intervalo;Critério;Intervalo_Soma)`           | `=SOMASE(L3:L29;"APROVADO";E3:E29)` | Soma valores que atendem a um determinado critério.                  |
| **PROCV**   | `=PROCV(Valor_Proc;Matriz;Nº_Coluna;FALSO)`            | `=PROCV(P2;A3:L29;2;FALSO)`         | Procura um valor em uma tabela e retorna uma informação relacionada. |

---

# 2. Função E

A função `E` é utilizada quando precisamos verificar **duas ou mais condições ao mesmo tempo**.

Para que o resultado seja **VERDADEIRO**, todas as condições precisam ser verdadeiras.

## Exemplo

```excel
=E(5>4;10>8)
```

Resultado:

**VERDADEIRO**

As duas condições são verdadeiras:

* `5>4` → VERDADEIRO
* `10>8` → VERDADEIRO

Agora:

```excel
=E(5>4;10<8)
```

Resultado:

**FALSO**

A primeira condição é verdadeira, mas a segunda é falsa.

### No exercício

O cliente precisa atender a duas condições:

1. A parcela deve ser menor que 1/3 da renda.
2. O cliente não pode ter restrições.

Por isso usamos:

```excel
=E(I3;J3)
```

Somente quando **I3 e J3 forem VERDADEIRO**, o resultado será VERDADEIRO.

---

# 3. Função SE

A função `SE` permite apresentar um resultado quando uma condição é verdadeira e outro resultado quando ela é falsa.

## Exemplo

```excel
=SE(A1>=5;"APROVADO";"NÃO APROVADO")
```

Se A1 for igual ou maior que 5:

**APROVADO**

Se A1 for menor que 5:

**NÃO APROVADO**

### No exercício

Depois de realizar os testes de renda e restrição, podemos usar:

```excel
=SE(K3;"APROVADO";"NÃO APROVADO")
```

Se `K3` for VERDADEIRO:

**APROVADO**

Se `K3` for FALSO:

**NÃO APROVADO**

Também podemos colocar as condições diretamente na função:

```excel
=SE(E(C3/3>G3;D3="Não");"APROVADO";"NÃO APROVADO")
```

---

# 4. Função PGTO

A função `PGTO` calcula o valor das parcelas de um financiamento ou empréstimo.

## Sintaxe

```excel
=PGTO(TAXA;NP;VP;VF;TIPO)
```

| Parâmetro | Significado                                                            |
| --------- | ---------------------------------------------------------------------- |
| **TAXA**  | Taxa de juros de cada período.                                         |
| **NP**    | Número total de parcelas.                                              |
| **VP**    | Valor presente, ou seja, valor financiado.                             |
| **VF**    | Valor futuro ou saldo devedor ao final.                                |
| **TIPO**  | `0` para pagamento no final do período e `1` para pagamento no início. |

### Exemplo

Um empréstimo de **R$ 25.000,00**, em **12 parcelas**, com juros de **1,5% ao mês**:

```excel
=PGTO(1,5%;12;-25000;0;0)
```

O Excel calculará o valor de cada parcela.

---

# 5. Função CONT.SE

A função `CONT.SE` conta quantas células atendem a um determinado critério.

## Exemplo

Imagine:

| A            |
| ------------ |
| APROVADO     |
| NÃO APROVADO |
| APROVADO     |
| APROVADO     |

Para contar quantos empréstimos foram aprovados:

```excel
=CONT.SE(A1:A4;"APROVADO")
```

Resultado:

**3**

No exercício:

```excel
=CONT.SE(L3:L29;"APROVADO")
```

Essa fórmula conta quantos clientes tiveram seus empréstimos aprovados.

---

# 6. Função SOMASE

A função `SOMASE` soma valores que atendem a determinado critério.

## Exemplo

| A            |         B |
| ------------ | --------: |
| APROVADO     | R$ 10.000 |
| NÃO APROVADO | R$ 20.000 |
| APROVADO     | R$ 15.000 |

Para somar apenas os valores dos aprovados:

```excel
=SOMASE(A1:A3;"APROVADO";B1:B3)
```

Resultado:

**R$ 25.000**

No exercício:

```excel
=SOMASE(L3:L29;"APROVADO";E3:E29)
```

A fórmula soma apenas os valores solicitados pelos clientes aprovados.

---

# 7. Função PROCV

A função `PROCV` permite procurar um valor na primeira coluna de uma tabela e retornar uma informação relacionada.

## Sintaxe

```excel
=PROCV(Valor_Proc;Matriz;Nº_Coluna;FALSO)
```

### Exemplo

Se procurarmos o código `103` em uma tabela:

```excel
=PROCV(F2;A2:D5;3;FALSO)
```

O Excel procura o código na primeira coluna da matriz e retorna o valor da terceira coluna.

O `FALSO` indica que queremos uma **correspondência exata**.

---

# 8. Banco Valor Ant – Análise de Crédito

## Regras para Liberação de Crédito

Para que um empréstimo seja aprovado, **as duas condições abaixo precisam ser verdadeiras**:

1. O cliente deve estar **sem restrições**.
2. O valor da parcela deve ser **menor que 1/3 da renda mensal**.

A taxa de juros do banco é de **1,5% ao mês**.

Portanto:

> **APROVADO = Parcela menor que 1/3 da renda E cliente sem restrição**

---

# 9. Tabela Principal

A tabela começa na **linha 2**.

As letras representam as **colunas do Excel** e os números representam as **linhas**.

| **A**          | **B**         | **C**             | **D**             | **E**                | **F**             | **G**                     | **H**           | **I**           | **J**               | **K**       | **L**         |
| -------------- | ------------- | ----------------- | ----------------- | -------------------- | ----------------- | ------------------------- | --------------- | --------------- | ------------------- | ----------- | ------------- |
| **CD_Cliente** | **Nome**      | **Renda Mensal**  | **Com Restrição** | **Valor Solicitado** | **Qtde Parcelas** | **Valor Parcelas (PGTO)** | **Valor Total** | **TESTE RENDA** | **TESTE RESTRIÇÃO** | **TESTE E** | **Aprovado?** |
| – 7678291   |  – Luna     | C3 – R$ 3.500,00  | D3 – Não          | E3 – R$ 25.000,00    | F3 – 12           | G3                        |               |               | J3                  | K3          | L3            |
| A4 – 7178625   | B4 – Gael     | C4 – R$ 4.000,00  | D4 – Não          | E4 – R$ 18.000,00    | F4 – 24           | G4                        | H4              | I4              | J4                  | K4          | L4            |
| A5 – 7583578   | B5 – Kai      | C5 – R$ 3.200,00  | D5 – Sim          | E5 – R$ 5.000,00     | F5 – 6            | G5                        | H5              | I5              | J5                  | K5          | L5            |
| A6 – 7778523   | B6 – Maitê    | C6 – R$ 5.000,00  | D6 – Não          | E6 – R$ 12.000,00    | F6 – 36           | G6                        | H6              | I6              | J6                  | K6          | L6            |
| A7 – 7883248   | B7 – Noah     | C7 – R$ 2.800,00  | D7 – Sim          | E7 – R$ 7.000,00     | F7 – 10           | G7                        | H7              | I7              | J7                  | K7          | L7            |
| A8 – 9078570   | B8 – Zion     | C8 – R$ 3.900,00  | D8 – Não          | E8 – R$ 18.000,00    | F8 – 30           | G8                        | H8              | I8              | J8                  | K8          | L8            |
| A9 – 6578647   | B9 – Aurora   | C9 – R$ 2.700,00  | D9 – Não          | E9 – R$ 6.000,00     | F9 – 8            | G9                        | H9              | I9              | J9                  | K9          | L9            |
| A10 – 8483170  | B10 – Theo    | C10 – R$ 3.300,00 | D10 – Sim         | E10 – R$ 25.000,00   | F10 – 24          | G10                       | H10             | I10             | J10                 | K10         | L10           |
| A11 – 6578470  | B11 – Ariel   | C11 – R$ 4.200,00 | D11 – Não         | E11 – R$ 16.000,00   | F11 – 36          | G11                       | H11             | I11             | J11                 | K11         | L11           |
| A12 – 7678225  | B12 – Liz     | C12 – R$ 3.100,00 | D12 – Não         | E12 – R$ 7.500,00    | F12 – 10          | G12                       | H12             | I12             | J12                 | K12         | L12           |
| A13 – 8283455  | B13 – Ravi    | C13 – R$ 2.900,00 | D13 – Sim         | E13 – R$ 6.500,00    | F13 – 12          | G13                       | H13             | I13             | J13                 | K13         | L13           |
| A14 – 8978404  | B14 – Yuki    | C14 – R$ 3.700,00 | D14 – Não         | E14 – R$ 9.500,00    | F14 – 15          | G14                       | H14             | I14             | J14                 | K14         | L14           |
| A15 – 7383367  | B15 – Isis    | C15 – R$ 8.100,00 | D15 – Sim         | E15 – R$ 25.000,00   | F15 – 16          | G15                       | H15             | I15             | J15                 | K15         | L15           |
| A16 – 6983791  | B16 – Enzo    | C16 – R$ 8.600,00 | D16 – Sim         | E16 – R$ 99.000,00   | F16 – 50          | G16                       | H16             | I16             | J16                 | K16         | L16           |
| A17 – 7678239  | B17 – Luca    | C17 – R$ 3.200,00 | D17 – Não         | E17 – R$ 3.500,00    | F17 – 34          | G17                       | H17             | I17             | J17                 | K17         | L17           |
| A18 – 6578967  | B18 – Ayla    | C18 – R$ 4.100,00 | D18 – Não         | E18 – R$ 14.000,00   | F18 – 20          | G18                       | H18             | I18             | J18                 | K18         | L18           |
| A19 – 7678994  | B19 – Levi    | C19 – R$ 6.700,00 | D19 – Não         | E19 – R$ 24.000,00   | F19 – 60          | G19                       | H19             | I19             | J19                 | K19         | L19           |
| A20 – 7883798  | B20 – Nalu    | C20 – R$ 7.100,00 | D20 – Sim         | E20 – R$ 160,00      | F20 – 60          | G20                       | H20             | I20             | J20                 | K20         | L20           |
| A21 – 9078164  | B21 – Zoe     | C21 – R$ 6.000,00 | D21 – Não         | E21 – R$ 30.000,00   | F21 – 24          | G21                       | H21             | I21             | J21                 | K21         | L21           |
| A22 – 6878777  | B22 – Davi    | C22 – R$ 3.400,00 | D22 – Não         | E22 – R$ 16.000,00   | F22 – 36          | G22                       | H22             | I22             | J22                 | K22         | L22           |
| A23 – 8383415  | B23 – Sol     | C23 – R$ 3.600,00 | D23 – Sim         | E23 – R$ 1.000,00    | F23 – 48          | G23                       | H23             | I23             | J23                 | K23         | L23           |
| A24 – 6678184  | B24 – Bella   | C24 – R$ 5.500,00 | D24 – Não         | E24 – R$ 65.000,00   | F24 – 18          | G24                       | H24             | I24             | J24                 | K24         | L24           |
| A25 – 6683143  | B25 – Benício | C25 – R$ 3.600,00 | D25 – Sim         | E25 – R$ 24.000,00   | F25 – 36          | G25                       | H25             | I25             | J25                 | K25         | L25           |
| A26 – 8383210  | B26 – Sky     | C26 – R$ 7.500,00 | D26 – Sim         | E26 – R$ 14.000,00   | F26 – 24          | G26                       | H26             | I26             | J26                 | K26         | L26           |
| A27 – 7778893  | B27 – Malu    | C27 – R$ 4.800,00 | D27 – Não         | E27 – R$ 30.000,00   | F27 – 48          | G27                       | H27             | I27             | J27                 | K27         | L27           |
| A28 – 7678531  | B28 – Luan    | C28 – R$ 9.900,00 | D28 – Não         | E28 – R$ 35.000,00   | F28 – 12          | G28                       | H28             | I28             | J28                 | K28         | L28           |
| A29 – 7478101  | B29 – Jade    | C29 – R$ 9.000,00 | D29 – Não         | E29 – R$ 99.999,00   | F29 – 72          | G29                       | H29             | I29             | J29                 | K29         | L29           |

---

# 10. Preenchendo a coluna G – Valor das Parcelas

Na célula **G3**, utilize:

```excel
=PGTO(1,5%;F3;-E3;0;0)
```

A fórmula utiliza:

* `1,5%` → taxa de juros mensal.
* `F3` → quantidade de parcelas.
* `-E3` → valor solicitado.
* `0` → valor futuro.
* `0` → pagamento no final do período.

Depois, arraste a fórmula de **G3 até G29**.

---

# 11. Preenchendo a coluna H – Valor Total

Na célula **H3**:

```excel
=F3*G3
```

A fórmula calcula:

> **Quantidade de parcelas × Valor da parcela**

Depois, arraste de **H3 até H29**.

---

# 12. Coluna I – TESTE RENDA

A regra é que a parcela seja menor que 1/3 da renda.

Na célula **I3**:

```excel
=C3/3>G3
```

A fórmula verifica:

> **1/3 da renda é maior que o valor da parcela?**

Se sim:

**VERDADEIRO**

Se não:

**FALSO**

Depois, arraste de **I3 até I29**.

### Exemplo

Renda:

**R$ 3.500,00**

1/3 da renda:

**R$ 1.166,67**

Se a parcela for R$ 1.000,00:

```text
1.166,67 > 1.000,00
```

Resultado:

**VERDADEIRO**

---

# 13. Coluna J – TESTE RESTRIÇÃO

Na célula **J3**:

```excel
=D3="Não"
```

Se D3 for **Não**:

**VERDADEIRO**

Se D3 for **Sim**:

**FALSO**

Depois, arraste de **J3 até J29**.

---

# 14. Coluna K – TESTE E

Agora vamos unir os dois testes.

Na célula **K3**:

```excel
=E(I3;J3)
```

Também podemos escrever:

```excel
=E(C3/3>G3;D3="Não")
```

O cliente só terá resultado **VERDADEIRO** se:

* Passar no teste de renda.
* Não possuir restrição.

Depois, arraste de **K3 até K29**.

---

# 15. Coluna L – Aprovado?

Na célula **L3**:

```excel
=SE(K3;"APROVADO";"NÃO APROVADO")
```

Outra opção:

```excel
=SE(E(C3/3>G3;D3="Não");"APROVADO";"NÃO APROVADO")
```

O resultado será:

* **APROVADO** → duas condições atendidas.
* **NÃO APROVADO** → uma ou mais condições não atendidas.

Depois, arraste de **L3 até L29**.

---

# 16. Resumo dos Empréstimos Liberados

Crie uma área de resumo:

| **D**                     | **E**          | **F**                | **G**           |
| ------------------------- | -------------- | -------------------- | --------------- |
| **Empréstimos Liberados** | **Quantidade** | **Valor Solicitado** | **Valor Total** |
|                           | Fórmula        | Fórmula              | Fórmula         |

## Quantidade

```excel
=CONT.SE(L3:L29;"APROVADO")
```

Conta quantos clientes foram aprovados.

## Valor Solicitado

```excel
=SOMASE(L3:L29;"APROVADO";E3:E29)
```

Soma o valor solicitado somente pelos clientes aprovados.

## Valor Total

```excel
=SOMASE(L3:L29;"APROVADO";H3:H29)
```

Soma o valor total dos financiamentos somente dos clientes aprovados.

---

# 17. Mecanismo de Busca com PROCV

Agora vamos criar uma área de consulta.

| **O**              | **P**   |
| ------------------ | ------- |
| **CD_Cliente**     | 7678291 |
| **Nome**           | Fórmula |
| **Aprovado?**      | Fórmula |
| **Valor Parcelas** | Fórmula |
| **Qtde Parcelas**  | Fórmula |

O código do cliente será digitado em **P2**.

Por exemplo:

```text
P2 = 7678291
```

A partir desse código, o Excel buscará os dados na tabela principal.

---

## Nome

Na célula **P3**:

```excel
=PROCV(P2;A3:L29;2;FALSO)
```

Retorna o nome do cliente.

Para `7678291`:

**Luna**

---

## Aprovado?

Na célula **P4**:

```excel
=PROCV(P2;A3:L29;12;FALSO)
```

Retorna:

**APROVADO**

ou

**NÃO APROVADO**

A coluna L é a **12ª coluna** da matriz `A:L`.

---

## Valor das Parcelas

Na célula **P5**:

```excel
=PROCV(P2;A3:L29;7;FALSO)
```

A coluna G é a **7ª coluna** da matriz.

---

## Quantidade de Parcelas

Na célula **P6**:

```excel
=PROCV(P2;A3:L29;6;FALSO)
```

A coluna F é a **6ª coluna** da matriz.

---

# 18. Estrutura Final do Mecanismo de Busca

| **O**              | **P**                        |
| ------------------ | ---------------------------- |
| **CD_Cliente**     | `7678291`                    |
| **Nome**           | `=PROCV(P2;A3:L29;2;FALSO)`  |
| **Aprovado?**      | `=PROCV(P2;A3:L29;12;FALSO)` |
| **Valor Parcelas** | `=PROCV(P2;A3:L29;7;FALSO)`  |
| **Qtde Parcelas**  | `=PROCV(P2;A3:L29;6;FALSO)`  |

Ao alterar o código em **P2**, o Excel automaticamente buscará as informações do novo cliente.

### Exemplo

Se alterarmos:

```text
P2 = 7178625
```

O mecanismo de busca encontrará os dados de:

**Gael**

---

# 19. Fluxo da Atividade

A análise de crédito funciona seguindo esta sequência:

**1. PGTO**

Calcula o valor da parcela.

↓

**2. Valor Total**

Calcula o valor total que será pago.

↓

**3. TESTE RENDA**

Verifica se a parcela é menor que 1/3 da renda.

↓

**4. TESTE RESTRIÇÃO**

Verifica se o cliente está sem restrição.

↓

**5. TESTE E**

Verifica se as duas condições são verdadeiras.

↓

**6. SE**

Classifica o cliente como **APROVADO** ou **NÃO APROVADO**.

↓

**7. CONT.SE**

Conta quantos empréstimos foram aprovados.

↓

**8. SOMASE**

Calcula o valor total solicitado e o valor total dos empréstimos aprovados.

↓

**9. PROCV**

Permite consultar os dados de um cliente pelo seu código.

---

# 20. Resumo Geral das Funções

| Função    | Objetivo                                             | Exemplo                             |
| --------- | ---------------------------------------------------- | ----------------------------------- |
| `E`       | Verifica se **todas as condições** são verdadeiras.  | `=E(I3;J3)`                         |
| `SE`      | Retorna um resultado dependendo de uma condição.     | `=SE(K3;"APROVADO";"NÃO APROVADO")` |
| `PGTO`    | Calcula o valor das parcelas.                        | `=PGTO(1,5%;F3;-E3;0;0)`            |
| `CONT.SE` | Conta células que atendem a um critério.             | `=CONT.SE(L3:L29;"APROVADO")`       |
| `SOMASE`  | Soma valores de acordo com um critério.              | `=SOMASE(L3:L29;"APROVADO";E3:E29)` |
| `PROCV`   | Busca um valor e retorna uma informação relacionada. | `=PROCV(P2;A3:L29;2;FALSO)`         |

---

# 21. Resumo para Memorizar

### E

> **TODAS as condições precisam ser verdadeiras.**

```excel
=E(condição1;condição2)
```

### SE

> **Se a condição for verdadeira, faça isso; caso contrário, faça aquilo.**

```excel
=SE(condição;resultado1;resultado2)
```

### PGTO

> **Calcula o valor da parcela.**

```excel
=PGTO(taxa;parcelas;valor)
```

### CONT.SE

> **Conta quantas vezes algo acontece.**

```excel
=CONT.SE(intervalo;critério)
```

### SOMASE

> **Soma valores de acordo com um critério.**

```excel
=SOMASE(intervalo;critério;intervalo_soma)
```

### PROCV

> **Procura um valor e traz uma informação relacionada.**

```excel
=PROCV(valor;matriz;coluna;FALSO)
```

---

## Dica final

Ao resolver a atividade, siga a ordem das colunas:

**G → H → I → J → K → L**

1. **G:** Calcule a parcela com `PGTO`.
2. **H:** Calcule o valor total.
3. **I:** Faça o teste da renda.
4. **J:** Faça o teste de restrição.
5. **K:** Una os testes com `E`.
6. **L:** Mostre o resultado com `SE`.

Depois:

7. Use `CONT.SE` para contar os aprovados.
8. Use `SOMASE` para somar os valores dos aprovados.
9. Use `PROCV` para criar o mecanismo de busca.

Dessa forma, a planilha transforma uma análise manual de crédito em um processo **automático e dinâmico**.
