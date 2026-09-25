# Operadores Lógicos no Excel

Os **operadores lógicos** são utilizados no Excel para realizar comparações entre valores. Eles permitem verificar se uma determinada condição é verdadeira ou falsa.

Quando realizamos um teste lógico no Excel, o resultado será sempre um dos dois valores:

* **VERDADEIRO**: quando a condição analisada é correta.
* **FALSO**: quando a condição analisada não é correta.

Por exemplo, podemos utilizar a fórmula `=67>42`. Nesse caso, o Excel verifica se **67 é maior que 42**. Como essa afirmação é verdadeira, o resultado será **VERDADEIRO**.

Da mesma forma, se utilizarmos a fórmula `=390<247`, o Excel verificará se **390 é menor que 247**. Como essa afirmação é falsa, o resultado será **FALSO**.

## Operador Maior que (>)

O símbolo `>` significa **maior que**. Ele é utilizado quando queremos verificar se um determinado valor é superior a outro.

Por exemplo:

```excel
=67>42
```

O Excel verifica se **67 é maior que 42**. Como a condição é verdadeira, o resultado será:

```text
VERDADEIRO
```

Outro exemplo seria:

```excel
=10>20
```

Nesse caso, como **10 não é maior que 20**, o resultado será:

```text
FALSO
```

---

## Operador Menor que (<)

O símbolo `<` significa **menor que**. Ele verifica se um valor é inferior a outro.

Por exemplo:

```excel
=390<247
```

O Excel verifica se **390 é menor que 247**. Como isso não é verdade, o resultado será:

```text
FALSO
```

Já na fórmula:

```excel
=15<30
```

A condição é verdadeira, pois **15 é menor que 30**. Portanto, o resultado será:

```text
VERDADEIRO
```

---

## Operador Maior ou igual (>=)

O símbolo `>=` significa **maior ou igual a**. Ele verifica se um valor é maior que outro ou se os dois valores são iguais.

Por exemplo:

```excel
=7>=6
```

Como **7 é maior que 6**, o resultado será:

```text
VERDADEIRO
```

Também podemos testar valores iguais:

```excel
=7>=7
```

O resultado também será:

```text
VERDADEIRO
```

Isso acontece porque a condição permite que o primeiro valor seja **maior ou igual** ao segundo.

Por outro lado:

```excel
=5>=8
```

O resultado será:

```text
FALSO
```

Isso ocorre porque **5 não é maior nem igual a 8**.

---

## Operador Menor ou igual (<=)

O símbolo `<=` significa **menor ou igual a**. Ele verifica se um valor é menor que outro ou se os dois valores são iguais.

Por exemplo:

```excel
=3<=5
```

Como **3 é menor que 5**, o resultado será:

```text
VERDADEIRO
```

Também podemos comparar valores iguais:

```excel
=5<=5
```

Nesse caso, o resultado será:

```text
VERDADEIRO
```

Isso acontece porque o operador permite que o primeiro valor seja **menor ou igual** ao segundo.

Já na fórmula:

```excel
=10<=5
```

O resultado será:

```text
FALSO
```

---

## Operador Diferente (<>)

O símbolo `<>` significa **diferente de**. Ele verifica se dois valores são diferentes.

Por exemplo:

```excel
=2<>8
```

Como **2 é diferente de 8**, o resultado será:

```text
VERDADEIRO
```

Já na fórmula:

```excel
=5<>5
```

O resultado será:

```text
FALSO
```

Isso acontece porque os dois valores são iguais.

Esse operador pode ser muito útil para verificar, por exemplo, se uma célula está diferente de determinado valor.

---

## Operador Igual (=)

O símbolo `=` pode ser utilizado para verificar se dois valores são iguais.

Por exemplo:

```excel
=3=3
```

Como os dois valores são iguais, o resultado será:

```text
VERDADEIRO
```

Se utilizarmos:

```excel
=3=8
```

Como os valores são diferentes, o resultado será:

```text
FALSO
```

O operador de igualdade é muito utilizado em testes lógicos para verificar se uma informação corresponde exatamente ao valor esperado.

---

# Resumo dos Operadores Lógicos

| Símbolo | Operação       | Exemplo    | Resultado  |
| ------- | -------------- | ---------- | ---------- |
| `>`     | Maior que      | `=67>42`   | VERDADEIRO |
| `<`     | Menor que      | `=390<247` | FALSO      |
| `>=`    | Maior ou igual | `=7>=6`    | VERDADEIRO |
| `<=`    | Menor ou igual | `=3<=5`    | VERDADEIRO |
| `<>`    | Diferente      | `=2<>8`    | VERDADEIRO |
| `=`     | Igual          | `=3=3`     | VERDADEIRO |

---

# Função SE

A função `SE` é uma das funções lógicas mais importantes do Excel. Ela permite realizar um teste e apresentar um resultado diferente dependendo de a condição ser **VERDADEIRA** ou **FALSA**.

A estrutura básica da função é:

```excel
=SE(Teste_Lógico;Valor_Se_Verdadeiro;Valor_Se_Falso)
```

Podemos entender essa fórmula da seguinte maneira:

> **SE** uma determinada condição for verdadeira, faça uma coisa; **caso contrário**, faça outra.

Por exemplo, imagine uma planilha de notas de alunos. Para ser aprovado, o aluno precisa obter uma média maior ou igual a **5**.

Podemos utilizar a seguinte fórmula:

```excel
=SE(C3>=5;"AP";"RP")
```

Nesse exemplo, o Excel verifica o valor que está na célula `C3`.

Se a nota for **maior ou igual a 5**, o resultado será:

```text
AP
```

Nesse caso, podemos interpretar `AP` como **Aprovado**.

Se a nota for **menor que 5**, o resultado será:

```text
RP
```

Nesse caso, podemos interpretar `RP` como **Reprovado**.

### Exemplo prático

Imagine que a célula `C3` contenha a nota `7`.

A fórmula:

```excel
=SE(C3>=5;"AP";"RP")
```

fará o seguinte teste:

```text
7 >= 5
```

Como a condição é verdadeira, o Excel exibirá:

```text
AP
```

Agora imagine que a célula `C3` contenha a nota `4`.

O Excel realizará o teste:

```text
4 >= 5
```

Como a condição é falsa, o resultado será:

```text
RP
```

---

# Exemplo com Situação do Aluno

Podemos utilizar a função `SE` para criar uma classificação mais clara para os alunos.

Suponha que a nota esteja na célula `C3`. A fórmula:

```excel
=SE(C3>=5;"Aprovado";"Reprovado")
```

retornará diretamente o texto correspondente à situação do aluno.

Se `C3` for igual a `8`, teremos:

```text
Aprovado
```

Se `C3` for igual a `3`, teremos:

```text
Reprovado
```

Dessa forma, a função `SE` permite transformar um simples teste lógico em uma informação mais fácil de interpretar.

---

# Exemplo com Controle de Estoque

A função `SE` também pode ser utilizada para controlar o estoque de produtos.

Imagine que a quantidade disponível de um produto esteja na célula `B2`. Podemos definir que o estoque mínimo é de **10 unidades**.

A fórmula:

```excel
=SE(B2<10;"Repor estoque";"Estoque normal")
```

irá verificar a quantidade disponível.

Se `B2` for igual a `5`, o resultado será:

```text
Repor estoque
```

Se `B2` for igual a `25`, o resultado será:

```text
Estoque normal
```

Assim, a função `SE` pode ajudar a identificar automaticamente quais produtos precisam ser repostos.

---

# Exemplo com Idade

Também podemos utilizar a função `SE` para verificar a idade de uma pessoa.

Suponha que a idade esteja na célula `A2`. Podemos utilizar:

```excel
=SE(A2>=18;"Maior de idade";"Menor de idade")
```

Se `A2` contiver `25`, o resultado será:

```text
Maior de idade
```

Se `A2` contiver `15`, o resultado será:

```text
Menor de idade
```

Nesse caso, o operador `>=` é utilizado para verificar se a idade é **maior ou igual a 18**.

---

# [PRÁTICA 3](https://docs.google.com/spreadsheets/d/1CafXz-PD7Q7TW4Xty8Y1SrNj6BcYmLRt/export?format=xlsx&utm_source=chatgpt.com)

# Notas - Excel Completo

**Professor:** Victor Passos

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| **Nº** | **Nomes** | **PROVA 1** | **PROVA 2** | **PROVA 3** | **Média** | **Aprovado?** | **Função SE** |
| 1 | Ambrósio | 8,4 | 5 | 7,2 | `=MÉDIA(C3:E3)` | `=F3>=5` | `=SE(F3>=5;"APROVADO";"REPROVADO")` |
| 2 | Arlinda | 5,1 | 4,7 | 3,7 | | | |
| 3 | Clarimunda | 7,8 | 5,4 | 7,7 | | | |
| 4 | Deodato | 10 | 9,5 | 10 | | | |
| 5 | Eudóxia | 7,3 | 8 | 5,9 | | | |
| 6 | Gervásio | 1 | 7,1 | 6,4 | | | |
| 7 | Laurinda | 10 | 10 | 5,3 | | | |
| 8 | Norberta | 9,8 | 4,2 | 8,3 | | | |
| 9 | Rufino | 10 | 10 | 7,1 | | | |
| 10 | Ubaldo | 8,4 | 4,5 | 9 | | | |
| 11 | Venância | 9,1 | 5,8 | 8 | | | |
| 12 | Zózimo | 3,3 | 8 | 3 | | | |

# Como resolver as colunas

## Coluna F — Média

Na coluna **Média**, devemos calcular a média das três provas de cada aluno.

Para isso, utilizamos a função `MÉDIA`, selecionando as células que contêm as notas das provas.

Na primeira linha de aluno, digite:

`=MÉDIA(C2:E2)`

Depois, pressione **Enter**. O Excel calculará automaticamente a média das três notas.

Após fazer o primeiro cálculo, utilize a **alça de preenchimento** para arrastar a fórmula para baixo e calcular a média dos demais alunos.

---

## Coluna G — Aprovado?

Na coluna **Aprovado?**, vamos verificar se a média do aluno é maior ou igual a 5.

Para isso, utilizamos um **teste lógico**:

`=F2>=5`

Se a média for maior ou igual a 5, o Excel retornará **VERDADEIRO**.

Se a média for menor que 5, o Excel retornará **FALSO**.

Depois, arraste a fórmula para baixo para verificar automaticamente os demais alunos.

---

## Coluna H — Função SE

Na coluna **Função SE**, vamos transformar o resultado do teste lógico em uma informação mais clara para o usuário.

Utilize a fórmula:

`=SE(F2>=5;"APROVADO";"REPROVADO")`

A fórmula verifica a média do aluno:

- Se a média for **maior ou igual a 5**, aparecerá **APROVADO**.
- Se a média for **menor que 5**, aparecerá **REPROVADO**.

Depois de inserir a fórmula no primeiro aluno, arraste-a para baixo para preencher automaticamente os demais resultados.

---

## Resumo

| Coluna | O que fazer | Fórmula |
|---|---|---|
| **F — Média** | Calcular a média das três provas | `=MÉDIA(C2:E2)` |
| **G — Aprovado?** | Verificar se a média é maior ou igual a 5 | `=F2>=5` |
| **H — Função SE** | Informar "APROVADO" ou "REPROVADO" | `=SE(F2>=5;"APROVADO";"REPROVADO")` |

**Dica:** Depois de preencher a primeira linha com as fórmulas, utilize a **alça de preenchimento** (pequeno quadrado no canto inferior direito da célula) e arraste para baixo. O Excel ajustará automaticamente as referências das células para cada aluno.

# Conclusão

Os operadores lógicos são fundamentais para realizar comparações no Excel. Eles permitem verificar relações entre valores e retornam os resultados **VERDADEIRO** ou **FALSO**.

A função `SE` utiliza esses testes lógicos para tomar decisões automaticamente. Com ela, podemos fazer com que o Excel apresente diferentes resultados dependendo da situação analisada.

Alguns exemplos de aplicações práticas são:

* Verificar se um aluno foi aprovado ou reprovado.
* Identificar se um produto precisa ser reposto no estoque.
* Classificar uma pessoa como maior ou menor de idade.
* Verificar se uma meta foi atingida.
* Identificar se uma venda atingiu determinado valor.
* Informar se uma conta está paga ou em atraso.

Em resumo, podemos pensar na função `SE` como uma pergunta feita ao Excel:

> **"Se essa condição for verdadeira, o que devo mostrar? E se for falsa, o que devo mostrar?"**

Essa lógica é a base para criar planilhas mais inteligentes e automatizadas.
