# Aula 9 – Funções ÍNDICE e CORRESP

## Objetivos da Aula

Ao final desta aula você será capaz de:

- Compreender como funciona a função **ÍNDICE**;
- Aprender a utilizar a função **CORRESP**;
- Combinar as duas funções para criar mecanismos de busca dinâmicos;
- Entender por que essa combinação é considerada uma alternativa mais poderosa ao PROCV.

---

# Introdução

Na aula anterior aprendemos a utilizar a função **PROCV** para buscar informações em outras tabelas.

Embora seja extremamente útil, o PROCV possui algumas limitações.

Por exemplo:

- Só consegue procurar informações da esquerda para a direita;
- Depende do número da coluna, que pode mudar caso novas colunas sejam inseridas;
- Fica menos flexível em bancos de dados maiores.

Para resolver essas limitações, utilizamos a combinação das funções **ÍNDICE** e **CORRESP**.

Essa dupla é considerada uma das ferramentas mais importantes para pesquisas no Excel.

---

# A Função ÍNDICE

A função **ÍNDICE** devolve o conteúdo existente em uma posição específica de uma tabela.

Imagine uma tabela como um grande tabuleiro.

Cada informação possui uma linha e uma coluna.

O ÍNDICE apenas responde:

> "Qual informação existe nessa posição?"

---

## Sintaxe

```excel
=ÍNDICE(Matriz;Linha;Coluna)
```

---

## Exemplo

Considere a tabela abaixo.

| Linha | Produto | Valor |
|-------:|----------|-------:|
|1|Samsung S25|3990|
|2|iPhone 17|6990|
|3|Vivo X200 Ultra|9990|

Queremos retornar o valor da **segunda linha** e da **primeira coluna**.

```excel
=ÍNDICE(A2:B4;2;1)
```

Resultado:

```
iPhone 17
```

Observe que o ÍNDICE **não procura informações**.

Ele apenas retorna aquilo que está em uma posição específica.

---

# A Função CORRESP

A função CORRESP faz exatamente o contrário.

Ela não retorna um valor.

Ela retorna **a posição** onde determinado valor foi encontrado.

---

## Sintaxe

```excel
=CORRESP(Valor Procurado;Matriz;0)
```

---

## Exemplo

Tabela

| Código |
|--------:|
|10001|
|10002|
|10003|
|10004|

Queremos descobrir em qual posição está o código **10003**.

```excel
=CORRESP(10003;A2:A5;0)
```

Resultado

```
3
```

Ou seja,

10003 está na **terceira posição** da lista.

---

# Outro exemplo

Tabela

| Produto |
|----------|
|Notebook|
|Mouse|
|Teclado|
|Monitor|

```excel
=CORRESP("Teclado";A2:A5;0)
```

Resultado

```
3
```

Mais uma vez, a função devolveu apenas a posição.

---

# Juntando as duas funções

Agora vem a grande vantagem.

O CORRESP encontra a posição.

O ÍNDICE utiliza essa posição para devolver a informação.

É como se uma função encontrasse o endereço e a outra buscasse o objeto naquele endereço.

---

# Exemplo simples

Tabela

| Código | Produto |
|--------:|----------|
|10001|Samsung|
|10002|iPhone|
|10003|Motorola|

Queremos descobrir o produto referente ao código **10002**.

Primeiro:

```excel
=CORRESP(10002;A2:A4;0)
```

Resultado

```
2
```

Depois:

```excel
=ÍNDICE(B2:B4;2)
```

Resultado

```
iPhone
```

Agora podemos juntar as duas funções.

```excel
=ÍNDICE(B2:B4;CORRESP(10002;A2:A4;0))
```

Resultado

```
iPhone
```

Tudo em apenas uma fórmula.

---

# Banco de Dados

Nesta aula utilizaremos novamente a planilha da empresa de transportes.

| Data da Saída | Código | Cidade de Origem | Cidade de Destino | Motorista | Placa do Veículo | Tipo de Carga | Peso (kg) | Distância (km) | Custo Frete | Mão de Obra | Valor Cliente |
|---------------|---------|------------------|-------------------|------------|------------------|---------------|-----------:|---------------:|-------------:|-------------:|---------------:|

O objetivo será construir um mecanismo de busca que permita localizar qualquer viagem apenas informando o código.

---

# Construindo o mecanismo de busca

Nossa estrutura será semelhante à seguinte.

| Mecanismo de Busca | |
|--------------------|----------------|
| Código | **10259042** |
| Cidade de Origem | |
| Cidade de Destino | |
| Motorista | |
| Tipo de Carga | |
| Distância | |
| Valor Cobrado | |

Sempre que o código for alterado, todas as informações deverão ser atualizadas automaticamente.

---

# Como a fórmula funciona?

A fórmula será composta por duas partes.

Primeiro o **CORRESP**.

Depois o **ÍNDICE**.

---

## Primeiro passo

Descobrir em qual linha está o código digitado.

```excel
=CORRESP($O$2;B:B;0)
```

Leitura

> Procure o código digitado em **O2** dentro da coluna **B**.

Resultado

```
42
```

Isso significa que a viagem está na linha 42 da tabela.

---

## Segundo passo

Agora precisamos descobrir em qual coluna está a informação desejada.

Por exemplo:

"Cidade de origem"

```excel
=CORRESP(N3;$1:$1;0)
```

Leitura

> Procure o texto **Cidade de origem** na primeira linha da tabela.

Resultado

```
3
```

Ou seja,

Cidade de origem é a terceira coluna da tabela.

---

# Agora juntamos tudo

```excel
=ÍNDICE(
Sereia;
CORRESP($O$2;B:B;0);
CORRESP(N3;$1:$1;0)
)
```

---

## Entendendo a fórmula

Primeiro o Excel encontra a linha do código.

```excel
CORRESP($O$2;B:B;0)
```

Depois encontra a coluna correspondente ao nome do campo.

```excel
CORRESP(N3;$1:$1;0)
```

Por fim, o ÍNDICE devolve a informação localizada exatamente naquela linha e coluna.

---

# Cidade de Origem

```excel
=ÍNDICE(Sereia;CORRESP($O$2;B:B;0);CORRESP(N3;$1:$1;0))
```

---

# Cidade de Destino

```excel
=ÍNDICE(Sereia;CORRESP($O$2;B:B;0);CORRESP(N4;$1:$1;0))
```

---

# Observe a vantagem

Não precisamos informar:

- coluna 3;
- coluna 5;
- coluna 8.

O próprio Excel descobre qual coluna deve utilizar.

Se uma nova coluna for inserida na tabela, as fórmulas continuarão funcionando normalmente.

Essa é a principal vantagem da combinação **ÍNDICE + CORRESP** em relação ao **PROCV**.

---



# Conclusão

A função **ÍNDICE** retorna o conteúdo existente em uma determinada posição de uma tabela, enquanto a função **CORRESP** localiza a posição onde um valor é encontrado. Quando utilizadas em conjunto, elas formam um mecanismo de busca extremamente flexível e poderoso, permitindo localizar qualquer informação em uma tabela sem depender da posição fixa das colunas.

Essa combinação é amplamente utilizada em planilhas profissionais e serve de base para compreender consultas em bancos de dados e modelos relacionais utilizados posteriormente em ferramentas como Access, SQL Server e Power BI.
