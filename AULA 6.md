# Aula 06 – Relacionando Tabelas com PROCV

## Objetivos da Aula

Ao final desta aula você será capaz de:

- Entender por que um banco de dados é dividido em várias tabelas;
- Utilizar a função **PROCV** para relacionar informações entre tabelas;
- Calcular automaticamente o valor das compras;
- Produzir relatórios utilizando **CONT.SE**, **SOMASE**, **CONT.SES** e **SOMASES**.

---

# Situação-Problema

Imagine que você trabalha em uma empresa especializada na venda de smartphones.

Todos os dias centenas de pedidos são realizados.

A empresa poderia armazenar todas as informações em uma única tabela, repetindo o nome do produto, o preço, o gênero do cliente, sua renda e estado em todas as vendas.

Porém isso geraria milhares de informações repetidas.

Para resolver esse problema, a empresa dividiu seu banco de dados em **três tabelas**.

```text
                 PRODUTOS

ID Produto
Produto
Valor Unitário

----------------------------

                 CLIENTES

ID Cliente
Nome
Data de Nascimento
Gênero
Renda
Estado

----------------------------

                  VENDAS

ID Venda
Data
ID Produto
Quantidade
Canal de Venda
ID Cliente
```

---

# Por que separar as tabelas?

Imagine que o **Samsung S25 Ultra** seja vendido 50.000 vezes.

Se o nome do aparelho estivesse escrito em todas as linhas da planilha, haveria milhares de informações repetidas.

O mesmo aconteceria com o preço do produto, o estado do cliente e sua renda.

Separando as informações em tabelas diferentes, cada dado é armazenado apenas uma vez.

Caso o preço do produto seja alterado, basta modificar um único registro na tabela Produtos.

Esse conceito recebe o nome de **normalização de dados**, sendo utilizado em praticamente todos os bancos de dados modernos.

---

# Estrutura do Banco de Dados

## Tabela PRODUTOS

Esta tabela armazena todas as informações dos produtos comercializados.

| ID_Produto | Produto | Valor Unitário | Unidades Vendidas |
|------------|---------------------------|---------------:|------------------:|
|10001|Samsung S25|R$ 3.990,00||
|10002|Samsung S25 Plus|R$ 4.990,00||
|10003|Samsung S25 Ultra|R$ 5.990,00||
|10004|iPhone 17|R$ 6.990,00||
|10005|iPhone 17 Pro Max|R$ 11.990,00||
|10006|Vivo X200 Ultra|R$ 9.990,00||
|10007|Motorola Edge 60 Pro|R$ 3.990,00||
|10008|Xiaomi 16 Ultra|R$ 8.990,00||

---

## Tabela CLIENTES

Nesta tabela ficam armazenadas todas as informações dos clientes.

| ID_Cliente | Nome | Data de Nascimento | Gênero | Renda Mensal | Estado de Residência |
|------------|----------------|-----------------|-----------|--------------:|----------------------|
|1000|João da Silva|26/03/1971|Masculino|R$ 3.200,00|SP|
|1001|Maria Oliveira|09/07/1996|Feminino|R$ 4.100,50|SC|
|1002|Pedro Souza|19/03/1973|Masculino|R$ 2.980,75|PR|

---

## Tabela VENDAS

A tabela de vendas registra apenas as informações referentes à venda.

Observe que ela **não possui** o nome do produto nem as informações do cliente.

| ID Venda | Data | ID Produto | Quantidade | Canal de Venda | ID Cliente |
|----------:|------|-----------:|------------:|----------------|-----------:|

Essas informações serão obtidas automaticamente utilizando o **PROCV**.

---

# Como as tabelas se relacionam?

```text
                  TABELA PRODUTOS

             ID Produto
             Produto
             Valor Unitário

                    ▲
                    │
                PROCV(ID Produto)
                    │
                    │

             TABELA DE VENDAS

ID Venda
Data
ID Produto
Quantidade
Canal de Venda
ID Cliente
Valor Unitário
Valor da Compra
Gênero
Estado
Renda

                    │
                    │
                PROCV(ID Cliente)
                    │
                    ▼

                 TABELA CLIENTES

ID Cliente
Nome
Nascimento
Gênero
Renda
Estado
```

Observe que a tabela de vendas funciona como uma tabela central, recebendo informações das outras duas tabelas.

---

# Conhecendo a Função PROCV

A função PROCV procura um valor em uma tabela e devolve uma informação relacionada a esse valor.

## Sintaxe

```excel
=PROCV(valor_procurado;matriz_tabela;número_da_coluna;FALSO)
```

---

## Entendendo cada parâmetro

### Valor Procurado

É o código que desejamos localizar.

Exemplo:

```excel
C2
```

---

### Matriz da Tabela

É a tabela onde a pesquisa será realizada.

Exemplo:

```excel
PRODUTOS
```

ou

```excel
A2:D100
```

---

### Número da Coluna

Corresponde à coluna que será retornada.

Na tabela Produtos temos:

| Coluna | Informação |
|--------:|------------|
|1|ID Produto|
|2|Produto|
|3|Valor Unitário|
|4|Unidades Vendidas|

Se desejarmos retornar o preço do produto:

```excel
3
```

---

### Procurar Exatamente

Sempre utilizaremos:

```excel
FALSO
```

ou

```excel
0
```

Assim o Excel retornará apenas resultados exatamente iguais ao código informado.

---

# Exemplo 1 – Buscar o Valor Unitário

Queremos descobrir o preço do produto vendido.

```excel
=PROCV(C2;PRODUTOS;3;FALSO)
```

Leitura da fórmula:

> Procure o código existente na célula C2 dentro da tabela PRODUTOS e devolva a terceira coluna.

Resultado:

```
R$ 5.990,00
```

---

# Exemplo 2 – Buscar o Gênero

Agora queremos descobrir o gênero do cliente.

```excel
=PROCV(F2;CLIENTES;4;FALSO)
```

Resultado:

```
Masculino
```

---

# Exemplo 3 – Buscar a Renda

```excel
=PROCV(F2;CLIENTES;5;FALSO)
```

Resultado:

```
R$ 3.200,00
```

---

# Exemplo 4 – Buscar o Estado

```excel
=PROCV(F2;CLIENTES;6;FALSO)
```

Resultado:

```
SP
```

---

# Observe a transformação da tabela

Antes:

| ID Produto | ID Cliente |
|------------|------------|
|10003|1000|

Depois do PROCV:

| ID Produto | Valor Unitário | ID Cliente | Gênero | Estado | Renda |
|------------|---------------:|------------|---------|---------|-------:|
|10003|R$ 5.990,00|1000|Masculino|SP|R$ 3.200,00|

Sem digitar nenhuma dessas informações manualmente.

---

# Calculando o Valor da Compra

Agora basta multiplicar o preço do produto pela quantidade vendida.

```excel
=G2*D2
```

ou

```excel
=Valor Unitário × Quantidade
```

---

# Agora podemos criar relatórios

Como todas as informações estão reunidas em uma única tabela, podemos utilizar diversas funções estatísticas.

---

# Quantidade de vendas por Estado

Utilizamos a função **CONT.SE**.

```excel
=CONT.SE(J:J;M3)
```

Essa fórmula conta quantas vendas aconteceram em cada estado.

---

# Quantidade de vendas por Canal

```excel
=CONT.SE(E:E;Q3)
```

Conta quantas vendas ocorreram em cada canal de venda.

---

# Valor vendido por Estado

Agora queremos descobrir quanto cada estado comprou.

```excel
=SOMASE(J:J;U3;H:H)
```

Essa função soma o valor de todas as compras realizadas naquele estado.

---

# Valor vendido por Estado e Gênero

Agora temos dois critérios.

- Estado
- Gênero

Para isso utilizamos **SOMASES**.

Clientes masculinos:

```excel
=SOMASES(H:H;J:J;Y3;I:I;"Masculino")
```

Clientes femininos:

```excel
=SOMASES(H:H;J:J;Y3;I:I;"Feminino")
```

---

# Resultado Final

Após concluir a planilha será possível responder perguntas como:

- Qual estado realizou mais compras?
- Qual estado movimentou mais dinheiro?
- Qual canal realizou mais vendas?
- Quanto foi vendido para homens?
- Quanto foi vendido para mulheres?
- Qual produto possui maior faturamento?

---

# Exercícios

## Exercício 1

Utilize o PROCV para preencher:

- Valor Unitário
- Gênero
- Estado
- Renda Mensal

---

## Exercício 2

Calcule o Valor da Compra.

---

## Exercício 3

Crie um relatório utilizando **CONT.SE** mostrando a quantidade de vendas por estado.

---

## Exercício 4

Crie um relatório utilizando **CONT.SE** mostrando a quantidade de vendas por canal.

---

## Exercício 5

Utilize **SOMASE** para calcular o faturamento de cada estado.

---

## Exercício 6

Utilize **SOMASES** para calcular:

- Valor vendido para clientes masculinos em cada estado.
- Valor vendido para clientes femininos em cada estado.

---

# Conclusão

Nesta aula aprendemos que bancos de dados profissionais não armazenam todas as informações em uma única tabela. Em vez disso, os dados são organizados em tabelas específicas, relacionadas por um identificador único (ID). A função **PROCV** permite trazer automaticamente essas informações para a tabela principal de vendas, evitando repetição de dados e facilitando a manutenção do banco.

Depois que os dados são reunidos, funções como **CONT.SE**, **SOMASE**, **CONT.SES** e **SOMASES** permitem construir relatórios e indicadores de maneira rápida e eficiente.

Essa mesma lógica será utilizada futuramente em bancos de dados relacionais, Microsoft Access, SQL Server e Power BI.
