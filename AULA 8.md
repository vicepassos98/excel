# Aula 08 – Funções SE e SES

## Objetivos da Aula

Ao final desta aula você será capaz de:

- Utilizar a função **SE** para tomar decisões no Excel;
- Compreender quando utilizar **SE aninhado**;
- Conhecer a função **SES**, uma alternativa mais organizada para vários testes lógicos;
- Automatizar cálculos de custos e valores de frete.

---

# O que é uma função lógica?

As funções lógicas permitem que o Excel **tome decisões automaticamente**.

Em vez de o usuário escolher manualmente um resultado, o Excel verifica uma condição e retorna uma resposta.

Exemplo:

> Se a média do aluno for maior ou igual a 6, então ele está aprovado.

Caso contrário:

> Está reprovado.

---

# Função SE

A função **SE** executa um teste lógico.

Se o teste for verdadeiro, o Excel retorna um resultado.

Caso contrário, retorna outro.

## Sintaxe

```excel
=SE(Teste Lógico;Valor se Verdadeiro;Valor se Falso)
```

---

## Exemplo 1

```excel
=SE(A2>=60;"Aprovado";"Reprovado")
```

Leitura:

> Se A2 for maior ou igual a 60, escreva **Aprovado**.

Caso contrário:

> Escreva **Reprovado**.

---

## Exemplo 2

```excel
=SE(B2="SP";"Sudeste";"Outro Estado")
```

---

# Quando o SE não é suficiente?

Imagine que precisamos classificar uma nota.

- Menor que 5 → Insuficiente
- Entre 5 e 6,9 → Regular
- Entre 7 e 8,9 → Bom
- Maior ou igual a 9 → Excelente

Nesse caso existem **quatro possibilidades**.

O SE tradicional possui apenas duas respostas.

Então precisamos utilizar um **SE dentro de outro SE**.

---

# SE Aninhado

```excel
=SE(A2<5;"Insuficiente";
SE(A2<7;"Regular";
SE(A2<9;"Bom";
"Excelente")))
```

Observe que cada novo teste fica dentro do resultado falso do teste anterior.

Embora funcione perfeitamente, a fórmula começa a ficar difícil de ler.

---

# Função SES

A função **SES** foi criada justamente para substituir vários SE aninhados.

Ela permite escrever diversos testes de forma muito mais organizada.

## Sintaxe

```excel
=SES(
Teste1;Resultado1;
Teste2;Resultado2;
Teste3;Resultado3;
...)
```

---

## O mesmo exemplo utilizando SES

```excel
=SES(
A2<5;"Insuficiente";
A2<7;"Regular";
A2<9;"Bom";
A2>=9;"Excelente")
```

Muito mais fácil de compreender.

---

# Comparando as funções

## Utilizando SE

```excel
=SE(A2<5;"Insuficiente";
SE(A2<7;"Regular";
SE(A2<9;"Bom";
"Excelente")))
```

---

## Utilizando SES

```excel
=SES(
A2<5;"Insuficiente";
A2<7;"Regular";
A2<9;"Bom";
A2>=9;"Excelente")
```

O resultado é exatamente o mesmo.

A diferença é que a função **SES** deixa a fórmula muito mais limpa.

---

# Exercício Prático

Nesta aula iremos automatizar os cálculos de uma empresa de transportes.

Nossa planilha registra diariamente as viagens realizadas pelos caminhões.

---

# Banco de Dados

| Data da Saída | Código | Cidade de Origem | Cidade de Destino | Motorista | Placa do Veículo | Tipo de Carga | Peso da Carga (kg) | Distância (km) | Custo do Frete | Custo Mão de Obra | Valor Cobrado ao Cliente |
|---------------|---------|------------------|-------------------|------------|------------------|---------------|-------------------:|---------------:|----------------:|-------------------:|--------------------------:|
|01/08/2025|10259001|Cabréuva - SP|Paulista - PE|Valéria Paschoal|ARM-2703|Perigosa|4833|2680|=3,1*I2+0,8*H2|=SE(G2="Comum";1,4*I2;SE(G2="Perigosa";2,1*I2;3,5*I2))|=(K2+J2)*125%|
|01/08/2025|10259002|Paulista - PE|Cabréuva - SP|Hilton Jesus|ARM-2765|Comum|4869|2680|R$ 12.203,20|R$ 3.752,00|R$ 19.944,00|
|02/08/2025|10259003|Uruguaiana - RS|Cabréuva - SP|Lupércia Fagundes|DEF-2025|Comum|4204|1570|R$ 8.230,20|R$ 2.198,00|R$ 13.035,25|
|02/08/2025|10259004|Queimados - RJ|Cabréuva - SP|Bino Correia|DXW-7047|Comum|1314|480|R$ 2.539,20|R$ 672,00|R$ 4.014,00|
|03/08/2025|10259005|Cabréuva - SP|Passa e Fica - RN|Bino Correia|DXW-7047|Comum|516|2760|R$ 8.968,80|R$ 3.864,00|R$ 16.041,00|

---

# Etapa 1 – Calculando o custo do frete

O custo do combustível depende apenas da distância percorrida e do peso transportado.

A fórmula será:

```excel
=3,1*Distância + 0,8*Peso
```

No Excel:

```excel
=3,1*I2+0,8*H2
```

---

# Etapa 2 – Calculando a mão de obra

Agora entra a função **SE**.

Cada tipo de carga possui um custo diferente.

| Tipo de carga | Valor pago por km |
|----------------|------------------:|
| Comum | R$ 1,40 |
| Perigosa | R$ 2,10 |
| Especial | R$ 3,50 |

Precisamos que o Excel descubra automaticamente qual valor utilizar.

---

## Utilizando SE Aninhado

```excel
=SE(
G2="Comum";
1,4*I2;
SE(
G2="Perigosa";
2,1*I2;
3,5*I2))
```

Leitura:

- Se a carga for **Comum**, multiplique a distância por 1,4.
- Caso contrário, verifique se é **Perigosa**.
- Se for perigosa, multiplique por 2,1.
- Caso contrário, considere que é **Especial** e multiplique por 3,5.

---

# A mesma fórmula utilizando SES

```excel
=SES(
G2="Comum";1,4*I2;
G2="Perigosa";2,1*I2;
G2="Especial";3,5*I2)
```

Observe como a leitura ficou muito mais simples.

---

# Etapa 3 – Valor cobrado do cliente

A empresa trabalha com uma margem de lucro de **25%**.

Logo, o valor cobrado será:

```excel
=(J2+K2)*125%
```

Ou seja:

```
(Custo do Frete + Mão de Obra)
+
25% de lucro
```

---

# Desafio

A empresa criou um novo tipo de carga.

| Tipo | Valor por km |
|-------|-------------:|
| Refrigerada | R$ 2,80 |

Atualize as fórmulas para que esse novo tipo seja calculado automaticamente.

Faça isso:

- utilizando **SE aninhado**;
- utilizando **SES**.

---

# Exercícios

### Exercício 1

Calcule o custo do frete utilizando a fórmula apresentada.

---

### Exercício 2

Calcule a mão de obra utilizando a função **SE**.

---

### Exercício 3

Reescreva a fórmula utilizando **SES**.

---

### Exercício 4

Calcule o valor cobrado do cliente.

---

### Exercício 5

Inclua o tipo de carga **Refrigerada** nas duas fórmulas.

---

# Conclusão

Nesta aula aprendemos duas importantes funções lógicas do Excel.

A função **SE** é ideal para decisões simples, envolvendo apenas duas possibilidades. Quando o número de condições aumenta, é possível utilizar o **SE aninhado**, inserindo uma função dentro da outra. Entretanto, fórmulas muito grandes tornam-se difíceis de interpretar e manter.

Para resolver esse problema, o Excel oferece a função **SES**, que permite testar diversas condições de forma sequencial, deixando a fórmula mais organizada, legível e fácil de modificar. Em situações com vários critérios, a função **SES** costuma ser a melhor escolha.
