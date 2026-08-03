# CONT.SES e SOMASES

As funções **CONT.SE** e **SOMASE** permitem trabalhar com **apenas um critério** de pesquisa. Elas são ideais quando precisamos contar ou somar valores levando em consideração apenas uma condição.

### CONT.SE
Conta quantas células atendem a um único critério.

**Sintaxe:**
```excel
=CONT.SE(intervalo; critério)
```

**Exemplo:**
```excel
=CONT.SE(D:D;"Clóvis do Carmo")
```

Conta quantas propostas foram atendidas pelo corretor **Clóvis do Carmo**.

---

### SOMASE
Soma valores utilizando apenas um critério.

**Sintaxe:**
```excel
=SOMASE(intervalo_critério; critério; intervalo_soma)
```

**Exemplo:**
```excel
=SOMASE(D:D;"Clóvis do Carmo";F:F)
```

Soma o valor de todos os imóveis vendidos pelo corretor **Clóvis do Carmo**.

---

## Quando utilizar CONT.SES e SOMASES?

As funções **CONT.SES** e **SOMASES** funcionam de maneira semelhante às anteriores, porém permitem utilizar **dois ou mais critérios ao mesmo tempo**.

Isso é muito útil quando queremos fazer análises mais específicas.

### CONT.SES

Conta quantas células atendem a **vários critérios simultaneamente**.

**Sintaxe:**

```excel
=CONT.SES(intervalo1; critério1; intervalo2; critério2; ...)
```

**Exemplo:**

```excel
=CONT.SES(D:D;"Clóvis do Carmo";K:K;"APROVADO(A)")
```

Resultado:

> Conta apenas as propostas do corretor **Clóvis do Carmo** que foram **APROVADAS**.

---

### SOMASES

Soma valores considerando **dois ou mais critérios**.

**Sintaxe:**

```excel
=SOMASES(intervalo_soma; intervalo1; critério1; intervalo2; critério2; ...)
```

**Exemplo:**

```excel
=SOMASES(F:F;D:D;"Clóvis do Carmo";K:K;"APROVADO(A)")
```

Resultado:

> Soma apenas o valor dos imóveis vendidos por **Clóvis do Carmo** que tiveram a proposta **APROVADA**.

---

## Comparação das funções

| Função | Quantidade de critérios | Finalidade |
|---------|------------------------|------------|
| CONT.SE | 1 | Conta registros que atendem um critério |
| CONT.SES | 2 ou mais | Conta registros que atendem vários critérios |
| SOMASE | 1 | Soma valores utilizando um critério |
| SOMASES | 2 ou mais | Soma valores utilizando vários critérios |

---

# Resumo de Vendas por Corretor

| Corretor | Ofertas Recebidas (CONT.SE) | Valor das Ofertas Recebidas (SOMASE) | Comissão Esperada (4%) | Ofertas Aprovadas (CONT.SES) | Valor Aprovado (SOMASES) | Comissão Aprovada |
|----------|----------------------------:|-------------------------------------:|-----------------------:|-----------------------------:|-------------------------:|------------------:|
| Clóvis do Carmo | 7 | R$ 2.183.621,00 | R$ 87.344,84 | 6 | R$ 1.988.297,00 | R$ 79.531,88 |
| Neide Fontoura | 5 | R$ 1.917.437,00 | R$ 76.697,48 | 3 | R$ 1.091.425,00 | R$ 43.657,00 |
| Gilmar Carvalho | 5 | R$ 1.736.879,00 | R$ 69.475,16 | 4 | R$ 1.423.229,00 | R$ 56.929,16 |
| Jurema Sales | 9 | R$ 3.264.439,00 | R$ 130.577,56 | 7 | R$ 2.500.659,00 | R$ 100.026,36 |
| Osvaldo Bilac | 7 | R$ 2.713.366,00 | R$ 108.534,64 | 4 | R$ 1.493.318,00 | R$ 59.732,72 |
| **Total** | **33** | **R$ 11.815.742,00** | **R$ 472.629,68** | **24** | **R$ 8.496.928,00** | **R$ 339.877,12** |

---

# Resumo de Vendas por Empreendimento

| Empreendimento | Ofertas Recebidas | Valor das Ofertas Recebidas | Ofertas Aprovadas (CONT.SES) | Valor Aprovado (SOMASES) |
|----------------|------------------:|----------------------------:|-----------------------------:|-------------------------:|
| Lux Tower II | 7 | R$ 3.586.534,00 | 4 | R$ 2.049.448,00 |
| Portal do Paraíso | 7 | R$ 3.150.910,00 | 6 | R$ 2.700.780,00 |
| Torres de Papel | 5 | R$ 976.620,00 | 3 | R$ 585.972,00 |
| Vale dos Sonhos | 8 | R$ 2.509.200,00 | 5 | R$ 1.568.250,00 |
| Vista Alegre | 6 | R$ 1.592.478,00 | 6 | R$ 1.592.478,00 |

---

# Exercícios Propostos

1. Utilize **CONT.SE** para descobrir quantas propostas cada corretor recebeu.

2. Utilize **SOMASE** para calcular o valor total das propostas de cada corretor.

3. Utilize **CONT.SES** para descobrir quantas propostas **aprovadas** cada corretor possui.

4. Utilize **SOMASES** para calcular o valor total das propostas aprovadas de cada corretor.

5. Repita os mesmos cálculos para os empreendimentos, utilizando as funções **CONT.SES** e **SOMASES**.
