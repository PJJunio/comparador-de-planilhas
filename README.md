# Ideia inicial

Enquanto eu prestava serviços para uma empresa, me deparei com uma tarefa simples e repetitiva, mas que consumia muito tempo,
que consistia em **verificar 2 planilhas de relatórios**, 1 da nossa empresa e outro da terceirizada, nessa tarefa era necessário
**procurar por erros dentro dos dados fornecidos pelos relatórios**, e caso houvesse algum, resolver.

Como as planilhas eram muito grandes, decidi criar um algorítimo simples para **comparar as células** da primeira coluna de ambas as planilhas
e se **caso fossem iguais, pintar elas de amarelo**.

Utilizei a **biblioteca openpyxl** para gerenciar minhas planilhas, fazendo a edição, comparação, e exibição das células desejadas.

O algorítimo está simples no momento, mas irei implementar mais funcionalidades em breve por conta de ações necessários, como comparar valores.

# Como usar

Até o momento, quando o programa é executado, vão abrir 2 janelas solicitando que o usuário selecione as planilhas.
Nos meus testes, usei a seguinte formatação
 * **Coluna A: Dado a ser comparado**
 * **Coluna E: Outro dado a ser comparado**
 * **Colunas B, C e D: Dados que não serão comparados**

<table>
  <thead>
    <tr>
      <th align="left">A</th>
      <th align="left">B</th>
      <th align="center">C</th>
      <th align="center">D</th>
      <th align="right">E</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td align="left">Joao</td>
      <td align="left">Hemograma Completo</td>
      <td align="center">Periódico</td>
      <td align="center">01/01/2025</td>
      <td align="right">R$ 9,00</td>
    </tr>
    <tr>
      <td align="left">Igor</td>
      <td align="left">Glicemia</td>
      <td align="center">Periódico</td>
      <td align="center">01/01/2025</td>
      <td align="right">R$ 6,30</td>
    </tr>
    <tr>
      <td align="left">Paulo</td>
      <td align="left">Hemograma Completo</td>
      <td align="center">Periódico</td>
      <td align="center">01/01/2025</td>
      <td align="right">R$ 9,00</td>
    </tr>
    <tr>
      <td align="left">Eduardo</td>
      <td align="left">Glicemia</td>
      <td align="center">Periódico</td>
      <td align="center">01/01/2025</td>
      <td align="right">R$ 6,30</td>
    </tr>
  </tbody>
</table>

As 2 planilhas estão com dados diferentes e estão formatadas nesse estilo, nesse caso ela reescrever a planilha um do lado da outra e comparar o nome e o valor de A1 e E1 com os da outra planilha
e se caso forem iguais, pinta os valores de amarelo.

Até o momento, caso queira modificar quais dados serão comparados, ainda é necessário digitar manualmente no If da linha 90.

* **if valor1_1 == valor2_1 and valor1_5 == valor2_5: = Compara as cêlulas das colunas A e E.**
* **if valor1_2 == valor2_2 and valor1_3== valor2_3: = Compara as cêlulas das colunas B e C.**
* **if valor1_1 == valor2_1: Compara as cêlulas apenas da coluna A.**
* **if valor1_2 == valor2_2: Compara as cêlulas apenas da coluna B.**

Em breve estarei postando uma solução para esse problema, adicionando a possibilidade de escolher quais colunas vão ser comparadas e tavez um script para organizar os dados automaticamente.
