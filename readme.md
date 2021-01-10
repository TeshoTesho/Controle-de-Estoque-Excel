# Controle de Estoque

Todo o projeto foi dezenvolvido por projeto pessoal. O uso é livre para estudos. Qualquer comercialização é repudiada.]: 




------------


**Sistema de Estoque desenvolvido em EXCEL com auxílio do DOS(batch file).**

O sistema é pensado de forma simples, dentro do excel existe as planilhas de controle de estoque, são elas:
-  Inicio: Tela inicial.
-  Aviso: Tela pedindo para habilitar o uso de macros (Segurança)
- Estoque: Pagina que faz a soma a subtração da entrada e saida de mercadorias, possui os campos de balanço.
- Baixas: Pagina que faz os controles das baixas mensalmente de todos os setores.
- Entradas: Pagina que faz os controles de entradas mensalmente de todos os setores.
- Configurações: Pagina para acessar configurações referentes aos calculos que são exibidos dentro da planilha, como controle de dias.

***ESSAS SÃO AS PLANILHAS BASE.***

Irei explicar detalhadamente cada uma das seções e suas funções.

## PLANILHA

A planilha foi inicialmente desenvolvida de forma crua, com formulas de soma e subtração simples. Porém ocasionava problemas. 
A situação da planilha atual é: Você só pode dar baixa ou entrada até o dia atual, tudo o que for posto posteriormente não será contado até que o dia chegue.

Na planilha á também guias ocultas e desativadas, vamos ver brevemente sobre elas:

### Guias Desativadas:

- **Qtd Pessoas: **A planilha foi desenvolvida para uma Colonia de Ferias, onde há refeições de funcionários e hóspedes. Nesta página é adicionado o numero de pessoas em relação aos dias. Há uma Sub página de Funcionários que calcula quantos funcionários estão presentes em cada dia do mês. A planilha ***"Qtd Pessoas"*** pega e ultiliza o valor para a contagem. São separados em 3 refeições, na aba de configurações é possivel contar somente funcionários, somente hóspedes e ambos, assim como é possível descontar a refeição: Café da Manhã. Esta planilha é ultilizada para o calculo de consumo percapto.

- **Previsão:** Essa planilha foi desativada. As formulas pararam de funcionar, em breve atualizações. A pagina baseia-se no consumo percapto e na quantidade de pessoas marcadas para datas futuras, tentando adivinhar a média de gasto para aquele dia. 

- **Estoque Mínimo:** Essa planilha foi desativada. As formulas pararam de funcionar, em breve atualizações. A pagina baseias-se no gasto e número de baixas dos produtos, junto ao consumo percapto ele calcula o mínimo de quantidade de cada produto precisa-se para a quantidade de dias definidos. 

------------

###  Guias Escondias:

- **.** : Pagina responsável pela configuração como de acesso aos macros e algumas funções.
- **Cadastro de produtos: **A pagina coma  função de armazenar cada produto e seu setor para clonar nas outras planilhas. 



### Macro das planilhas 

**A planilha é inteiramente baseada no uso do VBA do Excel. 
**

**Funções:**
- **Histórico:** Armazena cada alteração, mostrando data, hora, guia, coluna, linha, valor e usuário. Indica também quando o usuário abre  e fecha a planilha. 

- Abrir planilha: Ao abrir a planilha á diversas etapas:
1. Se for a primeira vez que a planilha será aberta, ele irá puxar o estoque final da planilha anterior. Baseia-se na formulação dos nomes dos arquivos, **NÃO ALTERE**.
2. Ligar Histórico, mensagens de boas vindas e de configurações caso seja ADMIN(Dentro da planilha)
3. Ocultar e exibir páginas. Ao fechar são execultadas novamente, escondento todas e deixando somente a guia **"Aviso"** visível. Caso o ADMIN defina, ele transformará o excel em tela cheia.
4. irá cadastrar na guia **"Histórico"** quem abriu a planilha.

> OBS: Todas as funções só serão execultadas caso a planilha seja do mês atual

- Fechar Planilha: Irá cadastrar na guia **"Historico"** quem fechou a planilha e retornará as configurações originais do EXCEL caso houvesem sido alteradas.

### FORMULÁRIOS

Existem 2 formulários na planilha.

1. Com a função de resetar e limpar a planilha

2. Com funções de limpeza e configurações como:
2.1. Exibir mensagens.
2.2. Ligar e desligar Histórico.
2.3. Permitir seleção dentro da planilha.
2.4. Abrir form com a planilha
2.5. Definir ADMIN
2.6. Definir senhas (Todas as guias são protegidas com esta senha automaticamente)
2.7. Desbloquear e bloquear planilas
2.8. Salvar planilha ao fechar
2.9 Outras...


> A planilha é inteiramente desenvolvida para um único mês, para resolver isso, foi desenvolvido um sistema com arquivo .bat para criar planilhas de vários meses.

## Batch File

O arquivo bat chamado: Abrir Planilha é projetado para ser execultado e criar os diretórios e os arquivos expecíficos. 

São 5 etapas:
1. Caso não exista criar a pasta "Histórico"
2. Caso não exista criar a planilha do mês atual
2.1. A planilha é criada com uma regra de nome, sendo ela:
"(Numero do mês) (Nome do mês) - (ano) - Historico de Balanco"
> **Exemplo de uma planilha de dezembro de 2020:** 
> "12 Dezembro - 2020 - Historico de Balanco"

3. Caso não exista criar execultavel da planilha.
>  O excultavel é criado com um script e compilado em C pelo Framework64 na versão**: v2.0.50727 **
>  A função dele é execultar uma planilha excel que force macros automaticamente, sem pedir a permissão do usuário.
4. Esconder arquivos desnecessários como o execultavel 
5. Abrir a planilha

------------
