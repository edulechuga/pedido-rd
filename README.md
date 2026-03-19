# Automação de Pedidos RD

Este é um projeto desenvolvido para automatizar a geração de planilhas de pedidos através da extração de dados de uma planilha geral para um *template* pré-configurado.

## Visão Geral

O script Python principal lerá a planilha que contém todas as aberturas de pedidos (ex: `Pedidos Aberturas - CUSTOM - 05.12.xlsx`) e, para cada linha validada, preencherá as células correspondentes em um *template* de formulário (ex: `RQCM-001 - Pedido de venda - V05.xlsx`).

### Extração Dinâmica de Colunas
A captura de informações na planilha de origem não está mais engessada em índices de colunas específicos. O script localiza os dados rastreando e combinando dinamicamente os nomes dos cabeçalhos, mesmo que sua ordem mude, tornando o código altamente resistente à inserção de novas colunas por erro humano ou novas requisições na planilha original.

Os dados que o algoritmo mapeia dinamicamente e que são transpostos, incluem:
- Observações (inserido na célula `B17`) a partir da coluna `'Observações NF'`
- Valores Unitários (`A27`, `A42`) a partir da coluna `'Qtde'`
- Valores de RD e Custo de Hardware (`D27`, `D42`) a partir das colunas `'VALOR RD - PDF'` e `'Custo UNT HW + Patrimonio'`
- PTAX (`B47`) a partir da coluna `'PTAX'`

O script ainda limpa quebras e formatações de colunas via `.strip()` e calcula fórmulas de valorização (`A27 * D27`) inserindo-as dinamicamente no destino. O arquivos resultantes e higienizados de caracteres especiais são salvos na pasta `pedidos_gerados_excel`.

## Estrutura do Repositório

- `gerar_pedidos_final.py`: Fonte de código principal e versão estável da automação.
- `.gitignore`: Evita o monitoramento pelo Git de arquivos brutos (`*.zip`, `*.xlsx`), do cache python (`__pycache__`) e das subpastas dinâmicas de saída geradas pelo script.
- `README.md`: Este documento sobre a estrutura e os processos do projeto.

---

### Executando o Script 

Certifique-se de ter as bibliotecas instaladas na sua máquina (como o `pandas` e o `openpyxl`).

```bash
pip install pandas openpyxl
python3 gerar_pedidos_final.py
```

### Guia Rápido de Git: Como criar atualizações seguras

Implementar melhorias e enviá-las para a produção (a *main*) da maneira correta exige que não alteremos diretamente a base de código primário. Para toda nova funcionalidade, é importante seguir este curto fluxo de ramificação (branches):

**1. Crie uma nova branch para proteger a main:**
No seu projeto atual na `main`, "clone-a" criando um ambiente paralelo:
```bash
git switch -c nome-da-sua-nova-melhoria-aqui
```
*(ex: `git switch -c colunas-dinamicas`)*

**2. Faça e Salve as alterações na nova branch:**
Realize todo o seu código normalmente. Quando acabar e as novidades funcionarem, empacote-as na nuvem da ramificação provisória:
```bash
git add .
git commit -m "Explicação breve do que foi modificado"
```

**3. Volte para a branch Main:**
```bash
git switch main
```

**4. Mescle os códigos e faça o Merge:**
Ao retornar à main, traga aquelas linhas que você acabou de comitar:
```bash
git merge nome-da-sua-nova-melhoria-aqui
```

**5. Envie a integração oficial ao Github:**
A sua main agora está atualizada na sua máquina! Por fim, sincronize com a internet, delete a ramificação temporária caso já não use mais e o fluxo está encerrado com êxito e organização.
```bash
git push
git branch -d nome-da-sua-nova-melhoria-aqui
```
