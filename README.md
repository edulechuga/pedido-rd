# Automação de Pedidos RD

Este é um projeto desenvolvido para automatizar a geração de planilhas de pedidos através da extração de dados de uma planilha geral para um *template* pré-configurado.

## Visão Geral

O script Python principal lerá a planilha que contém todas as aberturas de pedidos (ex: `Pedidos Aberturas - CUSTOM - 05.12.xlsx`) e, para cada linha validada, preencherá as células correspondentes em um *template* de formulário (ex: `RQCM-001 - Pedido de venda - V05.xlsx`).

Os dados transpostos incluem:
- Observações (inserido na célula `B17`)
- Valores Unitários (`A27`, `A42`)
- Valores RD e de Custo HW (`D27`, `D42`)
- PTAX (`B47`)

As fórmulas de valorização final de cada item (`A27 * D27`, etc) são aplicadas também dinamicamente. Os arquivos processados são higienizados em relação a caracteres especiais e salvos na subpasta `pedidos_gerados_excel`.

## Estrutura do Repositório

- `gerar_pedidos_final.py`: Fonte de código principal e versão estável da automação.
- `.gitignore`: Evita o monitoramento pelo Git de arquivos brutos (`*.zip`, `*.xlsx`), do cache python (`__pycache__`) e das subpastas dinâmicas de saída geradas pelo script.
- `README.md`: Este documento sobre a estrutura e os processos do projeto.

---

### Executando o Script 

Certifique-se de ter as bibliotecas instaladas (como o `pandas` e o `openpyxl`).

```bash
pip install pandas openpyxl
python gerar_pedidos_final.py
```

### Guia Rápido de Git

Sempre que concluir atualizações no código, utilize os seguintes comandos no terminal:

1. \`git add .\`
2. \`git commit -m "Explicação breve do que foi modificado"\`
3. \`git push\`
