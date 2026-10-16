# sei-programacoes

Coletor **somente leitura** das "Programações de Contratação" do módulo Contratações do SEI.
Abre o Edge/Chrome visível, você faz o login (senha/MFA) e deixa a lista na tela; o programa lê a tela
"Consultar" de cada programação e grava uma planilha (`.xlsx`, `.xls` ou `.json`).

- Nada é alterado no SEI; nenhuma senha ou cookie é guardado em disco.
- Cada execução grava `sei_programacoes.log` na pasta de destino.
- Interface gráfica (`sei_programacoes.py`) e versão de linha de comando (`coletar_programacoes.py`).

## Uso
1. Informe a URL do SEI (ex.: `https://sei.exemplo.gov.br/sei/`), o formato e a pasta; clique em **Abrir SEI**.
2. Faça o login, abra Contratações > Programações de Contratação, filtre e clique em Pesquisar.
3. Com a lista na tela, clique em **Continuar**. O campo *Limite* (opcional) serve para testes.

## Executável (Windows)
Com Python 3.10+ no PATH, rode `gerar-exe.bat` → `dist\SEI-Programacoes.exe`
(embute o Playwright e usa o Edge instalado).

## Execução pelo fonte
```
pip install playwright openpyxl xlwt
python sei_programacoes.py
```

Os seletores foram ajustados para o SEI do CJF; outros órgãos podem precisar de adaptação.
Desenvolvido pela STI/CJF. Licença MIT.
