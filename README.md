# aniversario-reframax

Gerador de listas de aniversariantes com GUI (PyQt6), layouts configuráveis, preferências salvas, bandeja do sistema e geração automática.

## Requisitos

- Python 3.10+
- PyQt6, Pillow, pandas, openpyxl, pywin32

## Uso (desenvolvimento)

```powershell
py script_gui.py
```

## Compilar

```powershell
pyinstaller Aniversariantes.spec
```

O executável compilado consulta `server.json` no GitHub. Se o campo `status` não for `"ativo"`, o app não inicia.

## Estrutura

- `planilhas/` — arquivos Excel (.xlsx) com colunas `Nome` e `Nascimento`
- `layouts/` — layouts com `layout.json` e imagem de fundo
- `server.json` — controle remoto de licença (`status: ativo`)
