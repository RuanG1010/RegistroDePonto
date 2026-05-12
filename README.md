# Registro de Ponto Manual

Aplicativo local em Python + SQLite para controle paralelo de ponto, com fechamento do dia 16 ao dia 15, exportacao para Excel/PDF, feriados automaticos e backup manual.

## Rodar pelo Python

```bat
python -m pip install -r requirements.txt
python registro_ponto_manual.py
```

## Gerar executavel sem Python na sua maquina

A opcao recomendada para quem nao pode instalar Python e usar o GitHub Actions:

1. Envie esta pasta para um repositorio no GitHub.
2. Abra a aba `Actions`.
3. Rode o workflow `Build Windows EXE`.
4. Baixe o artefato `RegistroPontoManual-Windows`.

Veja o passo a passo em [COMO_GERAR_EXE_SEM_PYTHON.md](COMO_GERAR_EXE_SEM_PYTHON.md).

## Gerar executavel em uma maquina com Python

No Windows, execute:

```bat
build_exe.bat
```

O executavel sera gerado em:

```text
dist\RegistroPontoManual.exe
```

## Onde ficam os dados

O banco, as configuracoes e os backups ficam no perfil do usuario:

```text
%USERPROFILE%\RegistroPontoManual
```

Arquivos principais:

- `ponto_manual.db`: banco SQLite
- `config.json`: configuracoes
- `backups\`: backups manuais do banco

## Regra revisada

Por padrao, dia util sem lancamento entra como debito no fechamento. Datas futuras aparecem como `Aguardando` e nao entram como debito antes do dia chegar. Essa regra pode ser desligada em `Configuracoes`, no item:

```text
Dia util sem lancamento entra como debito no fechamento
```
