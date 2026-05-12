# Como gerar o EXE sem instalar Python na sua maquina

Este caminho usa o GitHub Actions. O executavel e gerado em um Windows na nuvem e voce baixa apenas o arquivo pronto.

## Passo a passo

1. Crie um repositorio no GitHub.
2. Envie estes arquivos para o repositorio.
3. No GitHub, abra a aba `Actions`.
4. Clique em `Build Windows EXE`.
5. Clique em `Run workflow`.
6. Aguarde terminar.
7. Baixe o artefato `RegistroPontoManual-Windows`.

Dentro do artefato estara:

```text
RegistroPontoManual.exe
_internal\
```

Esse executavel nao exige Python instalado na maquina do usuario. Mantenha o `RegistroPontoManual.exe` na mesma pasta da pasta `_internal`.

## Observacao

O banco de dados, as configuracoes e os backups continuam ficando no perfil do usuario:

```text
%USERPROFILE%\RegistroPontoManual
```
