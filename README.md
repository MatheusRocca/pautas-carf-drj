# Pautas CARF/DRJ

Ferramenta em Python que lê o Diário Oficial da União, encontra as **pautas de julgamento do CARF e das DRJs** (Receita Federal) e gera uma planilha Excel com cada processo pautado. Com uma lista de clientes, ela também aponta quais processos pertencem a eles e pode avisar por e-mail.

Desenvolvido com apoio de IA (Claude Code).

## O que faz

1. **Baixa o DOU** do dia (seções 1, 2 e 3 e edições extras) pelo [INLabs](https://inlabs.in.gov.br/), da Imprensa Nacional.
2. **Filtra as pautas**: lê os XMLs e mantém só os artigos do tipo "Pauta" do CARF e das DRJs.
3. **Extrai os dados** de cada processo: órgão, turma, número do processo, recorrente, interessado, data e horário do julgamento, relator e link para a publicação.
4. **Cruza com uma lista de clientes** (Excel ou Google Sheets), pelo número do processo ou pelo nome da parte, e explica por que cada linha bateu.
5. **Gera a planilha** `pautas_AAAA-MM-DD.xlsx` e, no modo automático, **envia um e-mail** com as pautas dos clientes.

## Como usar

Requer Python 3 (testado no 3.10 e no 3.12).

```bash
pip install -r requirements.txt

# processar XMLs/ZIPs que você já tem
python pautas_carf.py ./2026-09-25-DO1 saida.xlsx

# cruzar com uma lista de clientes
python pautas_carf.py ./2026-09-25-DO1 saida.xlsx --clientes clientes.xlsx

# baixar o DOU de hoje (ou de outra data) e processar
python pautas_carf.py --baixar
python pautas_carf.py --baixar 2026-09-25

# sem argumentos: abre um menu no terminal
python pautas_carf.py
```

Para baixar do INLabs, crie um arquivo `inlabs.ini` (já está no `.gitignore`):

```ini
[inlabs]
email = seu-email@exemplo.com
senha = sua-senha
```

## Modo automático

- **No Mac:** `python pautas_carf.py --instalar-agendamento` agenda a execução em dias úteis às 8h30, 13h e 18h (launchd).
- **No GitHub Actions:** o workflow `.github/workflows/pautas.yml` roda o mesmo processo na nuvem. As credenciais ficam nos *secrets* do repositório (lista no topo do arquivo) e o estado entre execuções fica no branch `estado`, para não repetir avisos.

## Testes

São 104 testes automatizados (`unittest`) usando XMLs reais do DOU em `tests/fixtures`. Eles rodam a cada push pelo workflow `testes.yml`.

```bash
python -m unittest discover -s tests
```

## Estrutura

```
pautas_carf.py            script principal (extração, download, cruzamento, avisos)
tests/                    testes e fixtures (XMLs públicos do DOU)
.github/workflows/        testes a cada push e modo automático
```
