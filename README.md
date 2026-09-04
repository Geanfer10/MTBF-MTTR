# Painel PCM — MTTR / MTBF (Goiás Verde Alimentos)

Painel de indicadores de manutenção hospedado no GitHub Pages, atualizado
automaticamente sempre que uma planilha nova é enviada para a pasta `uploads/`.

## Como funciona

```
uploads/PCM_....xlsm   →  (você sobe o arquivo)
        │
        ▼
GitHub Actions roda scripts/parse_excel.py
        │
        ▼
data/2026-09-04.json, data/latest.json, data/history.json  (gerados sozinhos)
        │
        ▼
index.html (GitHub Pages) lê os JSON e desenha o painel
```

Você não edita HTML nem JSON manualmente. O único passo manual é subir o
arquivo `.xlsm` do dia.

## Passo a passo — configuração inicial (só uma vez)

1. **Criar o repositório**
   No GitHub, clique em *New repository*. Pode ser público ou privado (o
   GitHub Pages funciona nos dois casos em contas Pro/Team; em contas free,
   Pages só publica repositórios públicos).

2. **Subir estes arquivos**
   Envie toda esta pasta (`index.html`, `scripts/`, `.github/`, `data/`,
   `uploads/`, `README.md`) para o repositório — pode arrastar e soltar
   direto na interface do GitHub ("Add file → Upload files").

3. **Ativar o GitHub Pages**
   Vá em *Settings → Pages*. Em "Source", selecione a branch `main` e a
   pasta `/ (root)`. Salve. Em alguns minutos o GitHub te dá um link tipo:
   `https://SEU-USUARIO.github.io/NOME-DO-REPO/`

4. **Conferir as permissões do Actions**
   Vá em *Settings → Actions → General → Workflow permissions* e marque
   **"Read and write permissions"**. Isso é necessário para o robô conseguir
   salvar os arquivos de dados de volta no repositório.

Pronto — configuração feita uma única vez.

## Uso no dia a dia

1. Abra a pasta `uploads/` no seu repositório.
2. Clique em **Add file → Upload files**.
3. Envie o `.xlsm` atualizado (pode manter sempre o mesmo nome ou trocar,
   tanto faz).
4. Clique em *Commit changes*.
5. Aguarde 1–2 minutos — o robô roda sozinho (acompanhe em **Actions**, na
   aba do topo do repositório).
6. Abra o link do painel (passo 3 da configuração) — os dados já estarão
   atualizados, e a data mais recente aparece automaticamente selecionada
   no seletor de data do topo.

## Estrutura de pastas

```
index.html                      → o painel (não precisa editar)
scripts/parse_excel.py          → script que lê o Excel e gera os JSON
.github/workflows/process-excel.yml → o "robô" (GitHub Actions)
uploads/                        → onde você sobe o .xlsm de cada dia
data/                           → gerado automaticamente (não editar)
  ├── latest.json               → dados do dia mais recente
  ├── history.json              → série histórica (gráfico de tendência)
  └── AAAA-MM-DD.json           → um arquivo por dia processado
```

## Rodando localmente (opcional, para testar antes de subir)

Se quiser gerar os JSON no seu computador antes de subir pro GitHub:

```bash
pip install pandas openpyxl
python scripts/parse_excel.py --file "caminho/para/PCM_Indicadores.xlsm"
```

Isso cria os arquivos em `data/`. Depois é só subir a pasta `data/`
atualizada junto com o commit.

## Personalizações possíveis (me chame quando quiser mexer nisso)

- **KPI "Dias Parados"**: hoje é um valor manual (`--dias-parados`, padrão
  0) porque a definição exata desse indicador não está na planilha bruta.
  Se você me disser a regra certa, eu ajusto o script para calcular sozinho.
- **Meta MTBF mensal**: também é um valor de configuração (`--meta-mtbf`,
  padrão 300). Se quiser trocar por período, dá pra automatizar também.
- **Cores, seções, novos gráficos**: qualquer ajuste visual continua sendo
  só me pedir — a estrutura de dados já fica pronta para novos cruzamentos.
