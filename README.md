# KPI Sync — Central PMOC

Roda `kpi_manutencao.py` a cada 5 minutos via GitHub Actions, varrendo os
relatórios PDF do Nextcloud e publicando no Supabase (tabelas usadas pela
Central PMOC do app de orçamento). Sem credenciais no código — tudo vem de
Secrets do repositório.

## Configurar (uma vez só)

Em **Settings → Secrets and variables → Actions → New repository secret**,
cadastre:

- `NEXTCLOUD_URL`
- `NEXTCLOUD_USER`
- `NEXTCLOUD_APP_PASSWORD`
- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`

Depois disso o workflow (`.github/workflows/sync.yml`) roda sozinho, sem
nenhuma ação manual. Pra rodar uma vez fora do horário: aba **Actions** →
"Sincronizar KPI de Manutenção" → **Run workflow**.
