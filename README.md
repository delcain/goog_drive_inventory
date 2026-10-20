# Inventário dos Drives Compartilhados

## Configuração
1. Crie projeto no Google Cloud e ative **Google Drive API**.
2. Crie uma Service Account e habilite **Domain-Wide Delegation**. Guarde a chave JSON localmente.
3. No Admin Console: Segurança > Controle de acesso e dados > Controles de API > Delegação em todo o domínio, adicione o **Client ID numérico** da conta de serviço com escopo `https://www.googleapis.com/auth/drive.readonly`.
4. Use um superadministrador como `--admin`. O inventário depende da visibilidade/autorização da conta e dos controles da organização.
5. Execute em Python 3.11+:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python inventory.py --key service-account.json --admin administrador@seudominio.com
```

Relatórios: `relatorios/resumo.csv`, `relatorios/arquivos.csv`, `relatorios/resumo.xlsx`.

## Limitações e cuidados
- A API pode exigir permissões adicionais para listar arquivos em certos Drives; cada falha é registrada em `drive_scans`, e a execução fica `partial`.
- O script não inclui arquivos na lixeira nem itens de Meu Drive.
- O tamanho dos arquivos nativos do Google frequentemente não está disponível via `size`; o total de bytes é uma **estimativa parcial**, não a quota oficial faturada.
- O relatório de arquivos inclui o ID dos pais, não o caminho hierárquico reconstruído.
- Inventários são snapshots operacionais, não auditoria histórica exata de eventos; exclusões só são refletidas em Drives varridos integralmente.
- Se um Drive deixar de ser descoberto, seus registros anteriores permanecem no SQLite. Verifique `runs` e `drive_scans` antes de interpretar totais.
- Proteja a chave JSON e o SQLite. Nunca inclua a chave em repositórios Git.
