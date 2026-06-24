# Gestão360 Contábil

Plataforma web para controle de clientes de escritório contábil, atividades/serviços, contas a receber, dashboard e relatórios separados de atividades e recebimentos.

## Primeiro acesso e criação de contas

Não existe mais e-mail ou senha padrão.

Na primeira abertura, o sistema mostra a tela **Configuração inicial**. Nela, você cria:

- a primeira conta administradora;
- a senha pessoal do administrador;
- o código do escritório usado para autorizar novos cadastros.

Depois disso, cada membro da equipe pode clicar em **Criar minha conta** na tela de login e informar o código do escritório. Novas contas entram inicialmente como **Colaborador**.

O administrador pode liberar ou bloquear novos cadastros e trocar o código em **Configurações**.

## Como rodar no Windows

1. Instale o Python 3.11 ou superior.
2. Extraia o ZIP.
3. Dê dois cliques em `iniciar_windows.bat`.
4. Quando aparecer `Running on http://127.0.0.1:5000`, abra esse endereço no navegador.

## Como rodar no Mac/Linux

No terminal, entre na pasta do projeto e rode:

```bash
chmod +x iniciar_linux_mac.sh
./iniciar_linux_mac.sh
```

Depois acesse:

```text
http://127.0.0.1:5000
```

## Funcionalidades incluídas

- Login
- Dashboard principal
- Cadastro e lista de clientes
- Ficha completa do cliente
- Serviços / atividades por cliente
- Soma automática de horas por mês
- Contas a receber
- Total a receber, recebido, atraso e inadimplentes
- Relatório de atividades por cliente
- Relatório exclusivo das baixas/recebimentos do mês
- Cálculo de valor por hora: valor mensal cobrado ÷ horas gastas
- Classificação de rentabilidade
- Perfis de usuário: Administrador e Colaborador
- Bloqueio de dados financeiros para colaborador
- Exportação de relatórios em CSV, Excel e PDF
- Banco de dados SQLite local

## Publicação na Vercel

A Vercel executa o Flask normalmente, mas o banco SQLite local não é permanente no ambiente serverless. Esta versão aceita um banco **Turso/libSQL** em nuvem sem alterar as telas do sistema.

Antes de publicar, crie um banco no Turso e configure estas variáveis no projeto da Vercel:

```text
GESTAO360_SECRET=uma-chave-longa-e-aleatoria
TURSO_DATABASE_URL=libsql://...
TURSO_AUTH_TOKEN=...
```

O arquivo `.env.example` mostra os nomes corretos. Nunca envie o arquivo `.env` nem os tokens para o GitHub.

Quando o Turso estiver ativo:

- clientes, tarefas, cobranças, usuários e configurações ficam no banco em nuvem;
- anexos das atividades também são gravados no banco em nuvem;
- o backup local `.db` fica desativado e deve ser substituído pelos recursos de backup do provedor.

### Importante sobre os usuários

As contas criadas pelo botão **Criar minha conta** pertencem à mesma equipe e compartilham os dados do mesmo escritório. O código do escritório impede que visitantes desconhecidos criem acesso. Este modo não é uma plataforma SaaS com dados separados para escritórios diferentes.


## Atualização de contas a receber

Cobranças vencidas que ainda não foram baixadas como pagas continuam aparecendo nos meses seguintes até receberem baixa ou serem canceladas. Assim, pendências antigas não somem da tela do mês atual.

## Atualização: edição de cobrança inteira

Esta versão inclui a tela **Editar cobrança**, acessível pela ficha da cobrança.

Recursos adicionados:

- botão **Editar cobrança** dentro da ficha da cobrança;
- alteração de cliente, descrição, valor total, quantidade de parcelas, primeiro vencimento, forma de pagamento padrão e observações;
- recalculo automático das parcelas em aberto;
- parcelas já pagas são preservadas para não alterar o histórico financeiro;
- o campo de valor usa formato contábil em reais: **R$ 0,00**.

Ao atualizar o sistema, copie o arquivo `gestao360.db` da versão antiga para a nova pasta. Se estiver usando anexos, copie também a pasta `uploads`.

## Atualização 05 - Checklist tipo Trello

- Adicionado checklist dinâmico dentro da ficha da atividade.
- Agora é possível criar itens do checklist na hora, marcar/desmarcar, editar e excluir.
- O progresso do checklist aparece na ficha da atividade e nos cartões do Kanban.
- As ações do checklist também entram no histórico de andamento da atividade.

## Atualização 8 - Dashboard de tarefas

Foram adicionados novos indicadores ao dashboard principal para facilitar a gestão diária do escritório:

- tarefas vencidas em aberto;
- tarefas que vencem hoje;
- tarefas aguardando cliente ou documentos;
- tarefas urgentes/altas ainda abertas;
- abertas e concluídas no mês;
- listas rápidas de tarefas vencidas, para hoje e aguardando;
- status das tarefas no mês;
- prioridades no mês;
- ranking por responsável com abertas, vencidas e horas.

Esses indicadores funcionam junto com o Kanban, checklist e rotinas mensais já existentes.

## Atualização 9 - Backup e segurança

Foram adicionados recursos para proteger os dados do sistema:

- nova tela **Backup** no menu do administrador;
- geração de backup do banco `gestao360.db`;
- geração de backup completo em `.zip`, com banco de dados e anexos;
- listagem dos backups gerados, com opção de baixar ou excluir;
- tela **Alterar senha** para o usuário logado;
- alerta quando o usuário estiver usando senha padrão/fraca;
- criação de novos usuários agora exige senha inicial com pelo menos 8 caracteres.

Recomendação: gere um backup antes de excluir clientes, cobranças, anexos ou qualquer dado importante. Guarde uma cópia fora do computador principal.

## Atualização 11 - Relatórios separados

Os relatórios foram separados para não misturar operação com financeiro:

- **Relatório de atividades:** horas gastas, quantidade de atividades, valor por hora e rentabilidade;
- **Relatório de recebimentos:** somente as baixas efetivamente registradas nas contas a receber no mês escolhido;
- baixas parciais aparecem separadamente pelo valor realmente recebido;
- o relatório financeiro usa a **data da baixa**, não a data de vencimento;
- filtros por mês, cliente, forma de pagamento e responsável pela baixa;
- resumo por cliente e por forma de pagamento;
- exportação do relatório de recebimentos em CSV, Excel e PDF.


## Atualização 12 - Login próprio e preparação para Vercel

- removidos o e-mail e a senha preenchidos automaticamente na tela de login;
- removida a exibição pública das credenciais padrão;
- removida a criação automática do usuário `admin@gestao360.com` com senha `admin123`;
- primeira pessoa cria a própria conta administradora;
- membros autorizados criam a própria conta usando o código do escritório;
- novas contas entram como Colaborador;
- administrador pode ativar/desativar o cadastro próprio e trocar o código;
- sessões com cookies protegidos e limite de 20 MB por anexo;
- suporte opcional ao banco Turso/libSQL para persistência na Vercel;
- anexos são armazenados no banco quando o Turso está ativo;
- inicialização do banco também funciona quando o Flask é carregado pela Vercel.

## Publicação na Vercel

A Vercel detecta automaticamente o Flask pelo arquivo `app.py`. Não adicione uma regra `functions` apontando para `app.py`. Consulte `PUBLICAR_NA_VERCEL.md` e configure `GESTAO360_SECRET`, `TURSO_DATABASE_URL` e `TURSO_AUTH_TOKEN`.
