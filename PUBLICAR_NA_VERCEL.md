# Publicar o Gestão360 Contábil na Vercel

Esta versão foi preparada para rodar como aplicação Flask na Vercel. Antes de publicar, é obrigatório configurar um banco Turso/libSQL, porque o arquivo SQLite local não fica permanente no ambiente da Vercel.

## 1. Criar o banco em nuvem

1. Crie uma conta no Turso.
2. Crie um novo banco de dados.
3. Copie a URL do banco, que começa com `libsql://`.
4. Gere um token de acesso ao banco.

Você terá dois valores:

```text
TURSO_DATABASE_URL=libsql://...
TURSO_AUTH_TOKEN=...
```

## 2. Criar a chave de segurança

Crie uma sequência longa e aleatória para proteger as sessões do sistema. Exemplo de formato:

```text
GESTAO360_SECRET=uma-chave-longa-diferente-e-secreta
```

Não use o exemplo literalmente e não publique essa chave no GitHub.

## 3. Enviar o projeto ao GitHub

Envie todos os arquivos desta pasta, incluindo:

- `app.py`
- `requirements.txt`
- `vercel.json`
- pasta `templates`
- pasta `static`

Não envie arquivos `.env`, bancos locais, backups ou tokens.

## 4. Configurar na Vercel

No projeto da Vercel, abra **Settings → Environment Variables** e cadastre:

```text
GESTAO360_SECRET
TURSO_DATABASE_URL
TURSO_AUTH_TOKEN
```

Adicione as três variáveis para **Production**, **Preview** e **Development** quando quiser usar o mesmo banco em todos os ambientes. Para maior segurança, use bancos separados em produção e testes.

## 5. Fazer o deploy

Depois de salvar as variáveis, faça um novo deploy. Na primeira abertura, o sistema mostrará **Configuração inicial**.

Crie:

- seu nome;
- seu e-mail;
- sua senha pessoal;
- o código do escritório.

O código do escritório será usado pelos membros autorizados da equipe no botão **Criar minha conta**.

## Como funcionam as contas

- A primeira conta é Administrador.
- Novos cadastros próprios entram como Colaborador.
- O administrador pode mudar o perfil na tela **Usuários**.
- O administrador pode trocar o código ou bloquear novos cadastros em **Configurações**.
- Todos os usuários autorizados compartilham os dados do mesmo escritório.

## Atenção

Este modo é destinado à equipe de um único escritório. Ele não separa dados entre empresas diferentes. Para vender o sistema como SaaS para vários escritórios, será necessária uma estrutura de organizações isoladas.
