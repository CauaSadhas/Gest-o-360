# Publicar o Gestão360 Contábil na Vercel

Esta versão usa a detecção automática de Flask da Vercel. O arquivo principal é `app.py`. O `vercel.json` foi mantido apenas com o schema e não possui mais a regra inválida `functions`.

## 1. Estrutura que precisa estar na raiz do GitHub

Na primeira tela do repositório devem aparecer diretamente:

- `app.py`
- `requirements.txt`
- `.python-version`
- `vercel.json`
- pastas `templates`, `static` e `public`

Não deixe esses arquivos dentro de uma pasta extra. Na Vercel, o **Root Directory** deve apontar para a pasta onde está o `app.py`.

## 2. Criar o banco em nuvem

O arquivo SQLite local não é permanente na Vercel. Crie um banco Turso/libSQL e copie:

```text
TURSO_DATABASE_URL=libsql://...
TURSO_AUTH_TOKEN=...
```

## 3. Criar a chave de segurança

Crie uma sequência longa e aleatória para proteger as sessões:

```text
GESTAO360_SECRET=uma-chave-longa-diferente-e-secreta
```

Não publique essa chave no GitHub.

## 4. Configurar na Vercel

Abra **Settings → Environment Variables** e cadastre:

```text
GESTAO360_SECRET
TURSO_DATABASE_URL
TURSO_AUTH_TOKEN
```

Marque pelo menos **Production**. Para testar previews, marque também **Preview**.

## 5. Conferir as configurações do projeto

Em **Settings → Build and Deployment**:

- Framework Preset: `Other` ou detecção automática
- Build Command: deixe vazio
- Output Directory: deixe vazio
- Install Command: deixe automático
- Root Directory: pasta que contém `app.py`

## 6. Fazer novo deploy

Depois de enviar esta versão ao GitHub e salvar as variáveis, abra **Deployments**, escolha o último deploy e clique em **Redeploy**.

Na primeira abertura, o sistema mostrará **Configuração inicial** para criar a primeira conta administradora.

## Correção do erro “pattern app.py defined in functions”

Esse erro era causado por uma configuração antiga que tentava tratar `app.py` como se estivesse dentro da pasta `api`. A configuração foi removida. A Vercel reconhece automaticamente uma instância Flask chamada `app` em `app.py`.

## Atenção sobre anexos

Clientes, tarefas e cobranças ficam no banco Turso. Já os anexos enviados para atividades ainda usam armazenamento temporário da função e podem desaparecer após reinicializações. Antes de usar anexos em produção, conecte a área de arquivos a um armazenamento permanente, como Vercel Blob ou outro serviço compatível.
