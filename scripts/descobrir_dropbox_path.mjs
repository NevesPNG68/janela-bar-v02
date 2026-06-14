const TOKEN = process.env.DROPBOX_TOKEN;
const FILE_NAME = process.env.DROPBOX_SEARCH_NAME || 'Por venda_2026.xlsm';

if (!TOKEN) {
  console.error('❌ Secret DROPBOX_TOKEN não encontrado no GitHub Actions.');
  process.exit(1);
}

async function dropbox(endpoint, body) {
  const resp = await fetch(`https://api.dropboxapi.com/2/${endpoint}`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${TOKEN}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(body)
  });

  const text = await resp.text();
  let data;
  try { data = text ? JSON.parse(text) : {}; } catch { data = { raw: text }; }

  if (!resp.ok) {
    console.error(`❌ Erro Dropbox em ${endpoint}:`, JSON.stringify(data, null, 2));
    process.exit(1);
  }
  return data;
}

function extractPathFromSearchMatch(match) {
  const m = match?.metadata?.metadata || match?.metadata || match;
  return m?.path_display || m?.path_lower || '';
}

async function searchV2() {
  const data = await dropbox('files/search_v2', {
    query: FILE_NAME,
    options: {
      filename_only: true,
      max_results: 20
    }
  });
  return (data.matches || [])
    .map(extractPathFromSearchMatch)
    .filter(Boolean);
}

async function listFolderRecursive() {
  const found = [];
  let data = await dropbox('files/list_folder', {
    path: '',
    recursive: true,
    include_deleted: false,
    include_non_downloadable_files: false
  });

  while (true) {
    for (const entry of data.entries || []) {
      if (entry['.tag'] === 'file' && entry.name === FILE_NAME) {
        found.push(entry.path_display || entry.path_lower);
      }
    }
    if (!data.has_more) break;
    data = await dropbox('files/list_folder/continue', { cursor: data.cursor });
  }
  return found;
}

console.log(`🔎 Procurando no Dropbox: ${FILE_NAME}`);

let paths = await searchV2();
if (!paths.length) {
  console.log('Busca rápida não encontrou. Tentando varredura completa...');
  paths = await listFolderRecursive();
}

paths = [...new Set(paths)];

if (!paths.length) {
  console.log('❌ Nenhum arquivo encontrado com esse nome. Confira se o nome está exatamente igual.');
  process.exit(2);
}

console.log('✅ Arquivo(s) encontrado(s):');
for (const p of paths) {
  console.log(`DROPBOX_FILE_PATH=${p}`);
}

console.log('\nCopie exatamente o caminho acima e salve no secret DROPBOX_FILE_PATH.');
