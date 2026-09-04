// Проверка слияния журнала: удаление и переезд записи на другой месяц
// не должны откатываться при следующей синхронизации.
const fs = require('fs');
const path = require('path');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const script = [...html.matchAll(/<script(?![^>]*src=)[^>]*>([\s\S]*?)<\/script>/g)][0][1];

// Вытаскиваем только нужные функции слияния — остальной файл завязан на DOM.
const names = ['entryKeyOf', 'tombKey', 'entryStamp', 'pruneTombstones', 'ghMergePayload'];
const src = names.map(n => {
  const i = script.indexOf('function ' + n + '(');
  if (i < 0) throw new Error('не найдена функция ' + n);
  let depth = 0, j = script.indexOf('{', i);
  const start = i;
  for (let k = j; k < script.length; k++) {
    if (script[k] === '{') depth++;
    else if (script[k] === '}' && --depth === 0) { j = k + 1; break; }
  }
  return script.slice(start, j);
}).join('\n');

const ctx = {};
const preamble = 'const TOMBSTONE_TTL_MS = ' + (180 * 24 * 60 * 60 * 1000) + ';\n';
new Function('ctx', preamble + src + '\n' + names.map(n => `ctx.${n}=${n};`).join('')).call(null, ctx);
const { ghMergePayload } = ctx;

const now = Date.now();
const mk = (over) => Object.assign({ id: 1788502629463, kwh: 1, rubElec: 1, btcMined: 1, location: 'dema' }, over);
const pack = (entries, deleted) => ({ journals: { D: [], G: entries }, diffCache: {}, deleted: deleted || {} });

let failed = 0;
const check = (name, cond, extra) => {
  console.log((cond ? '  ok   ' : '  FAIL ') + name + (cond ? '' : '  <- ' + extra));
  if (!cond) failed++;
};
const months = (p) => p.journals.G.map(e => e.month + '|' + e.location).sort();

console.log('\n1. Переезд сентября в август (реальный баг)');
{
  const local = pack([mk({ month: '2026-08', updatedAt: now })], { 'G|2026-09|dema': now });
  const remote = pack([mk({ month: '2026-09' })]);
  const got = months(ghMergePayload(local, remote));
  check('сентябрь не воскресает', JSON.stringify(got) === JSON.stringify(['2026-08|dema']), got);
}

console.log('\n2. Удаление записи');
{
  const local = pack([], { 'G|2026-07|dema': now });
  const remote = pack([mk({ month: '2026-07', id: 1786338905320 })]);
  const got = months(ghMergePayload(local, remote));
  check('удалённая не возвращается', got.length === 0, got);
}

console.log('\n3. Запись создана заново после удаления');
{
  const local = pack([mk({ month: '2026-07', updatedAt: now + 5000 })], { 'G|2026-07|dema': now });
  const remote = pack([]);
  const got = months(ghMergePayload(local, remote));
  check('новая версия переживает надгробие', got.length === 1, got);
}

console.log('\n4. Партнёр добавил запись, которую мы не видели');
{
  const local = pack([mk({ month: '2026-08' })]);
  const remote = pack([mk({ month: '2026-08' }), mk({ month: '2026-05', id: 111 })]);
  const got = months(ghMergePayload(local, remote));
  check('чужая запись подтягивается', got.length === 2, got);
}

console.log('\n5. Правка побеждает старую версию с того же ключа');
{
  const local = pack([mk({ month: '2026-08', rubElec: 999, updatedAt: now })]);
  const remote = pack([mk({ month: '2026-08', rubElec: 111 })]);
  const merged = ghMergePayload(local, remote);
  check('осталась свежая правка', merged.journals.G[0].rubElec === 999, merged.journals.G[0].rubElec);
}

console.log('\n6. Старые записи без updatedAt');
{
  const local = pack([mk({ month: '2026-04', id: 1781271029463 })]);
  const remote = pack([mk({ month: '2026-04', id: 1781271029463 })]);
  const got = months(ghMergePayload(local, remote));
  check('не дублируются', got.length === 1, got);
}

console.log('\n7. Просроченные надгробия отбрасываются');
{
  const old = now - 200 * 24 * 3600 * 1000;
  const merged = ghMergePayload(pack([], { 'G|2026-01|dema': old }), pack([]));
  check('старое надгробие выброшено', Object.keys(merged.deleted).length === 0, merged.deleted);
}

console.log(failed ? '\n' + failed + ' проверок упало\n' : '\nвсе проверки прошли\n');
process.exit(failed ? 1 : 0);
