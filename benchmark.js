const { performance } = require('perf_hooks');

function generateList(size) {
  const list = [];
  for (let i = 0; i < size; i++) {
    list.push({ room: i % 2 === 0 ? `Room ${i % 10}` : null });
  }
  return list;
}

const list = generateList(1000);

function original(list) {
  return Array.from(new Set(list.map(b=>b.room).filter(Boolean))).join(' · ');
}

function optimized(list) {
  const uniqueRooms = new Set();
  for (let i = 0; i < list.length; i++) {
    const r = list[i].room;
    if (r) uniqueRooms.add(r);
  }
  return Array.from(uniqueRooms).join(' · ');
}

function runBenchmark(name, fn) {
  const start = performance.now();
  let result;
  for (let i = 0; i < 10000; i++) {
    result = fn(list);
  }
  const end = performance.now();
  console.log(`${name}: ${end - start} ms`);
  return result;
}

const r1 = runBenchmark('original', original);
const r2 = runBenchmark('optimized', optimized);

console.assert(r1 === r2, 'Results do not match');
