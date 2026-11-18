const fs = require('fs');
const path = require('path');
const vm = require('vm');

const repoRoot = path.resolve(__dirname, '..');

const gsFiles = [
  'src/project-a-audit-readonly/Codigo.gs',
  'src/project-a-audit-readonly/Roteamento_Etapa3.gs',
  'src/project-b-execution-rollback/ProjetoB_Execucao_Rollback.gs',
  'src/project-b-execution-rollback/Gerador_Teste.gs'
];

// 1. Static Verification for Forbidden Destructive Calls (Rule 2.A.3 & 8.2)
const forbiddenPatterns = [
  { name: 'Drive.Files.remove', regex: /Drive\.Files\.remove\s*\(/g },
  { name: 'Drive.Files.delete', regex: /Drive\.Files\.delete\s*\(/g },
  { name: 'emptyTrash', regex: /\bemptyTrash\s*\(/g },
  { name: 'setTrashed', regex: /\bsetTrashed\s*\(/g },
  { name: 'DriveApp.removeFile', regex: /\bremoveFile\s*\(/g },
  { name: 'trashed: true', regex: /trashed\s*:\s*true/gi }
];

console.log('=== 1. STATIC SECURITY SCAN (FORBIDDEN CALLS) ===');
for (const pat of forbiddenPatterns) {
  let count = 0;
  for (const rel of gsFiles) {
    const content = fs.readFileSync(path.join(repoRoot, rel), 'utf8');
    // Strip comments before scanning so docstrings mentioning rules don't false-positive
    const codeOnly = content
      .replace(/\/\*[\s\S]*?\*\//g, '')
      .replace(/\/\/.*$/gm, '');
    const matches = codeOnly.match(pat.regex);
    if (matches) count += matches.length;
  }
  console.log(`  ${pat.name}: ${count} ocorrências`);
  if (count > 0) {
    throw new Error(`Forbidden call detected: ${pat.name}`);
  }
}

// 2. Execute Pure Logic Unit Tests in V8 Sandbox
console.log('\n=== 2. UNIT TESTS (PROJECT A & PROJECT B PURE LOGIC) ===');
const logs = [];
const sandboxProjectA = {
  Logger: { log: (msg) => { logs.push(msg); console.log('  [Project A]', msg); } },
  Date, Math, Number, String, Boolean, Array, Object, Set, Map, RegExp, Error, JSON
};
vm.createContext(sandboxProjectA);
vm.runInContext(fs.readFileSync(path.join(repoRoot, 'src/project-a-audit-readonly/Codigo.gs'), 'utf8'), sandboxProjectA);
vm.runInContext(fs.readFileSync(path.join(repoRoot, 'src/project-a-audit-readonly/Roteamento_Etapa3.gs'), 'utf8'), sandboxProjectA);

sandboxProjectA.testes_logica();
sandboxProjectA.testes_logica_roteamento();

const sandboxProjectB = {
  Logger: { log: (msg) => { logs.push(msg); console.log('  [Project B]', msg); } },
  Date, Math, Number, String, Boolean, Array, Object, Set, Map, RegExp, Error, JSON
};
vm.createContext(sandboxProjectB);
vm.runInContext(fs.readFileSync(path.join(repoRoot, 'src/project-b-execution-rollback/ProjetoB_Execucao_Rollback.gs'), 'utf8'), sandboxProjectB);

sandboxProjectB.testes_logica_execucao();

console.log('\nALL STATIC CHECKS & V8 UNIT TESTS PASSED SUCCESSFULLY.');
