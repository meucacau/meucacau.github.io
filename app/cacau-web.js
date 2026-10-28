// Configuração pública do login web (um ID de cliente OAuth não é segredo).
// Preencher com o "ID do cliente" do tipo "Aplicativo da Web" criado no Console do Google Cloud,
// com a origem https://luishfilho.github.io autorizada.
window.CACAU_WEB_CONFIG = {
  clientId: '423371820715-put6v1djqjsgp42dlhnkc7ogcbns13mg.apps.googleusercontent.com',
  // Seletor de planilhas (Google Picker), usado na permissão mínima (drive.file):
  projectNumber: '423371820715', // número do projeto no Google Cloud (prefixo do ID do cliente)
  apiKey: 'AIzaSyCv8a8qGtNf09fEz0RoJkNAvjyd1deMWT8' // chave de API do navegador (pública por natureza; restrita ao site no Google Cloud)
};

/**
 * Emulação mínima do Apps Script para rodar o código do servidor (src/*.js) no navegador.
 *
 * O servidor lê e grava a planilha de forma síncrona (SpreadsheetApp). No navegador a
 * Sheets API é assíncrona, então o backend (backend.js) baixa um "snapshot" das abas ANTES
 * de cada chamada e este arquivo expõe esse snapshot com a mesma interface do
 * SpreadsheetApp: o código do servidor continua igual, só lê da memória.
 *
 * Três modos (setSnapshot options.mode):
 *   'readonly'  gravar numa aba real lança erro; aba inexistente vira "efêmera" (só memória)
 *   'discard'   gravações valem só na cópia em memória (descartada na próxima chamada)
 *   'write'     gravações valem na memória E ficam registradas em pendingOps(), para o backend
 *               enviar à Sheets API (sheets-writer.js traduz)
 */
var CacauGAS = (function () {
  var SCRIPT_TZ = 'America/Sao_Paulo'; // mesmo timeZone do appsscript.json do beta
  var SERIAL_EPOCH_MS = Date.UTC(1899, 11, 30);
  var READ_ONLY_MESSAGE = 'Modo somente leitura: para gravar, ative o "modo gravação" (use uma cópia da planilha para testar).';

  // ---------- datas ----------
  var dtfCache = {};
  function fieldsAt_(ms, tz) {
    var f = dtfCache[tz] || (dtfCache[tz] = new Intl.DateTimeFormat('en-US', {
      timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit'
    }));
    var out = {};
    f.formatToParts(new Date(ms)).forEach(function (p) {
      if (p.type !== 'literal') out[p.type] = parseInt(p.value, 10);
    });
    return out;
  }

  /** Deslocamento (ms) do fuso em relação ao UTC naquele instante. */
  function offsetAt_(ms, tz) {
    var f = fieldsAt_(ms, tz);
    var wallAsUtc = Date.UTC(f.year, f.month - 1, f.day, f.hour, f.minute, f.second);
    return wallAsUtc - Math.floor(ms / 1000) * 1000;
  }

  function pad_(n, size) {
    var s = String(n);
    while (s.length < size) s = '0' + s;
    return s;
  }

  /** Utilities.formatDate: aceita os símbolos usados pelo app (yyyy yy MM M dd d HH H mm m ss s SSS e 'literal'). */
  function formatDate(date, tz, pattern) {
    var ms = date.getTime();
    var f = fieldsAt_(ms, tz);
    var millis = ((ms % 1000) + 1000) % 1000;
    return String(pattern).replace(/'([^']*)'|yyyy|yy|MM|M|dd|d|HH|H|mm|m|ss|s|SSS/g, function (token, literal) {
      if (literal !== undefined) return literal;
      switch (token) {
        case 'yyyy': return pad_(f.year, 4);
        case 'yy': return pad_(f.year % 100, 2);
        case 'MM': return pad_(f.month, 2);
        case 'M': return String(f.month);
        case 'dd': return pad_(f.day, 2);
        case 'd': return String(f.day);
        case 'HH': return pad_(f.hour, 2);
        case 'H': return String(f.hour);
        case 'mm': return pad_(f.minute, 2);
        case 'm': return String(f.minute);
        case 'ss': return pad_(f.second, 2);
        case 's': return String(f.second);
        case 'SSS': return pad_(millis, 3);
      }
      return token;
    });
  }

  /**
   * Número de série do Sheets (dias desde 30/12/1899, no fuso da planilha) -> Date, como o
   * getValues() do Apps Script devolve para células de data/hora.
   */
  function serialToDate(serial, tz) {
    var wallMs = SERIAL_EPOCH_MS + Math.round(serial * 86400000);
    var instant = wallMs;
    for (var i = 0; i < 3; i++) instant = wallMs - offsetAt_(instant, tz);
    return new Date(instant);
  }

  /** Inverso de serialToDate: Date -> número de série do Sheets no fuso da planilha. */
  function dateToSerial(date, tz) {
    var f = fieldsAt_(date.getTime(), tz);
    var millis = ((date.getTime() % 1000) + 1000) % 1000;
    var wall = Date.UTC(f.year, f.month - 1, f.day, f.hour, f.minute, f.second) + millis;
    return (wall - SERIAL_EPOCH_MS) / 86400000;
  }

  function uuid_() {
    if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID();
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
      var r = Math.random() * 16 | 0;
      return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
    });
  }

  var Utilities = { formatDate: formatDate, getUuid: uuid_ };

  // ---------- sessão, locks, propriedades, log ----------
  var identity = { email: '' };
  function setEmail(email) { identity.email = email || ''; }

  var Session = {
    getScriptTimeZone: function () { return SCRIPT_TZ; },
    getActiveUser: function () { return { getEmail: function () { return identity.email; } }; }
  };

  function fakeLock_() {
    return { waitLock: function () {}, tryLock: function () { return true; }, releaseLock: function () {}, hasLock: function () { return true; } };
  }
  var LockService = { getUserLock: fakeLock_, getScriptLock: fakeLock_, getDocumentLock: fakeLock_ };

  var memoryStore = {};
  var storage = {
    getItem: function (k) { return Object.prototype.hasOwnProperty.call(memoryStore, k) ? memoryStore[k] : null; },
    setItem: function (k, v) { memoryStore[k] = String(v); },
    removeItem: function (k) { delete memoryStore[k]; }
  };
  function useStorage(store) { storage = store; }

  function propertyStore_(prefix) {
    return {
      getProperty: function (key) { return storage.getItem(prefix + key); },
      setProperty: function (key, value) { storage.setItem(prefix + key, value); return this; },
      deleteProperty: function (key) { storage.removeItem(prefix + key); return this; }
    };
  }
  var PropertiesService = {
    getUserProperties: function () { return propertyStore_('cacau-prop:'); },
    // No navegador não existe propriedade de script (é onde ficava o ALLOWED_EMAILS do beta).
    getScriptProperties: function () { return { getProperty: function () { return null; } }; }
  };

  var Logger = { log: function () { if (typeof console !== 'undefined') console.log.apply(console, arguments); } };

  // ---------- planilha em memória ----------
  var ctx = { current: null, mode: 'readonly', ops: [] };

  /** snapshot = { id, url, tz, sheets: [{ title, rowCount, columnCount, values: [[...]] }] } */
  function setSnapshot(snapshot, options) {
    var sheets = {};
    var order = [];
    snapshot.sheets.forEach(function (spec) {
      sheets[spec.title] = {
        name: spec.title,
        rows: spec.values.map(function (row) { return row.slice(); }),
        maxRows: Math.max(spec.rowCount || 0, spec.values.length),
        maxCols: Math.max(spec.columnCount || 0, widest_(spec.values)),
        ephemeral: false,
        created: false,
        source: spec.values // linhas originais (não são mexidas): base da checagem de conflito
      };
      order.push(spec.title);
    });
    order.forEach(function (title) { sheets[title].origLastRow = lastRowOf_(sheets[title]); });
    ctx.current = { id: snapshot.id, url: snapshot.url || '', tz: snapshot.tz || SCRIPT_TZ, sheets: sheets, order: order };
    // 'discard': leituras do servidor que "arrumam" a planilha (ex: criar o orçamento
    // "Sem orçamento") mexem só na cópia em memória, que é descartada na próxima chamada.
    ctx.mode = (options && options.mode) || 'readonly';
    ctx.ops = [];
  }

  function widest_(rows) {
    var w = 0;
    rows.forEach(function (r) { if (r.length > w) w = r.length; });
    return w;
  }

  function current_() {
    if (!ctx.current) throw new Error('Planilha ainda não carregada.');
    return ctx.current;
  }

  function guard_(sheet) {
    if (sheet.ephemeral) return;
    if (ctx.mode === 'readonly') throw new Error(READ_ONLY_MESSAGE);
  }

  function isEmptyCell_(v) { return v === '' || v === undefined || v === null; }

  function lastRowOf_(sheet) {
    for (var i = sheet.rows.length - 1; i >= 0; i--) {
      var row = sheet.rows[i];
      for (var j = 0; j < row.length; j++) if (!isEmptyCell_(row[j])) return i + 1;
    }
    return 0;
  }

  function writeCells_(s, row, col, values) {
    for (var i = 0; i < values.length; i++) {
      while (s.rows.length < row - 1 + i + 1) s.rows.push([]);
      var target = s.rows[row - 1 + i];
      for (var j = 0; j < values[i].length; j++) {
        while (target.length < col - 1 + j) target.push('');
        target[col - 1 + j] = values[i][j];
      }
    }
    if (row - 1 + values.length > s.maxRows) s.maxRows = row - 1 + values.length;
  }

  function sheetApi_(name) {
    function raw() {
      var s = current_().sheets[name];
      if (!s) throw new Error('A aba "' + name + '" não existe mais nesta planilha.');
      return s;
    }
    function rangeApi(row, col, numRows, numCols) {
      var nr = numRows === undefined ? 1 : numRows;
      var nc = numCols === undefined ? 1 : numCols;
      return {
        getValues: function () {
          var s = raw();
          var out = [];
          for (var i = 0; i < nr; i++) {
            var line = [];
            var source = s.rows[row - 1 + i] || [];
            for (var j = 0; j < nc; j++) {
              var v = source[col - 1 + j];
              line.push(v === undefined ? '' : v);
            }
            out.push(line);
          }
          return out;
        },
        getValue: function () { return this.getValues()[0][0]; },
        setValues: function (values) {
          var s = raw(); guard_(s);
          writeCells_(s, row, col, values);
          // append: só linhas novas, depois do fim lido (o backend usa "acrescentar" da API,
          // que não sobrescreve linhas que outro aparelho tenha adicionado nesse meio tempo)
          ctx.ops.push({ op: 'setValues', sheet: name, row: row, col: col, values: values,
            append: !s.created && !s.ephemeral && row > s.origLastRow });
          return this;
        },
        clearContent: function () {
          var s = raw(); guard_(s);
          for (var i = 0; i < nr; i++) {
            var target = s.rows[row - 1 + i];
            if (!target) continue;
            for (var j = 0; j < nc; j++) if (col - 1 + j < target.length) target[col - 1 + j] = '';
          }
          ctx.ops.push({ op: 'clearContent', sheet: name, row: row, col: col, numRows: nr, numCols: nc });
          return this;
        },
        setNumberFormat: function (format) {
          var s = raw(); guard_(s);
          ctx.ops.push({ op: 'setNumberFormat', sheet: name, row: row, col: col, numRows: nr, numCols: nc, format: format });
          return this;
        }
      };
    }
    return {
      getName: function () { return name; },
      getLastRow: function () { return lastRowOf_(raw()); },
      getMaxRows: function () { return raw().maxRows; },
      getMaxColumns: function () { return raw().maxCols; },
      getRange: rangeApi,
      appendRow: function (values) {
        var s = raw(); guard_(s);
        writeCells_(s, lastRowOf_(s) + 1, 1, [values]);
        ctx.ops.push({ op: 'appendRow', sheet: name, values: values, created: s.created });
        return this;
      },
      deleteRow: function (n) {
        var s = raw(); guard_(s);
        s.rows.splice(n - 1, 1);
        s.maxRows = Math.max(1, s.maxRows - 1);
        ctx.ops.push({ op: 'deleteRow', sheet: name, row: n });
        return this;
      },
      insertColumnsAfter: function (position, count) {
        var s = raw(); guard_(s);
        s.rows.forEach(function (r) {
          if (r.length > position) Array.prototype.splice.apply(r, [position, 0].concat(new Array(count).fill('')));
        });
        s.maxCols += count;
        ctx.ops.push({ op: 'insertColumnsAfter', sheet: name, position: position, count: count });
        return this;
      },
      setFrozenRows: function (n) {
        var s = raw(); guard_(s);
        ctx.ops.push({ op: 'setFrozenRows', sheet: name, count: n });
        return this;
      }
    };
  }

  // O servidor guarda o objeto da planilha em variável global entre as leituras; no Apps
  // Script cada execução zera isso, aqui não. Por isso este objeto sempre consulta o
  // snapshot ATUAL em vez de guardar um.
  var spreadsheet = {
    getId: function () { return current_().id; },
    getUrl: function () { return current_().url; },
    getSheetByName: function (name) { return current_().sheets[name] ? sheetApi_(name) : null; },
    getSheets: function () { return current_().order.map(sheetApi_); },
    insertSheet: function (name) {
      var cur = current_();
      if (cur.sheets[name]) throw new Error('Já existe uma aba chamada "' + name + '".');
      var real = ctx.mode === 'write';
      cur.sheets[name] = { name: name, rows: [], maxRows: 1000, maxCols: 26, ephemeral: !real, created: real, origLastRow: 0, source: [] };
      cur.order.push(name);
      if (real) ctx.ops.push({ op: 'insertSheet', sheet: name });
      return sheetApi_(name);
    },
    deleteSheet: function (sheet) {
      var cur = current_();
      var name = sheet.getName();
      guard_(cur.sheets[name]);
      if (ctx.mode === 'write' && !cur.sheets[name].ephemeral) ctx.ops.push({ op: 'deleteSheet', sheet: name });
      delete cur.sheets[name];
      cur.order = cur.order.filter(function (n) { return n !== name; });
    }
  };

  var SpreadsheetApp = {
    openById: function (id) {
      if (!ctx.current || ctx.current.id !== id) throw new Error('Planilha ' + id + ' não está carregada.');
      return spreadsheet;
    },
    getActiveSpreadsheet: function () { return spreadsheet; },
    getUi: function () { throw new Error('Menus da planilha não existem na versão web.'); }
  };

  return {
    Utilities: Utilities, Session: Session, LockService: LockService, PropertiesService: PropertiesService,
    Logger: Logger, SpreadsheetApp: SpreadsheetApp,
    setSnapshot: setSnapshot, setEmail: setEmail, useStorage: useStorage, serialToDate: serialToDate, dateToSerial: dateToSerial,
    /** Coluna (1-based) das linhas ORIGINAIS lidas (sem o cabeçalho): base da checagem de conflito. */
    originalColumn: function (sheetName, col) {
      var s = current_().sheets[sheetName];
      if (!s || s.created) return [];
      return s.source.slice(1, s.origLastRow).map(function (r) { var v = r[col - 1]; return v === undefined ? '' : v; });
    },
    mode: function () { return ctx.mode; },
    formatDate: formatDate, scriptTimeZone: SCRIPT_TZ, pendingOps: function () { return ctx.ops.slice(); },
    READ_ONLY_MESSAGE: READ_ONLY_MESSAGE
  };
})();

/**
 * Traduz as operações que o servidor fez no CacauGAS (pendingOps) em chamadas da Sheets API.
 *
 * Duas etapas, as duas puras (sem rede), para dar para testar:
 *   plan(ops)                        -> passos em ordem (operações vizinhas do mesmo tipo se juntam)
 *   request(step, ids, tz, toSerial) -> { path, method, body } de um passo
 * "ids" é { nomeDaAba: sheetId numérico }; só precisa conter as abas que o passo usa.
 *
 * Tipos de passo:
 *   structure  um spreadsheets:batchUpdate (criar/apagar aba, apagar linha, inserir coluna,
 *              congelar linha, formato de texto). As requests vão na ordem em que o servidor
 *              as fez, então os índices de linha se comportam como no Apps Script.
 *   values     um values:batchUpdate (vários setValues)
 *   append     um values:append (acrescentar linhas; o Sheets escolhe a linha, então não
 *              sobrescreve o que outro aparelho tenha acrescentado)
 *   clear      um values:batchClear
 */
var CacauWriter = (function () {
  function colLetter_(n) {
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  function quoteSheet_(name) { return "'" + String(name).replace(/'/g, "''") + "'"; }

  /** A1 de um bloco: 'Aba'!A2:C4 */
  function a1_(sheet, row, col, numRows, numCols) {
    return quoteSheet_(sheet) + '!' + colLetter_(col) + row + ':' + colLetter_(col + numCols - 1) + (row + numRows - 1);
  }

  function kindOf_(op) {
    switch (op.op) {
      case 'insertSheet': case 'deleteSheet': case 'deleteRow': case 'insertColumnsAfter':
      case 'setFrozenRows': case 'setNumberFormat':
        return 'structure';
      case 'clearContent': return 'clear';
      case 'appendRow': return 'append';
      case 'setValues': return op.append ? 'append' : 'values';
    }
    throw new Error('Operação desconhecida: ' + op.op);
  }

  // Operações em abas DIFERENTES não dependem uma da outra, e estas quatro (criar aba, congelar
  // linha, formato de texto, inserir coluna) também não mudam o que as demais gravam: então
  // sobem para o começo, em lotes. Sem isso, uma planilha nova (20 abas) virava ~60 chamadas.
  // deleteRow/deleteSheet/clearContent/setValues/append ficam na ordem original (índices de linha).
  var HOISTABLE = { setFrozenRows: true, setNumberFormat: true, insertColumnsAfter: true };

  function plan(ops) {
    var inserts = [], hoisted = [], rest = [];
    ops.forEach(function (op) {
      if (op.op === 'insertSheet') inserts.push(op);
      else if (HOISTABLE[op.op]) hoisted.push(op);
      else rest.push(op);
    });
    var steps = [];
    inserts.concat(hoisted, rest).forEach(function (op) {
      var kind = kindOf_(op);
      var last = steps[steps.length - 1];
      // aba criada no próprio passo ainda não tem sheetId: quem a usar vai para o passo seguinte
      var usesNewSheet = kind === 'structure' && op.op !== 'insertSheet' && last && last.kind === 'structure' &&
        last.ops.some(function (o) { return o.op === 'insertSheet' && o.sheet === op.sheet; });
      var mergeable = last && last.kind === kind && !usesNewSheet &&
        (kind !== 'append' || last.ops[0].sheet === op.sheet); // appends só se juntam na mesma aba
      if (mergeable) last.ops.push(op);
      else steps.push({ kind: kind, ops: [op] });
    });
    return steps;
  }

  /** Nomes de abas (existentes na planilha) que o passo precisa traduzir para sheetId. */
  function sheetsNeedingIds(step) {
    if (step.kind !== 'structure') return [];
    var names = {};
    step.ops.forEach(function (op) { if (op.op !== 'insertSheet') names[op.sheet] = true; });
    return Object.keys(names);
  }

  function cell_(value, tz, toSerial) {
    if (value === null || value === undefined) return '';
    if (value instanceof Date) return toSerial(value, tz);
    return value;
  }
  function matrix_(values, tz, toSerial) {
    return values.map(function (row) { return row.map(function (v) { return cell_(v, tz, toSerial); }); });
  }

  function structureRequest_(op, ids) {
    var id = ids[op.sheet];
    if (op.op !== 'insertSheet' && id === undefined) throw new Error('Não achei o código da aba "' + op.sheet + '".');
    switch (op.op) {
      case 'insertSheet': return { addSheet: { properties: { title: op.sheet } } };
      case 'deleteSheet': return { deleteSheet: { sheetId: id } };
      case 'deleteRow':
        return { deleteDimension: { range: { sheetId: id, dimension: 'ROWS', startIndex: op.row - 1, endIndex: op.row } } };
      case 'insertColumnsAfter':
        return { insertDimension: { range: { sheetId: id, dimension: 'COLUMNS', startIndex: op.position, endIndex: op.position + op.count }, inheritFromBefore: op.position > 0 } };
      case 'setFrozenRows':
        return { updateSheetProperties: { properties: { sheetId: id, gridProperties: { frozenRowCount: op.count } }, fields: 'gridProperties.frozenRowCount' } };
      case 'setNumberFormat':
        return { repeatCell: {
          range: { sheetId: id, startRowIndex: op.row - 1, endRowIndex: op.row - 1 + op.numRows, startColumnIndex: op.col - 1, endColumnIndex: op.col - 1 + op.numCols },
          cell: { userEnteredFormat: { numberFormat: op.format === '@' ? { type: 'TEXT' } : { type: 'NUMBER', pattern: op.format } } },
          fields: 'userEnteredFormat.numberFormat'
        } };
    }
    throw new Error('Operação de estrutura desconhecida: ' + op.op);
  }

  function request(step, ids, tz, toSerial) {
    var base = '';
    switch (step.kind) {
      case 'structure':
        return { method: 'POST', path: base + ':batchUpdate', body: { requests: step.ops.map(function (op) { return structureRequest_(op, ids); }) } };
      case 'values':
        return { method: 'POST', path: '/values:batchUpdate', body: {
          valueInputOption: 'USER_ENTERED',
          data: step.ops.map(function (op) {
            var vals = matrix_(op.values, tz, toSerial);
            var width = vals.reduce(function (w, r) { return Math.max(w, r.length); }, 1);
            return { range: a1_(op.sheet, op.row, op.col, vals.length, width), majorDimension: 'ROWS', values: vals };
          })
        } };
      case 'append': {
        var op = step.ops[0];
        var rows = [];
        step.ops.forEach(function (o) { rows = rows.concat(o.op === 'appendRow' ? [o.values] : o.values); });
        var vals = matrix_(rows, tz, toSerial);
        return {
          method: 'POST',
          path: '/values/' + encodeURIComponent(quoteSheet_(op.sheet) + '!A1') + ':append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS',
          body: { majorDimension: 'ROWS', values: vals }
        };
      }
      case 'clear':
        return { method: 'POST', path: '/values:batchClear', body: { ranges: step.ops.map(function (op) { return a1_(op.sheet, op.row, op.col, op.numRows, op.numCols); }) } };
    }
    throw new Error('Passo desconhecido: ' + step.kind);
  }

  /**
   * Abas e colunas que precisam ser conferidas antes de gravar: qualquer aba existente em que o
   * servidor mexeu em linhas que já existiam (editar, apagar, limpar). Acrescentar não precisa.
   */
  function sheetsToVerify(ops) {
    var names = {};
    ops.forEach(function (op) {
      if (op.op === 'deleteRow' || op.op === 'clearContent' || (op.op === 'setValues' && !op.append)) names[op.sheet] = true;
    });
    // abas criadas nesta mesma chamada não existiam: nada a conferir
    ops.forEach(function (op) { if (op.op === 'insertSheet') delete names[op.sheet]; });
    return Object.keys(names);
  }

  function createsSheets(ops) {
    return ops.some(function (op) { return op.op === 'insertSheet' || op.op === 'deleteSheet'; });
  }

  return {
    plan: plan, request: request, sheetsNeedingIds: sheetsNeedingIds, sheetsToVerify: sheetsToVerify,
    createsSheets: createsSheets, colLetter: colLetter_, a1: a1_
  };
})();

var CacauServer = (function (Utilities, Session, LockService, PropertiesService, SpreadsheetApp, Logger) {
// ===== Accounts.js =====
/** CRUD de contas + cálculo de saldo. */

var ACCOUNT_TYPES = ['checking', 'savings', 'cash', 'credit_card', 'investment'];

// Descrições dos lançamentos que o próprio sistema cria pra explicar mudança de saldo de
// conta (ver appendBalanceTransaction_).
var ACCOUNT_OPENING_DESCRIPTION = 'Início da conta - Saldo inicial';
var ACCOUNT_CONSOLIDATION_DESCRIPTION = 'Consolidação de saldo (edição manual da conta)';

function amountSignForKind_(kind) {
  return (kind === 'income' || kind === 'transfer_in') ? 1 : -1;
}

/** account: objeto de linha da aba Accounts (não o ID) — evita reler a aba pra cada conta. */
function computeAccountBalance_(account, allTransactions) {
  var txs = allTransactions || readAll_('Transactions');
  var sum = txs.reduce(function (acc, tx) {
    if (tx.AccountID !== account.ID) return acc;
    return acc + amountSignForKind_(tx.Kind) * Number(tx.Amount);
  }, 0);
  return Number(account.InitialBalance || 0) + sum;
}

// Paleta de 10 cores fixas (índices 0-9, mesma ordem de ACCOUNT_COLOR_PALETTE_ em
// JavaScript.html — o servidor só precisa saber QUANTAS existem, não os tons em si, já
// que a cor de cada índice é só uma questão de renderização no cliente). -1 é o
// sentinela de "sem cor" ESCOLHIDA pelo usuário, distinto de célula vazia (conta legada
// que nunca passou pelo seletor, ou tinha um matiz livre de antes da paleta fixa).
var ACCOUNT_COLOR_COUNT_ = 10;
function randomColorIndex_() {
  return Math.floor(Math.random() * ACCOUNT_COLOR_COUNT_);
}
/**
 * Cor padrão de uma conta que ainda não tem índice salvo — derivada do próprio ID, não
 * sorteada e gravada. Antes, listAccounts ESCREVIA (updateRowById_) pra preencher essa
 * lacuna, o que fazia um caminho de LEITURA (getBootstrapData, chamado a cada refresh)
 * pegar o lock de escrita — e logo depois de renomear a coluna ColorHue -> ColorIndex,
 * TODA conta caía nesse caso de uma vez, tornando cada bootstrap uma rajada de escritas
 * concorrendo com o que o usuário estava editando (contribuindo pros "Lock timeout").
 * Derivar do ID dá o mesmo resultado estável pra sempre, sem escrever nada — e a escolha
 * explícita do usuário (índice 0-9 ou -1 = sem cor) continua mandando quando existe.
 */
function defaultColorIndexForId_(id) {
  var text = String(id || '');
  var hash = 0;
  for (var i = 0; i < text.length; i++) {
    hash = (hash * 31 + text.charCodeAt(i)) % 100000007;
  }
  return hash % ACCOUNT_COLOR_COUNT_;
}

/**
 * Grava um lançamento que EXPLICA uma mudança de saldo da conta (abertura da conta ou
 * consolidação por edição manual), sempre com data E horário do momento da ação.
 *
 * Toda alteração de saldo/valor de conta tem que ser rastreável pelo extrato: por isso
 * o saldo inicial de uma conta nova não vive mais só na coluna InitialBalance (um número
 * que aparecia do nada, sem data, sem hora e sem linha no extrato) — ele entra como
 * lançamento "Início da conta - Saldo inicial" e InitialBalance fica 0, senão
 * computeAccountBalance_ (InitialBalance + soma dos lançamentos) contaria o valor duas
 * vezes.
 */
function appendBalanceTransaction_(accountId, accountType, closingDay, amountSigned, description) {
  var cents = Math.round(amountSigned * 100);
  if (cents === 0) return null;
  var date = todayYmd_();
  return appendRow_('Transactions', {
    Date: date,
    Time: nowHm_(),
    AccountID: accountId,
    CategoryID: '',
    Kind: cents > 0 ? 'income' : 'expense',
    Amount: Math.abs(cents) / 100,
    Description: description,
    Counterparty: '',
    BillMonth: accountType === 'credit_card' ? billMonthForDate_(date, Number(closingDay)) : '',
    InstallmentGroupID: '',
    InstallmentSeq: '',
    InstallmentTotal: '',
    TransferGroupID: '',
    ForecastID: ''
  });
}

function listAccounts(includeArchived, allTransactions) {
  var accounts = readAll_('Accounts').filter(function (a) {
    return includeArchived || !isArchived_(a.Archived);
  });
  var txs = allTransactions || readAll_('Transactions');
  return accounts.map(function (a) {
    var raw = a.ColorIndex;
    var colorIndex;
    if (raw === -1 || raw === '-1') {
      colorIndex = null; // "sem cor", escolha explícita do usuário — não é lacuna a preencher
    } else if (raw === '' || raw === undefined || raw === null || !Number.isInteger(Number(raw)) || Number(raw) < 0 || Number(raw) >= ACCOUNT_COLOR_COUNT_) {
      // Célula vazia (nunca passou pelo seletor) OU valor fora da faixa 0-9 (ex: um matiz
      // livre de antes da paleta fixa) — deriva do ID, SEM gravar (ver comentário em
      // defaultColorIndexForId_: escrever aqui colocava o lock de escrita no caminho de
      // leitura do bootstrap).
      colorIndex = defaultColorIndexForId_(a.ID);
    } else {
      colorIndex = Number(raw);
    }
    return {
      id: a.ID,
      name: a.Name,
      type: a.Type,
      initialBalance: Number(a.InitialBalance || 0),
      closingDay: a.ClosingDay === '' ? null : Number(a.ClosingDay),
      dueDay: a.DueDay === '' ? null : Number(a.DueDay),
      description: a.Description || '',
      investmentDate: a.InvestmentDate || null,
      maturityDate: a.MaturityDate || null,
      archived: isArchived_(a.Archived),
      colorIndex: colorIndex,
      balance: computeAccountBalance_(a, txs),
      hasTransactions: txs.some(function (t) { return t.AccountID === a.ID; })
    };
  });
}

/** Compara nomes ignorando maiúsculas/minúsculas e espaços nas pontas. */
function namesMatch_(a, b) {
  return (a || '').trim().toLowerCase() === (b || '').trim().toLowerCase();
}

function createAccount(input) {
  if (!input.name) throw new Error('Nome da conta é obrigatório.');
  if (ACCOUNT_TYPES.indexOf(input.type) === -1) throw new Error('Tipo de conta inválido: ' + input.type);
  var duplicate = readAll_('Accounts').some(function (a) { return !isArchived_(a.Archived) && namesMatch_(a.Name, input.name); });
  if (duplicate) throw new Error('Já existe uma conta com esse nome.');
  if (input.type === 'credit_card') {
    if (!input.closingDay || !input.dueDay) {
      throw new Error('Cartão de crédito precisa de dia de fechamento e vencimento.');
    }
  }
  var openingBalance = Number(input.initialBalance || 0);
  var saved = appendRow_('Accounts', {
    Name: input.name,
    Type: input.type,
    // 0 de propósito: o saldo inicial entra como LANÇAMENTO logo abaixo (ver
    // appendBalanceTransaction_), pra ter data/hora e aparecer no extrato.
    InitialBalance: 0,
    ClosingDay: input.type === 'credit_card' ? Number(input.closingDay) : '',
    DueDay: input.type === 'credit_card' ? Number(input.dueDay) : '',
    Description: input.type === 'investment' ? (input.description || '') : '',
    InvestmentDate: input.type === 'investment' ? (input.investmentDate || '') : '',
    MaturityDate: input.type === 'investment' ? (input.maturityDate || '') : '',
    Archived: false,
    ColorIndex: input.colorIndex === undefined ? randomColorIndex_() : (input.colorIndex === null ? -1 : Number(input.colorIndex))
  });
  var openingRow = appendBalanceTransaction_(saved.ID, input.type, saved.ClosingDay, openingBalance, ACCOUNT_OPENING_DESCRIPTION);
  // O cliente adota esse lançamento na hora (troca a prévia otimista por ele — ver
  // adoptRealId_/resolveCreatedTransactions_ em openAccountModal, JavaScript.html) em vez
  // de esperar o refresh coalescido; sem devolvê-lo aqui não haveria como fazer isso.
  return { account: saved, transaction: openingRow ? mapTransactionRow_(openingRow) : null };
}

function updateAccount(input) {
  if (!input.id) throw new Error('ID da conta é obrigatório.');
  var current = findById_('Accounts', input.id);
  if (!current) throw new Error('Conta não encontrada.');

  var updates = {};
  if (input.name !== undefined) {
    if (!input.name) throw new Error('Nome da conta é obrigatório.');
    var duplicate = readAll_('Accounts').some(function (a) {
      return a.ID !== input.id && !isArchived_(a.Archived) && namesMatch_(a.Name, input.name);
    });
    if (duplicate) throw new Error('Já existe uma conta com esse nome.');
    updates.Name = input.name;
  }

  if (input.type !== undefined && input.type !== current.Type) {
    if (ACCOUNT_TYPES.indexOf(input.type) === -1) throw new Error('Tipo de conta inválido: ' + input.type);
    var hasTx = readAll_('Transactions').some(function (t) { return t.AccountID === input.id; });
    if (hasTx) throw new Error('Não é possível mudar o tipo de uma conta que já tem lançamentos.');
    updates.Type = input.type;
  }

  if (input.closingDay !== undefined) updates.ClosingDay = input.closingDay === null ? '' : Number(input.closingDay);
  if (input.dueDay !== undefined) updates.DueDay = input.dueDay === null ? '' : Number(input.dueDay);
  if (input.description !== undefined) updates.Description = input.description || '';
  if (input.investmentDate !== undefined) updates.InvestmentDate = input.investmentDate || '';
  if (input.maturityDate !== undefined) updates.MaturityDate = input.maturityDate || '';
  if (input.colorIndex !== undefined) updates.ColorIndex = input.colorIndex === null ? -1 : Number(input.colorIndex);

  var resultingType = updates.Type !== undefined ? updates.Type : current.Type;
  if (resultingType === 'credit_card') {
    var closing = updates.ClosingDay !== undefined ? updates.ClosingDay : current.ClosingDay;
    var due = updates.DueDay !== undefined ? updates.DueDay : current.DueDay;
    if (!closing || !due) throw new Error('Cartão de crédito precisa de dia de fechamento e vencimento.');
  }

  // Saldo negativo só é bloqueado fora de cartão de crédito — lá, negativo é o normal (dívida).
  var balanceMustNotBeNegative = resultingType !== 'credit_card';

  // Editar saldo é SEMPRE "saldo atual": a diferença vira um lançamento de consolidação
  // com data e hora (input.currentBalance). InitialBalance não é mais editável por aqui —
  // ele só existe como resíduo de conta antiga; conta nova nasce com 0 e o saldo inicial
  // dela é um lançamento (ver appendBalanceTransaction_).
  var consolidationRow = null;
  if (input.currentBalance !== undefined) {
    var target = Number(input.currentBalance);
    if (balanceMustNotBeNegative && target < 0) throw new Error('Saldo não pode ser negativo nesse tipo de conta.');
    var computedBalance = computeAccountBalance_(current, readAll_('Transactions'));
    var diffCents = Math.round(target * 100) - Math.round(computedBalance * 100);
    var closingDayForBill = updates.ClosingDay !== undefined ? updates.ClosingDay : current.ClosingDay;
    consolidationRow = appendBalanceTransaction_(input.id, resultingType, closingDayForBill, diffCents / 100, ACCOUNT_CONSOLIDATION_DESCRIPTION);
  }

  var updated = updateRowById_('Accounts', input.id, updates);
  if (!updated) throw new Error('Conta não encontrada.');
  return { account: updated, transaction: consolidationRow ? mapTransactionRow_(consolidationRow) : null };
}

/**
 * Migração de uma vez (menu Finanças): converte o InitialBalance que sobrou em conta
 * ANTIGA num lançamento "Início da conta - Saldo inicial", zerando a coluna. Depois
 * disso, TODO saldo de conta é formado só por lançamentos — nada de número solto sem
 * data, hora nem linha no extrato (conta nova já nasce assim, ver createAccount).
 *
 * A data do lançamento é a de criação da conta (CreatedAt), não hoje: o dinheiro já
 * estava lá desde então, e datar hoje jogaria o valor todo pro mês atual, estragando
 * qualquer comparação de receita/despesa por mês. O horário vai junto quando o CreatedAt
 * tem hora; senão fica vazio.
 *
 * Idempotente por construção: só mexe em conta com InitialBalance != 0, e zera a coluna
 * no mesmo passo — rodar de novo não acha mais nada pra fazer. Não depende de procurar o
 * lançamento pela descrição (que o usuário pode ter editado).
 */
function migrateInitialBalancesToTransactions() {
  var accounts = readAll_('Accounts');
  var migrated = [];
  accounts.forEach(function (a) {
    var initial = Number(a.InitialBalance || 0);
    if (Math.round(initial * 100) === 0) return;
    var createdAt = a.CreatedAt ? new Date(a.CreatedAt) : null;
    var valid = createdAt && !isNaN(createdAt.getTime());
    var tz = Session.getScriptTimeZone();
    var date = valid ? Utilities.formatDate(createdAt, tz, 'yyyy-MM-dd') : todayYmd_();
    var time = valid ? Utilities.formatDate(createdAt, tz, 'HH:mm') : '';
    appendRow_('Transactions', {
      Date: date,
      Time: time,
      AccountID: a.ID,
      CategoryID: '',
      Kind: initial > 0 ? 'income' : 'expense',
      Amount: Math.abs(Math.round(initial * 100)) / 100,
      Description: ACCOUNT_OPENING_DESCRIPTION,
      Counterparty: '',
      BillMonth: a.Type === 'credit_card' ? billMonthForDate_(date, Number(a.ClosingDay)) : '',
      InstallmentGroupID: '',
      InstallmentSeq: '',
      InstallmentTotal: '',
      TransferGroupID: '',
      ForecastID: ''
    });
    updateRowById_('Accounts', a.ID, { InitialBalance: 0 });
    migrated.push(a.Name + ' (' + initial.toFixed(2) + ' em ' + date + ')');
  });
  return migrated;
}

/** Versão pro menu da planilha — mostra o resultado num alert. */
function migrateInitialBalancesMenu() {
  var migrated = migrateInitialBalancesToTransactions();
  var ui = SpreadsheetApp.getUi();
  if (!migrated.length) {
    ui.alert('Nenhuma conta tinha saldo inicial solto — todos os saldos já vêm de lançamentos.');
    return;
  }
  ui.alert(migrated.length + ' conta(s) migrada(s):\n\n' + migrated.join('\n') +
    '\n\nO saldo de cada uma continua o mesmo; agora ele vem de um lançamento com data e hora.');
}

function archiveAccount(id) {
  var updated = updateRowById_('Accounts', id, { Archived: true });
  if (!updated) throw new Error('Conta não encontrada.');
  return updated;
}

// ===== Branding.js =====
/**
 * Nome e ícone do app. O beta multiusuário se chama "Cacau" (nome provisório); o app do
 * dono (produção) continua "Finanças". Funções, e não variáveis, porque o Apps Script
 * carrega os arquivos em ordem não garantida e MULTIUSER_MODE vem de Mode.js.
 */
function appName_() {
  return MULTIUSER_MODE ? 'Cacau' : 'Finanças';
}

// Favicon só no beta; o do dono mantém o ícone padrão. O Apps Script deduz o tipo da imagem pela
// extensão da URL, então precisa ser um https público terminando em .png (um data: não funciona).
// O arquivo fica no site do Cacau (repositório luishfilho/cacau, GitHub Pages).
var CACAU_FAVICON_URL = 'https://luishfilho.github.io/cacau/favicon.png';

/** Aplica título e (no beta) favicon a uma página HtmlService. O favicon é opcional: se o Apps Script recusar a URL, a página abre igual. */
function brandPage_(output) {
  output.setTitle(appName_());
  if (MULTIUSER_MODE) {
    try {
      output.setFaviconUrl(CACAU_FAVICON_URL);
    } catch (err) {
      Logger.log('favicon recusado: ' + err);
    }
  }
  return output;
}

// ===== Budgets.js =====
/**
 * CRUD de orçamento por categoria (ver Budgets no client — "Orçamentos"). Um orçamento
 * é uma TAXA — um valor esperado a cada N dias (ex: R$100 a cada 7 dias) — em vez de um
 * valor fixo mensal. Isso permite tanto acompanhar "nos últimos N dias" quanto projetar
 * pra um mês de calendário específico (escalando pelo número de dias do mês ÷ N, ver
 * JavaScript.html). Repete automaticamente até ser alterado ou removido; todo o cálculo
 * de gasto/progresso é feito no cliente a partir de STATE.allTransactions, então essa
 * aba só guarda a configuração.
 *
 * CategoryID guarda uma lista separada por vírgula (mesmo padrão de Categories.Keywords)
 * — um orçamento agora pode cobrir várias categorias/subcategorias de uma vez (ver
 * createBudget), não mais só uma categoria de topo. O nome da coluna ficou o mesmo
 * (singular) de propósito, pra orçamentos já existentes (uma única categoria, sem
 * vírgula) continuarem lendo certo sem precisar migrar nada.
 */

function parseCategoryIdList_(raw) {
  return (raw || '').split(',').map(function (id) { return id.trim(); }).filter(function (id) { return id; });
}

/**
 * categoryIds de verdade de um orçamento: se AutoIncludeParentId estiver preenchido,
 * IGNORA o CategoryID guardado (que é só uma foto de quando foi criado/editado) e
 * recalcula categoria + subcategorias ATUAIS — é isso que faz uma subcategoria nova
 * entrar sozinha num orçamento de "família completa" sem precisar editar nada. allCategories
 * é passado de fora (readAll_('Categories')) pra não reler a aba uma vez por orçamento.
 */
function effectiveCategoryIdsForBudgetRow_(b, allCategories) {
  if (b.AutoIncludeParentId) {
    var childrenIds = allCategories
      .filter(function (c) { return c.ParentID === b.AutoIncludeParentId && !isArchived_(c.Archived); })
      .map(function (c) { return c.ID; });
    return [b.AutoIncludeParentId].concat(childrenIds);
  }
  return parseCategoryIdList_(b.CategoryID);
}

/**
 * Decide se a seleção de categorias de um orçamento cai num dos 2 casos que mudam
 * nome/comportamento (ver createBudget/updateBudget):
 *  - "família completa": a categoria de topo escolhida + TODAS as suas subcategorias
 *    ATUAIS estão selecionadas, e nada mais -> autoIncludeParentId (nome vira só o da
 *    categoria; uma subcategoria nova criada depois passa a contar sozinha).
 *  - "só a categoria, sem subs": só a categoria de topo está selecionada, sem nenhuma
 *    subcategoria (porque ela não tem, ou porque não foram escolhidas) -> soloParentId
 *    (nome ganha o sufixo " (sem subcategorias)").
 * Qualquer outra combinação (múltiplas categorias de topo, seleção parcial de
 * subcategorias, subcategoria sem a categoria-mãe) não bate com nenhum caso — nome
 * continua sendo a junção alfabética de sempre.
 */
function classifyBudgetCategorySelection_(categoryIds, allCategories) {
  var categoriesById = {};
  allCategories.forEach(function (c) { categoriesById[c.ID] = c; });
  var topLevelSelected = categoryIds.filter(function (id) { return categoriesById[id] && !categoriesById[id].ParentID; });
  if (topLevelSelected.length !== 1) return {};
  var parentId = topLevelSelected[0];
  var otherSelectedIds = categoryIds.filter(function (id) { return id !== parentId; });
  if (otherSelectedIds.length === 0) return { soloParentId: parentId };

  var currentChildrenIds = allCategories
    .filter(function (c) { return c.ParentID === parentId && !isArchived_(c.Archived); })
    .map(function (c) { return c.ID; });
  var otherSelectedSet = {};
  otherSelectedIds.forEach(function (id) { otherSelectedSet[id] = true; });
  var matchesAllChildren = currentChildrenIds.length > 0 &&
    currentChildrenIds.length === otherSelectedIds.length &&
    currentChildrenIds.every(function (id) { return otherSelectedSet[id]; });
  return matchesAllChildren ? { autoIncludeParentId: parentId } : {};
}

/**
 * Garante que a linha especial de "não orçamentado" (ver computeBudgetAutoProgress_,
 * JavaScript.html) já exista, criando-a na primeira vez que alguém abrir o app depois
 * desse recurso existir -- sem precisar de nenhuma migração manual. CategoryID fica
 * vazio de propósito (a cobertura de categoria é sempre recalculada ao vivo, nunca lida
 * dessa linha). PeriodMode 'auto' sempre -- essa barra não tem período/valor manual
 * configurável, unificada com todo orçamento Automático (usa o mesmo X derivado de Y),
 * ver listBudgets/grilling.
 */
function ensureUnbudgetedCatchAllBudget_(rows) {
  var existing = rows.filter(function (b) { return parseBoolCell_(b.IsUnbudgetedCatchAll); })[0];
  if (existing) return rows;
  rows.push(appendRow_('Budgets', {
    CategoryID: '', MonthlyLimit: 0, IntervalDays: 30, AutoIncludeParentId: '',
    Archived: false, IsUnbudgetedCatchAll: true, PeriodMode: 'auto'
  }));
  return rows;
}

function listBudgets(includeArchived) {
  var allCategories = readAll_('Categories');
  var categoriesById = {};
  allCategories.forEach(function (c) { categoriesById[c.ID] = c; });

  var rows = ensureUnbudgetedCatchAllBudget_(readAll_('Budgets'));

  return rows
    .filter(function (b) { return includeArchived || !isArchived_(b.Archived); })
    .map(function (b) {
      var categoryIds = effectiveCategoryIdsForBudgetRow_(b, allCategories);
      var categoryName;
      if (b.AutoIncludeParentId && categoriesById[b.AutoIncludeParentId]) {
        categoryName = categoriesById[b.AutoIncludeParentId].Name;
      } else {
        var names = categoryIds.map(function (id) {
          var c = categoriesById[id];
          return c ? c.Name : 'Categoria removida';
        }).sort(function (a, b2) { return a.localeCompare(b2, 'pt-BR'); });
        var isSoloTopLevel = categoryIds.length === 1 && categoriesById[categoryIds[0]] && !categoriesById[categoryIds[0]].ParentID;
        categoryName = isSoloTopLevel ? names[0] + ' (sem subcategorias)' : names.join(' / ');
      }
      var isCatchAll = parseBoolCell_(b.IsUnbudgetedCatchAll);
      return {
        id: b.ID,
        categoryIds: categoryIds,
        categoryName: categoryName,
        autoIncludeParentId: b.AutoIncludeParentId || null,
        amount: Number(b.MonthlyLimit),
        // Orçamentos criados antes de existir intervalo de dias equivalem a "30 dias".
        intervalDays: b.IntervalDays ? Number(b.IntervalDays) : 30,
        archived: isArchived_(b.Archived),
        isUnbudgetedCatchAll: isCatchAll,
        // A barra catch-all é SEMPRE 'auto' (nunca teve PeriodMode próprio antes desse
        // recurso existir) -- pra todo o resto, linha antiga sem essa coluna = 'manual'
        // (ver SCHEMA, SheetService.js).
        periodMode: isCatchAll ? 'auto' : (b.PeriodMode || 'manual'),
        createdAt: b.CreatedAt ? String(b.CreatedAt).slice(0, 10) : null
      };
    });
}

function validateIntervalDays_(intervalDays) {
  var n = Number(intervalDays);
  if (!Number.isInteger(n) || n < 1 || n > 999) throw new Error('Intervalo precisa ser um número inteiro entre 1 e 999 dias.');
  return n;
}

/**
 * Valida uma lista de categoryIds pra um orçamento: todas precisam existir e ser de
 * DESPESA (sem mais a restrição de serem só de topo — ver
 * rankedTopCategoryItems_/buildCategoryFilterField_ no client, que agora oferece o mesmo
 * seletor multi-seleção do filtro do Extrato), e nenhuma pode já estar em OUTRO orçamento
 * ativo — usa effectiveCategoryIdsForBudgetRow_ (não o CategoryID cru) pra também
 * enxergar subcategorias que um orçamento de "família completa" inclui automaticamente,
 * mesmo que tenham sido criadas depois dele e por isso não apareçam no CategoryID
 * guardado. excludeBudgetId (opcional, usado por updateBudget): ignora as categorias do
 * PRÓPRIO orçamento sendo editado na checagem de "já está em outro orçamento".
 */
function validateBudgetCategoryIds_(categoryIds, excludeBudgetId) {
  if (!categoryIds.length) throw new Error('Escolha pelo menos uma categoria.');
  var categories = categoryIds.map(function (id) {
    var category = findById_('Categories', id);
    if (!category) throw new Error('Categoria não encontrada.');
    if (category.Kind !== 'expense') throw new Error('Orçamento só se aplica a categorias de despesa.');
    return category;
  });
  var allCategories = readAll_('Categories');
  var alreadyBudgeted = {};
  readAll_('Budgets').forEach(function (b) {
    if (isArchived_(b.Archived) || b.ID === excludeBudgetId) return;
    effectiveCategoryIdsForBudgetRow_(b, allCategories).forEach(function (id) { alreadyBudgeted[id] = true; });
  });
  var duplicate = categories.find(function (c) { return alreadyBudgeted[c.ID]; });
  if (duplicate) throw new Error('"' + duplicate.Name + '" já está em outro orçamento.');
}

/** periodMode 'auto' (ver grilling): amount/intervalDays viram só placeholders inertes
 *  (o client já manda um valor válido qualquer, nunca lido de verdade -- ver
 *  computeBudgetAutoProgress_/openBudgetModal, JavaScript.html) -- só valida de verdade
 *  em modo 'manual', onde eles são o teto/período reais da barra. */
function createBudget(input) {
  var categoryIds = Array.isArray(input.categoryIds) ? input.categoryIds : parseCategoryIdList_(input.categoryIds);
  validateBudgetCategoryIds_(categoryIds, null);

  var periodMode = input.periodMode === 'auto' ? 'auto' : 'manual';
  var amount, intervalDays;
  if (periodMode === 'auto') {
    amount = Number(input.amount) || 1;
    intervalDays = validateIntervalDays_(input.intervalDays || 30);
  } else {
    amount = Number(input.amount);
    if (!amount || amount <= 0) throw new Error('Valor precisa ser maior que zero.');
    intervalDays = validateIntervalDays_(input.intervalDays);
  }
  var classification = classifyBudgetCategorySelection_(categoryIds, readAll_('Categories'));

  return appendRow_('Budgets', {
    CategoryID: categoryIds.join(','), MonthlyLimit: amount, IntervalDays: intervalDays,
    AutoIncludeParentId: classification.autoIncludeParentId || '', Archived: false, PeriodMode: periodMode
  });
}

/** input.periodMode ('auto'|'manual', opcional): trocar PRA 'auto' nunca vem acompanhado
 *  de amount/intervalDays (ver openBudgetModal, JavaScript.html) -- o valor/período
 *  manual anterior fica intacto na planilha sozinho, sem precisar de coluna extra pra
 *  "lembrar" nada (ver comentário no SCHEMA, SheetService.js). */
function updateBudget(input) {
  if (!input.id) throw new Error('ID do orçamento é obrigatório.');
  var updates = {};
  if (input.periodMode !== undefined) updates.PeriodMode = input.periodMode === 'auto' ? 'auto' : 'manual';
  if (input.amount !== undefined) {
    var amount = Number(input.amount);
    if (!amount || amount <= 0) throw new Error('Valor precisa ser maior que zero.');
    updates.MonthlyLimit = amount;
  }
  if (input.intervalDays !== undefined) updates.IntervalDays = validateIntervalDays_(input.intervalDays);
  if (input.categoryIds !== undefined) {
    var categoryIds = Array.isArray(input.categoryIds) ? input.categoryIds : parseCategoryIdList_(input.categoryIds);
    validateBudgetCategoryIds_(categoryIds, input.id);
    updates.CategoryID = categoryIds.join(',');
    var classification = classifyBudgetCategorySelection_(categoryIds, readAll_('Categories'));
    updates.AutoIncludeParentId = classification.autoIncludeParentId || '';
  }
  var updated = updateRowById_('Budgets', input.id, updates);
  if (!updated) throw new Error('Orçamento não encontrado.');
  return updated;
}

function removeBudget(id) {
  var removed = deleteRowById_('Budgets', id);
  if (!removed) throw new Error('Orçamento não encontrado.');
  return true;
}

var BUDGET_SUGGESTION_DISMISS_DAYS = 30;

/** Descarta uma sugestão de NOVO orçamento (ver computeBudgetSuggestions_,
 *  JavaScript.html) por BUDGET_SUGGESTION_DISMISS_DAYS dias, não pra sempre. Upsert por
 *  categoryId: descartar de novo só atualiza a data, não acumula linha. SuggestionType
 *  continua gravado (sempre 'new' agora -- a sugestão de AUMENTO de período que também
 *  usava essa aba foi removida) só pra linhas antigas continuarem lendo certo, sem
 *  precisar migrar nada. */
function dismissBudgetSuggestion(categoryId) {
  if (!categoryId) throw new Error('ID é obrigatório.');
  var existing = readAll_('BudgetSuggestionDismissals').find(function (d) {
    return (d.SuggestionType || 'new') === 'new' && d.CategoryID === categoryId;
  });
  if (existing) {
    updateRowById_('BudgetSuggestionDismissals', existing.ID, { DismissedAt: nowIso_() });
  } else {
    appendRow_('BudgetSuggestionDismissals', { DismissedAt: nowIso_(), SuggestionType: 'new', CategoryID: categoryId });
  }
  return true;
}

/** Descarta uma sugestão de TROCAR O MODO (Manual<->Automático) de um orçamento já
 *  existente (ver budgetModeSwitchSuggestion_, JavaScript.html) por BUDGET_SUGGESTION_
 *  DISMISS_DAYS dias -- mesmo mecanismo/janela de dismissBudgetSuggestion, só que upsert
 *  por BudgetID (um orçamento pode cobrir várias categorias, CategoryID sozinho não seria
 *  uma chave única aqui -- é exatamente pra esse caso que a coluna BudgetID já existia,
 *  ver comentário da SCHEMA em SheetService.js). SuggestionType vira 'switch-manual' ou
 *  'switch-auto' conforme a direção, pra sugerir Automático de novo não ficar preso pelo
 *  descarte de Manual (e vice-versa) — são sugestões independentes. */
function dismissBudgetModeSwitchSuggestion(input) {
  if (!input || !input.budgetId) throw new Error('ID do orçamento é obrigatório.');
  var suggestionType = input.direction === 'to-manual' ? 'switch-manual' : 'switch-auto';
  var existing = readAll_('BudgetSuggestionDismissals').find(function (d) {
    return d.SuggestionType === suggestionType && d.BudgetID === input.budgetId;
  });
  if (existing) {
    updateRowById_('BudgetSuggestionDismissals', existing.ID, { DismissedAt: nowIso_() });
  } else {
    appendRow_('BudgetSuggestionDismissals', { DismissedAt: nowIso_(), SuggestionType: suggestionType, BudgetID: input.budgetId });
  }
  return true;
}

/**
 * Expõe os descartes ainda VÁLIDOS (dentro dos 30 dias) pro cliente decidir quais
 * sugestões não devem aparecer de novo — ver computeBudgetSuggestions_/
 * budgetModeSwitchSuggestion_. Purga sozinha (autocorretiva, igual
 * purgeJunkPendingEntries_) os descartes mais antigos que isso, que já perderam efeito e
 * só ocupariam espaço à toa. Ignora linhas antigas do tipo 'increase' (a sugestão de
 * AUMENTO de período foi removida) -- ficam de fora do resultado e são purgadas pelo
 * cutoff normal como qualquer outra, sem precisar de migração.
 */
function listBudgetSuggestionDismissals() {
  var cutoff = new Date(Date.now() - BUDGET_SUGGESTION_DISMISS_DAYS * 24 * 60 * 60 * 1000);
  var rows = readAll_('BudgetSuggestionDismissals');
  var stillValid = rows.filter(function (d) { return new Date(d.DismissedAt) >= cutoff; });
  if (stillValid.length !== rows.length) {
    withSheetLock_(function () {
      var sheet = getOrCreateSheet_('BudgetSuggestionDismissals');
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      var headers = SCHEMA.BudgetSuggestionDismissals;
      sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
      if (stillValid.length) {
        sheet.getRange(2, 1, stillValid.length, headers.length).setValues(
          stillValid.map(function (row) { return objectToRow_(headers, row); })
        );
      }
    });
  }
  return stillValid
    .filter(function (d) { return (d.SuggestionType || 'new') === 'new' || d.SuggestionType === 'switch-manual' || d.SuggestionType === 'switch-auto'; })
    .map(function (d) {
      var type = d.SuggestionType || 'new';
      return type === 'new'
        ? { type: 'new', categoryId: d.CategoryID || null, dismissedAt: isoCell_(d.DismissedAt) }
        : { type: type, budgetId: d.BudgetID || null, dismissedAt: isoCell_(d.DismissedAt) };
    });
}

// ===== Categories.js =====
/** CRUD de categorias (com subcategorias via ParentID). */

var CATEGORY_RECENT_USE_DAYS = 30;

/**
 * IDs de categoria/subcategoria com pelo menos um lançamento CRIADO (CreatedAt, não a
 * Date do lançamento em si) nos últimos CATEGORY_RECENT_USE_DAYS dias — usado pelo
 * cliente pra decidir o que esconder atrás do "+" nas cápsulas de categoria (ver
 * collapseUnused em createCapsuleField, JavaScript.html) quando há muitas.
 */
function recentlyUsedCategoryIds_(allTransactions) {
  var cutoff = new Date(Date.now() - CATEGORY_RECENT_USE_DAYS * 24 * 60 * 60 * 1000);
  var ids = {};
  (allTransactions || []).forEach(function (t) {
    if (t.CategoryID && new Date(t.CreatedAt) >= cutoff) ids[t.CategoryID] = true;
  });
  return ids;
}

/**
 * allTransactions é opcional: sem ele, nada é marcado como recentlyUsed (o cliente
 * simplesmente não esconde nada atrás do "+").
 *
 * Uma subcategoria conta como arquivada quando a categoria PAI está arquivada, mesmo que
 * a linha dela própria não esteja: arquivar uma categoria de topo passou a arquivar as
 * filhas junto (ver archiveCategory), mas dado ANTIGO ficou com filhas soltas — e elas
 * reapareciam nos seletores como "órfãs" muito depois do usuário achar que tinha
 * arquivado aquele galho inteiro.
 */
function listCategories(includeArchived, allTransactions) {
  var recentIds = recentlyUsedCategoryIds_(allTransactions);
  var all = readAll_('Categories');
  var archivedById = {};
  all.forEach(function (c) { archivedById[c.ID] = isArchived_(c.Archived); });
  function effectivelyArchived_(c) {
    return archivedById[c.ID] || (c.ParentID ? !!archivedById[c.ParentID] : false);
  }
  return all
    .filter(function (c) { return includeArchived || !effectivelyArchived_(c); })
    .map(function (c) {
      return {
        id: c.ID,
        name: c.Name,
        parentId: c.ParentID || null,
        kind: c.Kind,
        archived: isArchived_(c.Archived),
        keywords: c.Keywords || '',
        description: c.Description || '',
        // '' = sem ícone escolhido à mão -- o cliente casa um automaticamente pelo nome
        // (ver matchCategoryIconKey_/resolveCategoryIconKey_, JavaScript.html). Só tem
        // valor aqui quando o usuário escolheu explicitamente no seletor de ícones.
        icon: c.Icon || '',
        recentlyUsed: !!recentIds[c.ID]
      };
    });
}

function categoryNamesMatch_(a, b) {
  return (a || '').trim().toLowerCase() === (b || '').trim().toLowerCase();
}

/** "airbnb,  Booking.com ,hotel" -> "airbnb, Booking.com, hotel" (trim cada uma, remove vazias). */
function normalizeKeywordsInput_(raw) {
  return (raw || '').split(',').map(function (k) { return k.trim(); }).filter(function (k) { return k; }).join(', ');
}

/** "airbnb, Booking.com" -> ["airbnb", "booking.com"] (minúsculo, pronto pra comparar). */
function parseKeywords_(raw) {
  return (raw || '').split(',').map(function (k) { return k.trim().toLowerCase(); }).filter(function (k) { return k; });
}

/**
 * Primeira categoria (não arquivada) cuja lista de palavras-chave contém alguma como
 * substring (case-insensitive) do texto dado — usado como sugestão automática de
 * categoria pra pendências/lançamentos (ex: "airbnb" na subcategoria Hospedagem casa
 * com o "Pago a" de uma pendência da Airbnb). Recebe a lista de categorias já carregada
 * pra não reler a aba a cada pendência (ver listPendingEntries, Inbox.js).
 */
function findCategoryIdByKeyword_(text, categories) {
  if (!text) return null;
  var lower = text.toLowerCase();
  var match = categories.filter(function (c) { return !isArchived_(c.Archived) && c.Keywords; }).find(function (c) {
    return parseKeywords_(c.Keywords).some(function (kw) { return lower.indexOf(kw) !== -1; });
  });
  return match ? match.ID : null;
}

function createCategory(input) {
  if (!input.name) throw new Error('Nome da categoria é obrigatório.');
  if (input.kind !== 'income' && input.kind !== 'expense') throw new Error('Kind inválido: ' + input.kind);
  if (input.parentId) {
    var parent = findById_('Categories', input.parentId);
    if (!parent) throw new Error('Categoria pai não encontrada.');
    if (parent.Kind !== input.kind) throw new Error('Subcategoria precisa ter o mesmo tipo (receita/despesa) da categoria pai.');
  }
  var duplicate = readAll_('Categories').some(function (c) { return !isArchived_(c.Archived) && categoryNamesMatch_(c.Name, input.name); });
  if (duplicate) throw new Error('Já existe uma categoria com esse nome.');
  return appendRow_('Categories', {
    Name: input.name,
    ParentID: input.parentId || '',
    Kind: input.kind,
    Archived: false,
    Keywords: normalizeKeywordsInput_(input.keywords),
    Description: input.description || '',
    Icon: input.icon || ''
  });
}

function updateCategory(input) {
  if (!input.id) throw new Error('ID da categoria é obrigatório.');
  var updates = {};
  if (input.name !== undefined) {
    if (!input.name) throw new Error('Nome da categoria é obrigatório.');
    var duplicate = readAll_('Categories').some(function (c) {
      return c.ID !== input.id && !isArchived_(c.Archived) && categoryNamesMatch_(c.Name, input.name);
    });
    if (duplicate) throw new Error('Já existe uma categoria com esse nome.');
    updates.Name = input.name;
  }
  if (input.keywords !== undefined) updates.Keywords = normalizeKeywordsInput_(input.keywords);
  if (input.description !== undefined) updates.Description = input.description || '';
  // '' explícito LIMPA um ícone escolhido antes (volta a casar automaticamente pelo
  // nome) -- por isso o teste é !== undefined, não uma checagem de truthy.
  if (input.icon !== undefined) updates.Icon = input.icon || '';
  var updated = updateRowById_('Categories', input.id, updates);
  if (!updated) throw new Error('Categoria não encontrada.');
  return updated;
}

/**
 * Arquiva a categoria E as subcategorias dela. Antes só a linha pedida era marcada, o que
 * deixava as filhas visíveis nos seletores (sem a mãe por perto) — o usuário arquivava um
 * galho e continuava vendo pedaços dele. Ver também effectivelyArchived_ em
 * listCategories, que cobre o dado antigo que já ficou nesse estado.
 */
function archiveCategory(id) {
  var updated = updateRowById_('Categories', id, { Archived: true });
  if (!updated) throw new Error('Categoria não encontrada.');
  readAll_('Categories')
    .filter(function (c) { return c.ParentID === id && !isArchived_(c.Archived); })
    .forEach(function (c) { updateRowById_('Categories', c.ID, { Archived: true }); });
  return updated;
}

// ===== Code.js =====
/** Ponto de entrada do web app. */

// Deployment "Só eu" usado pela interface humana. O webhook do MacroDroid roda num
// SEGUNDO deployment (acesso "Qualquer pessoa", criado manualmente pelo editor do
// Apps Script — ver instruções). Esse segundo deployment também executa doGet, então
// é essencial bloquear a interface nele: sem essa checagem, qualquer pessoa com o link
// do webhook veria o app inteiro (e o google.script.run funcionaria com a autoridade
// de quem publicou, já que o modo de execução é "Executar como: Eu").
var MAIN_WEBAPP_DEPLOYMENT_ID = 'AKfycbxqeloHyULywGtZ6YbFgregoHOaSEHaPK5pmlUlAY-AIQbnss1WPZ3MCgcWslczfDrWpA';

function doGet() {
  if (MULTIUSER_MODE) return doGetMultiUser_();
  var currentUrl = ScriptApp.getService().getUrl() || '';
  if (currentUrl.indexOf(MAIN_WEBAPP_DEPLOYMENT_ID) === -1) {
    return ContentService.createTextOutput('Not found');
  }
  return renderIndex_(false);
}

/**
 * Beta multiusuário: o script roda COMO QUEM ACESSA, então não há deployment secundário
 * pra proteger (o MAIN_WEBAPP_DEPLOYMENT_ID acima é do dono). No primeiro acesso cria a
 * planilha da pessoa no Drive dela (onboarding) — todo o resto do app lê/grava nela.
 */
function doGetMultiUser_() {
  var email = Session.getActiveUser().getEmail();
  if (!isUserAllowed_(email)) {
    return brandPage_(HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;padding:24px">Este e-mail (' + escapeHtml_(email) +
      ') ainda não tem acesso ao beta do ' + appName_() + '.</p>'
    ));
  }
  ensureUserSpreadsheet_();
  try {
    getSpreadsheet_(); // a planilha guardada ainda abre com esta conta?
  } catch (err) {
    return renderSpreadsheetRecoveryPage_(email);
  }
  return renderIndex_(true);
}

/**
 * A planilha guardada não abre (apagada, sem acesso). Em vez de erro em toda tela, oferece
 * criar uma nova (a antiga fica registrada, não é apagada) ou reconectar a uma que a pessoa
 * já tenha. Nada é criado sozinho: pode ser só uma falha passageira de permissão.
 */
function renderSpreadsheetRecoveryPage_(email) {
  var html =
    '<base target="_top"><meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<style>body{font-family:sans-serif;max-width:480px;margin:24px auto;padding:0 16px}' +
    'button,input{font-size:16px;padding:10px;margin:6px 0;width:100%;box-sizing:border-box}</style>' +
    '<h2>Não consegui abrir sua planilha</h2>' +
    '<p>A planilha de finanças da conta <b>' + escapeHtml_(email) + '</b> foi apagada ou ficou sem acesso. ' +
    'Se foi só uma falha passageira, <a href="" onclick="location.reload();return false">tente de novo</a>.</p>' +
    '<button id="novo">Criar uma planilha nova</button>' +
    '<p>Ou reconecte a uma planilha que você já tem (código ou link):</p>' +
    '<input id="alvo" placeholder="Código ou link da planilha">' +
    '<button id="reconectar">Reconectar</button>' +
    '<p id="msg" style="color:#b00"></p>' +
    '<script>' +
    'function pronto(r){document.getElementById("msg").style.color="#070";' +
    'document.getElementById("msg").innerHTML=r&&r.url?"Pronto. <a href=\\""+r.url+"\\" target=\\"_top\\">Abrir o app</a>":"Pronto. Recarregue a página.";}' +
    'function erro(e){document.getElementById("msg").style.color="#b00";document.getElementById("msg").textContent=(e&&e.message)||String(e);}' +
    'document.getElementById("novo").onclick=function(){google.script.run.withSuccessHandler(pronto).withFailureHandler(erro).recreateUserSpreadsheet();};' +
    'document.getElementById("reconectar").onclick=function(){google.script.run.withSuccessHandler(pronto).withFailureHandler(erro).reconnectUserSpreadsheet(document.getElementById("alvo").value);};' +
    '</script>';
  return brandPage_(HtmlService.createHtmlOutput(html));
}

function renderIndex_(multiuser) {
  var template = HtmlService.createTemplateFromFile('Index');
  template.multiuser = !!multiuser;
  return brandPage_(template.evaluate())
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Camada de acesso do beta: lista opcional de e-mails liberados (ex: quem pagou) em
 * Script Properties > ALLOWED_EMAILS, separados por vírgula. Sem a propriedade, todo
 * mundo entra. É só uma porta na frente: o código do projeto é público por link (precisa
 * ser, pros gatilhos de cada usuário rodarem), então isto NÃO é segurança de dados — os
 * dados ficam na planilha de cada pessoa, no Drive dela.
 */
function isUserAllowed_(email) {
  var raw = PropertiesService.getScriptProperties().getProperty('ALLOWED_EMAILS');
  if (!raw) return true;
  var allowed = raw.split(',').map(function (e) { return e.trim().toLowerCase(); }).filter(Boolean);
  return allowed.indexOf(String(email || '').trim().toLowerCase()) !== -1;
}

function escapeHtml_(text) {
  return String(text || '').replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

/**
 * Conta e planilha da pessoa. Usado pela janela de configuração e pela ponte com o app
 * Android, que grava as notificações direto nesta planilha (Sheets API, com a conta Google
 * da pessoa) e por isso precisa do ID dela.
 */
function getPhoneConnectionInfo() {
  var ss = getSpreadsheet_();
  return {
    email: Session.getActiveUser().getEmail(),
    sheetId: ss.getId(),
    sheetUrl: ss.getUrl(),
    multiuser: !!MULTIUSER_MODE
  };
}

/**
 * Lançamentos em CSV (separador ";" e vírgula decimal, que é o que o Excel/Planilhas em
 * português esperam), com conta e categoria pelo NOME. Texto que começa com = + - @ ganha um
 * apóstrofo na frente pra nunca virar fórmula ao abrir o arquivo.
 */
function exportTransactionsCsv() {
  var accountNames = {};
  readAll_('Accounts').forEach(function (a) { accountNames[a.ID] = a.Name; });
  var categoryNames = {};
  readAll_('Categories').forEach(function (c) { categoryNames[c.ID] = c.Name; });
  var kindLabels = {
    income: 'Receita', expense: 'Despesa', transfer_in: 'Transferência (entrada)', transfer_out: 'Transferência (saída)'
  };

  function text_(value) {
    var s = String(value == null ? '' : value);
    if (/^[=+\-@]/.test(s)) s = "'" + s;
    return '"' + s.replace(/"/g, '""') + '"';
  }
  var lines = [['Data', 'Hora', 'Conta', 'Categoria', 'Tipo', 'Valor', 'Referente a', 'Pago a/Recebido de']
    .map(text_).join(';')];
  readAll_('Transactions').forEach(function (t) {
    lines.push([
      text_(t.Date), text_(t.Time), text_(accountNames[t.AccountID] || ''),
      text_(categoryNames[t.CategoryID] || ''), text_(kindLabels[t.Kind] || t.Kind),
      String(Number(t.Amount) || 0).replace('.', ','),
      text_(t.Description), text_(t.Counterparty)
    ].join(';'));
  });
  return lines.join('\r\n');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Carga inicial usada pelo frontend ao abrir o app. Inclui também a lista completa
 * de transações, para que a aba Extrato abra na hora (filtrada no navegador) em vez
 * de fazer uma chamada nova ao servidor toda vez. Cada parte é buscada de forma
 * independente: uma falha ao agregar uma delas não deve impedir as demais de carregar.
 */
function getBootstrapData() {
  var errors = [];
  var allTxs = [];
  try {
    allTxs = readAll_('Transactions');
  } catch (err) {
    allTxs = [];
    errors.push('lançamentos: ' + err.message);
  }
  var accounts = [];
  try {
    accounts = listAccounts(false, allTxs);
  } catch (err) {
    accounts = [];
    errors.push('contas: ' + err.message);
  }
  var categories = [];
  try {
    categories = listCategories(false, allTxs);
  } catch (err) {
    categories = [];
    errors.push('categorias: ' + err.message);
  }
  var dashboard = null;
  try {
    dashboard = getDashboardData(allTxs);
  } catch (err) {
    dashboard = null;
    errors.push('painel: ' + err.message);
  }
  var usage = { accounts: {}, categories: {} };
  try {
    usage = getUsageStats(allTxs);
  } catch (err) {
    usage = { accounts: {}, categories: {} };
    errors.push('estatísticas de uso: ' + err.message);
  }
  var transactions = [];
  try {
    transactions = listTransactions({}, allTxs);
  } catch (err) {
    transactions = [];
    errors.push('extrato: ' + err.message);
  }
  var pendingEntries = [];
  try {
    pendingEntries = listPendingEntries();
  } catch (err) {
    pendingEntries = [];
    errors.push('pendências: ' + err.message);
  }
  var forecasts = [];
  try {
    forecasts = listForecasts(false);
  } catch (err) {
    forecasts = [];
    errors.push('lançamentos recorrentes: ' + err.message);
  }
  var budgets = [];
  try {
    budgets = listBudgets(false);
  } catch (err) {
    budgets = [];
    errors.push('orçamentos: ' + err.message);
  }
  // Categoria lembrada por contraparte (ver Inbox.js) — exposta pro cliente sugerir
  // categoria em tempo real no Novo lançamento também, não só nas pendências.
  var counterpartyMappings = [];
  try {
    counterpartyMappings = listCounterpartyMappings();
  } catch (err) {
    counterpartyMappings = [];
    errors.push('memória de contraparte: ' + err.message);
  }
  // Nome de estabelecimento lembrado pro texto cru de "Pago a"/"Recebido de" (ver
  // Inbox.js) — exposto pro cliente também reconhecer o nome lembrado (não só o cru) ao
  // sugerir vincular um lançamento a um Lançamento recorrente por nome, ver
  // forecastCounterpartyNames_/JavaScript.html, grillado com o usuário.
  var counterpartyAliases = [];
  try {
    counterpartyAliases = listCounterpartyAliases();
  } catch (err) {
    counterpartyAliases = [];
    errors.push('nomes lembrados: ' + err.message);
  }
  // Descartes de sugestão de orçamento ainda válidos (ver computeBudgetSuggestions_,
  // JavaScript.html, e listBudgetSuggestionDismissals, Budgets.js).
  var budgetSuggestionDismissals = [];
  try {
    budgetSuggestionDismissals = listBudgetSuggestionDismissals();
  } catch (err) {
    budgetSuggestionDismissals = [];
    errors.push('descartes de sugestão: ' + err.message);
  }
  // Planos (ver Plans.js) — eventos únicos e futuros usados pelo gráfico de tendência
  // de saldo (computeBalanceTrend_, JavaScript.html).
  var plans = [];
  try {
    plans = listPlans();
  } catch (err) {
    plans = [];
    errors.push('planos: ' + err.message);
  }
  // Reembolsos (ver Refunds.js) — cru, o cliente deriva líquido/completo a partir
  // disso + allTransactions (mesma convenção de forecastEffectiveAmount).
  var people = [];
  try {
    people = listPeople(false);
  } catch (err) {
    people = [];
    errors.push('pessoas: ' + err.message);
  }
  var refundShares = [];
  try {
    refundShares = listRefundShares();
  } catch (err) {
    refundShares = [];
    errors.push('partes de reembolso: ' + err.message);
  }
  var refundLinks = [];
  try {
    refundLinks = listRefundLinks();
  } catch (err) {
    refundLinks = [];
    errors.push('vínculos de reembolso: ' + err.message);
  }
  // "Ignorar mês" de ciclos não pagos de Lançamentos recorrentes (ver
  // computeForecastCycleStatus_/JavaScript.html e listForecastCycleIgnores/Forecasts.js).
  var forecastCycleIgnores = [];
  try {
    forecastCycleIgnores = listForecastCycleIgnores();
  } catch (err) {
    forecastCycleIgnores = [];
    errors.push('ciclos ignorados de previsão: ' + err.message);
  }
  // Descartes ainda válidos da "Sugestão: R$X Aplicar" de Lançamento recorrente (ver
  // dismissForecastSuggestion/listForecastSuggestionDismissals, Forecasts.js).
  var forecastSuggestionDismissals = [];
  try {
    forecastSuggestionDismissals = listForecastSuggestionDismissals();
  } catch (err) {
    forecastSuggestionDismissals = [];
    errors.push('descartes de sugestão de previsão: ' + err.message);
  }
  // Meta de saldo do gráfico de projeção (ver getBalanceGoal, Plans.js) -- número em R$
  // ou null.
  var balanceGoal = null;
  try {
    balanceGoal = getBalanceGoal();
  } catch (err) {
    balanceGoal = null;
    errors.push('meta de saldo: ' + err.message);
  }
  // Filtro de abertura do card de Orçamentos ('rolling' ou 'current', ver
  // getBudgetViewPreference) -- preferência de UI, não dado financeiro.
  var budgetViewPreference = 'rolling';
  try {
    budgetViewPreference = getBudgetViewPreference();
  } catch (err) {
    budgetViewPreference = 'rolling';
    errors.push('filtro de orçamentos: ' + err.message);
  }
  // Rede de segurança: google.script.run devolve null pro cliente se QUALQUER ponto do
  // retorno tiver um Date (ou outro valor não serializável) -- e aí o app inteiro falha
  // com "Cannot read properties of null (reading 'errors')". Passar por JSON converte
  // Date em texto ISO em vez de derrubar tudo (ver isoCell_, SheetService.js).
  return JSON.parse(JSON.stringify({
    accounts: accounts, categories: categories, dashboard: dashboard, usage: usage,
    transactions: transactions, pendingEntries: pendingEntries, forecasts: forecasts, budgets: budgets,
    counterpartyMappings: counterpartyMappings, counterpartyAliases: counterpartyAliases, budgetSuggestionDismissals: budgetSuggestionDismissals,
    plans: plans, people: people, refundShares: refundShares, refundLinks: refundLinks,
    forecastCycleIgnores: forecastCycleIgnores, forecastSuggestionDismissals: forecastSuggestionDismissals,
    balanceGoal: balanceGoal, budgetViewPreference: budgetViewPreference,
    // Seções que falharam ao carregar. Cada try/catch acima devolve vazio pra uma falha
    // não derrubar as outras — mas devolver vazio EM SILÊNCIO já enganou de um jeito
    // caro: quando readAll_('Transactions') estourava, todo saldo de conta passava a
    // mostrar só o InitialBalance e o extrato ficava vazio, com cara de "o app perdeu
    // meus dados" em vez de "uma leitura falhou". O cliente avisa (ver applyBootstrapData).
    errors: errors
  }));
}

/**
 * Filtro com que o card de Orçamentos (Início) ABRE: 'rolling' ("Últimos dias", o padrão)
 * ou 'current' (o mês atual, seja qual for -- por isso não guarda "2026-09", que viraria
 * um mês passado no mês seguinte). Um mês passado nunca é candidato a abertura, então o
 * cliente simplesmente não grava nada quando o usuário escolhe um (ver
 * budgetViewPreferenceToSave_, JavaScript.html).
 *
 * Fica nas propriedades do USUÁRIO (PropertiesService), não numa aba da planilha nem em
 * localStorage: é preferência de interface, não dado financeiro; vive no servidor, então
 * vale nos dois aparelhos (o localStorage do iframe do Apps Script não é confiável pra
 * isso) e lê em milissegundos no bootstrap, sem mais uma leitura de aba.
 */
var BUDGET_VIEW_PREFERENCE_KEY_ = 'BUDGET_VIEW_PREFERENCE';
function getBudgetViewPreference() {
  var value = PropertiesService.getUserProperties().getProperty(BUDGET_VIEW_PREFERENCE_KEY_);
  return value === 'current' ? 'current' : 'rolling';
}
function setBudgetViewPreference(value) {
  if (value !== 'rolling' && value !== 'current') throw new Error('Filtro de orçamentos inválido.');
  PropertiesService.getUserProperties().setProperty(BUDGET_VIEW_PREFERENCE_KEY_, value);
  return value;
}

// ===== Dashboard.js =====
/** Agregações para a tela inicial. */

function currentYm_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
}

/** Nome da categoria de topo (para agrupar gastos de subcategorias no gráfico). */
function topLevelCategoryName_(categoryId, categoriesById) {
  var cat = categoriesById[categoryId];
  if (!cat) return 'Sem categoria';
  if (!cat.ParentID) return cat.Name;
  var parent = categoriesById[cat.ParentID];
  return parent ? parent.Name : cat.Name;
}

function getDashboardData(allTransactions) {
  var allTxs = allTransactions || readAll_('Transactions');
  var accounts = listAccounts(false, allTxs);
  var totalBalance = accounts.reduce(function (sum, a) { return sum + a.balance; }, 0);

  var categories = readAll_('Categories');
  var categoriesById = {};
  categories.forEach(function (c) { categoriesById[c.ID] = c; });

  var ym = currentYm_();
  // Sem os lançamentos marcados "Não considerar em análises?" (pedido explícito do usuário): a
  // rosca "Gastos por categoria" e os totais do mês mostram só o gasto/receita "normal". O saldo
  // das contas (acima) continua contando TUDO -- por isso essas somas deixam de bater com ele.
  var monthTxs = allTxs.filter(function (t) { return t.Date && t.Date.indexOf(ym) === 0 && !parseBoolCell_(t.ExcludeFromAnalysis); });

  var expensesByCategory = {};
  monthTxs.forEach(function (t) {
    if (t.Kind !== 'expense') return;
    var name = topLevelCategoryName_(t.CategoryID, categoriesById);
    expensesByCategory[name] = (expensesByCategory[name] || 0) + Number(t.Amount);
  });

  var totalIncomeMonth = monthTxs.reduce(function (sum, t) { return t.Kind === 'income' ? sum + Number(t.Amount) : sum; }, 0);
  var totalExpenseMonth = monthTxs.reduce(function (sum, t) { return t.Kind === 'expense' ? sum + Number(t.Amount) : sum; }, 0);

  var recentTransactions = listTransactions({}, allTxs).slice(0, 15);

  var creditCardSummaries = accounts
    .filter(function (a) { return a.type === 'credit_card'; })
    .map(function (a) {
      var bills = getCreditCardBillSummary(a.id, allTxs);
      return { accountId: a.id, accountName: a.name, bills: bills };
    });

  return {
    totalBalance: totalBalance,
    accounts: accounts,
    expensesByCategory: Object.keys(expensesByCategory).map(function (name) {
      return { category: name, amount: expensesByCategory[name] };
    }),
    totalIncomeMonth: totalIncomeMonth,
    totalExpenseMonth: totalExpenseMonth,
    recentTransactions: recentTransactions,
    creditCardSummaries: creditCardSummaries,
    month: ym
  };
}

/**
 * Conta quantas vezes cada conta/categoria apareceu em lançamentos registrados
 * (CreatedAt) nos últimos 30 dias — usado para ordenar as cápsulas de seleção
 * pelas opções mais usadas recentemente.
 */
function getUsageStats(allTransactions) {
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);

  var accounts = {};
  var categories = {};
  (allTransactions || readAll_('Transactions')).forEach(function (t) {
    if (!t.CreatedAt) return;
    var createdAt = new Date(t.CreatedAt);
    if (isNaN(createdAt.getTime()) || createdAt < cutoff) return;
    if (t.AccountID) accounts[t.AccountID] = (accounts[t.AccountID] || 0) + 1;
    if (t.CategoryID) categories[t.CategoryID] = (categories[t.CategoryID] || 0) + 1;
  });
  return { accounts: accounts, categories: categories };
}

// ===== Forecasts.js =====
/**
 * CRUD de Lançamentos recorrentes: um gasto/receita que se repete a cada IntervalDays
 * dias (ex: "Diarista" a cada 7, "Salário" a cada 30), mas que não faz sentido modelar
 * como um limite de Orçamento. Pode ser numa categoria OU subcategoria, e vale tanto
 * pra despesa quanto pra receita. Repete automaticamente com o mesmo ExpectedAmount até
 * ser editado ou removido.
 *
 * O vínculo entre um Lançamento recorrente e o lançamento real de uma ocorrência
 * específica fica no campo ForecastID da própria Transaction (ver Transactions.js) —
 * por isso, se ele "foi cumprido" (dentro da janela dos últimos IntervalDays dias) e
 * com qual valor real é algo calculado no cliente (que já tem STATE.allTransactions
 * carregado), não aqui.
 */

/**
 * 'days' reaproveita a mesma faixa de Orçamentos (1-999, ver validateIntervalDays_ em
 * Budgets.js); 'months' tem faixa própria bem mais estreita (1-60 = até "a cada 5 anos",
 * já além do que faz sentido pra um lançamento recorrente) — são unidades diferentes, não
 * dá pra validar as duas com o mesmo limite numérico.
 */
function validateForecastInterval_(value, unit) {
  if (unit === 'months') {
    var m = Number(value);
    if (!Number.isInteger(m) || m < 1 || m > 60) throw new Error('Intervalo precisa ser um número inteiro entre 1 e 60 meses.');
    return { days: m, unit: 'months' };
  }
  return { days: validateIntervalDays_(value), unit: 'days' };
}

/** Só existe pra período Mensal (é o cliente quem só manda isso quando IntervalUnit é
 *  'months', ver openForecastModal) — puramente informativo (um lembrete de QUANDO no
 *  mês esperar, ver forecastRowHtml), nunca muda a janela de "cumprido" (que continua
 *  sempre rolante até hoje, ver forecastLinkedTransactionInPeriod_). '' (vazio) = não
 *  informado, aceito de volta pra permitir limpar um valor já salvo. */
function validateForecastDayOfMonth_(value) {
  if (value === undefined || value === null || value === '') return '';
  var n = Number(value);
  if (!Number.isInteger(n) || n < 1 || n > 31) throw new Error('Dia do mês precisa ser um número inteiro entre 1 e 31.');
  return n;
}

function listForecasts(includeArchived) {
  var categoriesById = {};
  readAll_('Categories').forEach(function (c) { categoriesById[c.ID] = c; });

  return readAll_('MonthlyForecasts')
    .filter(function (f) { return includeArchived || !isArchived_(f.Archived); })
    .map(function (f) {
      var cat = categoriesById[f.CategoryID];
      return {
        id: f.ID,
        kind: f.Kind,
        categoryId: f.CategoryID,
        categoryName: cat ? cat.Name : 'Categoria removida',
        label: f.Label,
        expectedAmount: Number(f.ExpectedAmount),
        // Linhas antigas sem IntervalDays equivalem a "30 dias" — mesma convenção de Budgets.
        intervalDays: f.IntervalDays ? Number(f.IntervalDays) : 30,
        intervalUnit: f.IntervalUnit === 'months' ? 'months' : 'days',
        dayOfMonth: f.DayOfMonth ? Number(f.DayOfMonth) : null,
        archived: isArchived_(f.Archived),
        // Usado pra limitar quantos ciclos passados a linha de bolinhas de status
        // mostra -- não faz sentido marcar "Atrasado" um mês anterior à própria
        // criação do lançamento recorrente (ver forecastCycleWindows_/JavaScript.html).
        createdAt: f.CreatedAt || null
      };
    });
}

function createForecast(input) {
  if (!input.categoryId) throw new Error('Categoria é obrigatória.');
  if (!findById_('Categories', input.categoryId)) throw new Error('Categoria não encontrada.');
  if (!input.label) throw new Error('Nome é obrigatório.');
  var amount = Number(input.expectedAmount);
  if (!amount || amount <= 0) throw new Error('Valor esperado precisa ser maior que zero.');
  var interval = validateForecastInterval_(input.intervalDays, input.intervalUnit);
  var dayOfMonth = validateForecastDayOfMonth_(input.dayOfMonth);

  return appendRow_('MonthlyForecasts', {
    Kind: input.kind, CategoryID: input.categoryId, Label: input.label,
    ExpectedAmount: amount, IntervalDays: interval.days, IntervalUnit: interval.unit, DayOfMonth: dayOfMonth, Archived: false
  });
}

function updateForecast(input) {
  if (!input.id) throw new Error('ID da previsão é obrigatório.');
  var updates = {};
  if (input.categoryId !== undefined) {
    if (!input.categoryId) throw new Error('Categoria é obrigatória.');
    if (!findById_('Categories', input.categoryId)) throw new Error('Categoria não encontrada.');
    updates.CategoryID = input.categoryId;
  }
  if (input.label !== undefined) {
    if (!input.label) throw new Error('Nome é obrigatório.');
    updates.Label = input.label;
  }
  if (input.expectedAmount !== undefined) {
    var amount = Number(input.expectedAmount);
    if (!amount || amount <= 0) throw new Error('Valor esperado precisa ser maior que zero.');
    updates.ExpectedAmount = amount;
  }
  if (input.intervalDays !== undefined) {
    var interval = validateForecastInterval_(input.intervalDays, input.intervalUnit);
    updates.IntervalDays = interval.days;
    updates.IntervalUnit = interval.unit;
  }
  if (input.dayOfMonth !== undefined) updates.DayOfMonth = validateForecastDayOfMonth_(input.dayOfMonth);
  if (input.archived !== undefined) updates.Archived = input.archived;
  var updated = updateRowById_('MonthlyForecasts', input.id, updates);
  if (!updated) throw new Error('Previsão não encontrada.');
  return updated;
}

function removeForecast(id) {
  var removed = deleteRowById_('MonthlyForecasts', id);
  if (!removed) throw new Error('Previsão não encontrada.');
  return true;
}

/** Descartes ainda válidos de ciclo de Lançamento recorrente (ver "Ignorar mês",
 *  computeForecastCycleStatus_/JavaScript.html) -- forma mínima pro cliente (não expõe
 *  ID/IgnoredAt, que não importam pra decidir se uma bolinha está ignorada). */
function listForecastCycleIgnores() {
  return readAll_('ForecastCycleIgnores').map(function (x) {
    return { forecastId: x.ForecastID, cycleStart: x.CycleStart };
  });
}

/** Marca um ciclo (mês/período) de um Lançamento recorrente como "ignorado" -- upsert
 *  por (ForecastID, CycleStart), mesmo padrão de dismissBudgetSuggestion/Budgets.js.
 *  Idempotente: marcar de novo um ciclo já ignorado não duplica linha. */
function ignoreForecastCycle(input) {
  if (!input.forecastId) throw new Error('ID do lançamento recorrente é obrigatório.');
  if (!input.cycleStart) throw new Error('Início do ciclo é obrigatório.');
  var existing = readAll_('ForecastCycleIgnores').find(function (x) {
    return x.ForecastID === input.forecastId && x.CycleStart === input.cycleStart;
  });
  if (!existing) {
    appendRow_('ForecastCycleIgnores', { ForecastID: input.forecastId, CycleStart: input.cycleStart, IgnoredAt: nowIso_() });
  }
  return true;
}

/** Desfaz "Ignorar mês" (ver ignoreForecastCycle acima) -- reversível, pedido explícito
 *  do usuário. Sem efeito (não é erro) se o ciclo já não estava ignorado. */
function unignoreForecastCycle(input) {
  if (!input.forecastId) throw new Error('ID do lançamento recorrente é obrigatório.');
  if (!input.cycleStart) throw new Error('Início do ciclo é obrigatório.');
  var existing = readAll_('ForecastCycleIgnores').find(function (x) {
    return x.ForecastID === input.forecastId && x.CycleStart === input.cycleStart;
  });
  if (existing) deleteRowById_('ForecastCycleIgnores', existing.ID);
  return true;
}

var FORECAST_SUGGESTION_DISMISS_DAYS = 30;

/** Descarta a sugestão de novo valor de um Lançamento recorrente (ver
 *  forecastSuggestedAmount/JavaScript.html) por FORECAST_SUGGESTION_DISMISS_DAYS dias --
 *  mesmo padrão de dismissBudgetSuggestion/Budgets.js. Upsert por ForecastID: descartar
 *  de novo só atualiza a data, não acumula linha. */
function dismissForecastSuggestion(forecastId) {
  if (!forecastId) throw new Error('ID do lançamento recorrente é obrigatório.');
  var existing = readAll_('ForecastSuggestionDismissals').find(function (d) { return d.ForecastID === forecastId; });
  if (existing) {
    updateRowById_('ForecastSuggestionDismissals', existing.ID, { DismissedAt: nowIso_() });
  } else {
    appendRow_('ForecastSuggestionDismissals', { DismissedAt: nowIso_(), ForecastID: forecastId });
  }
  return true;
}

/** Expõe os descartes ainda VÁLIDOS (dentro dos 30 dias) pro cliente decidir quais
 *  sugestões não devem aparecer de novo — ver forecastRowHtml/JavaScript.html. Purga
 *  sozinha (autocorretiva, mesmo padrão de listBudgetSuggestionDismissals) os descartes
 *  mais antigos que isso. */
function listForecastSuggestionDismissals() {
  var cutoff = new Date(Date.now() - FORECAST_SUGGESTION_DISMISS_DAYS * 24 * 60 * 60 * 1000);
  var rows = readAll_('ForecastSuggestionDismissals');
  var stillValid = rows.filter(function (d) { return new Date(d.DismissedAt) >= cutoff; });
  if (stillValid.length !== rows.length) {
    withSheetLock_(function () {
      var sheet = getOrCreateSheet_('ForecastSuggestionDismissals');
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      var headers = SCHEMA.ForecastSuggestionDismissals;
      sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
      if (stillValid.length) {
        sheet.getRange(2, 1, stillValid.length, headers.length).setValues(
          stillValid.map(function (row) { return objectToRow_(headers, row); })
        );
      }
    });
  }
  return stillValid.map(function (d) { return { forecastId: d.ForecastID, dismissedAt: isoCell_(d.DismissedAt) }; });
}

// ===== Inbox.js =====
/**
 * Caixa de pendências: notificações bancárias capturadas pelo MacroDroid (ver doPost em
 * Code.js) e por email (ver EmailInbox.js) esperando revisão antes de virarem uma
 * Transaction de verdade. Tenta adivinhar valor e tipo (receita/despesa) do texto da
 * notificação; quando não reconhece o padrão, deixa em branco pro usuário preencher.
 *
 * O MacroDroid grava só o texto cru (ver captureRawNotification_) — nenhuma interpretação
 * de conteúdo roda no caminho do doPost, de propósito, porque esse webhook só existe atrás
 * de um deployment FIXO (a URL que o celular chama), preso na versão do dia em que foi
 * criado/atualizado pela última vez. Toda a extração de valor/tipo/contraparte/categoria e
 * o filtro de spam/duplicata rodam em listPendingEntries/purgeJunkPendingEntries_ (ver
 * deriveNotificationFields_), sob o deployment principal — esse sim sempre atualizado.
 * Email e importação de documento continuam gravando estruturado desde a criação (ver
 * buildPendingEntryRow_) — não sofrem desse problema, já que rodam via trigger/deployment
 * principal, não atrás de uma URL fixa externa.
 */

// Notificações cujo texto contém qualquer uma dessas frases são ignoradas por
// completo (não viram pendência) — avisos/lembretes/promoções, não transações de
// verdade. Compartilhado entre o webhook do MacroDroid e a leitura de email.
var IGNORED_PENDING_PHRASES = [
  'ainda falta pagar aquele último boleto do mês',
  'bônus de 50% do cdi',
  'acelerador de lucros',
  'um cashback de',
  'tarifa atual',
  'não esqueça seus objetos pessoais',
  'pagamento de boleto não finalizada',
  'aprovação de crédito',
  'gravação automática',
  'seu boleto vence amanhã',
  'aviso transferência',
  'localizando motoristas',
  'receba de qualquer cartão',
  'o valor já virou limite',
  'você pode ganhar',
  'para buscar você',
  'seu saldo pode render',
  'mais agilidade pros seus envios',
  'mais fôlego pra fazer pix',
  'você tem um novo boleto',
  'lembrete de transferência futura',
  'parcele seu boleto'
];

// Corrida 99 cancelada: o 99Pay debita a corrida e, se cancelada, reembolsa em seguida
// com uma notificação separada contendo essa frase. Quando o valor bate, as duas
// pendências (débito da corrida + reembolso) se cancelam e nenhuma delas deve aparecer
// pra revisão — ver checkAndRemoveRefundPair_ em createPendingEntry.
var CORRIDA_99_TEXT = 'corrida 99';
var REEMBOLSO_99_TEXT = 'um valor foi adicionado';

function shouldIgnorePendingText_(text) {
  var lower = (text || '').toLowerCase();
  return IGNORED_PENDING_PHRASES.some(function (phrase) { return lower.indexOf(phrase.toLowerCase()) !== -1; });
}

/** "99" (o app) aparece no App de qualquer jeito que a notificação tenha chegado --
 *  pelo MacroDroid (nome que o celular manda) ou por email (ver EmailInbox.js: app =
 *  'email:99pay') -- por isso checa SUBSTRING, não igualdade exata. */
function isNinetyNineApp_(appIdentifier) {
  // "99" puro na célula App vira o NÚMERO 99 na leitura do Sheets (autoconversão de
  // texto puramente numérico) -- String(...) garante que .toLowerCase() nunca quebra,
  // independente de a célula ter sido lida como string ou como número (bug relatado:
  // "(appIdentifier || "").toLowerCase is not a function").
  return String(appIdentifier || '').toLowerCase().indexOf('99') !== -1;
}

/**
 * Grava o texto cru de uma notificação do 99 assim que ela vira lançamento confirmado
 * de verdade (ver confirmPendingEntry) -- só se ainda não existir uma linha com o MESMO
 * texto (dedupe por Text, não por notificação individual: o objetivo é catalogar os
 * FORMATOS diferentes que já apareceram, não contar repetições). Substitui a estratégia
 * anterior de só ir crescendo IGNORED_PENDING_PHRASES adivinhando o que NÃO deveria
 * entrar -- pedido explícito do usuário: olhando essa aba, ele mesmo identifica o
 * padrão do que É uma notificação legítima do 99, sem precisar de mais nenhuma frase
 * nova na lista de ignoradas.
 */
function recordNinetyNineConfirmedSample_(pending) {
  if (!isNinetyNineApp_(pending.App) || !pending.Text) return;
  var already = readAll_('Notification99Samples').some(function (r) { return r.Text === pending.Text; });
  if (already) return;
  appendRow_('Notification99Samples', { Text: pending.Text, App: pending.App, Title: pending.Title || '' });
}

/** Extrai "R$ 1.234,56" -> 1234.56. Retorna null se não achar nenhum valor no texto. */
function extractAmountFromText_(text) {
  var match = (text || '').match(/R\$\s*([\d.,]+)/);
  if (!match) return null;
  var normalized = match[1].replace(/\./g, '').replace(',', '.');
  var value = parseFloat(normalized);
  return isNaN(value) ? null : value;
}

/** Palpite de receita/despesa com base em palavras comuns de notificação bancária. */
function guessKindFromText_(text) {
  var lower = (text || '').toLowerCase();
  var incomeHints = ['recebeu', 'recebido', 'depósito', 'deposito', 'creditado', 'estorno', 'estornado'];
  var expenseHints = ['pagamento', 'pagou', 'compra', 'enviou', 'enviado', 'debitado', 'saque', 'fatura'];
  for (var i = 0; i < incomeHints.length; i++) if (lower.indexOf(incomeHints[i]) !== -1) return 'income';
  for (var i2 = 0; i2 < expenseHints.length; i2++) if (lower.indexOf(expenseHints[i2]) !== -1) return 'expense';
  return '';
}

function counterpartyKey_(name) {
  return (name || '').trim().toLowerCase();
}

/**
 * Notificações de compra cruas (ex: RecargaPay) costumam vir como "Em WL STEAM P com o
 * Cartão final 7072" — texto inteiro em caixa alta, com um preâmbulo/sufixo que não é
 * nome de estabelecimento nenhum. Extrai só o meio da frase quando esse padrão aparece
 * (chamado só pelo doPost/webhook do MacroDroid — ver Code.js; fontes que já sabem a
 * contraparte, como email e importação de documento, nunca passam por aqui). Quando o
 * padrão não bate, devolve o texto como veio — melhor não mexer num formato desconhecido
 * do que arriscar estragar uma contraparte que já esteja bem formatada.
 */
function guessCounterpartyFromNotificationText_(text) {
  if (!text) return '';
  var match = text.match(/\bem\s+(.+?)\s+com\s+o\s+cart[ãa]o\s+final\s+\d+\b/i);
  if (!match) return text.trim();
  return titleCasePt_(match[1].trim());
}

/** "WL STEAM P" -> "Wl Steam P" (capitaliza só a primeira letra de cada palavra). */
function titleCasePt_(text) {
  return text.toLowerCase().replace(/(^|\s)(\S)/g, function (whole, sep, ch) { return sep + ch.toUpperCase(); });
}

/**
 * Grava a notificação do MacroDroid CRUA na planilha — sem nenhuma interpretação de
 * conteúdo (nem valor, nem tipo, nem contraparte, nem categoria). De propósito: o doPost
 * (Code.js) só existe atrás de um deployment FIXO — a URL que o MacroDroid chama no
 * celular — preso pra sempre na versão do dia em que foi criado/atualizado pela última
 * vez. Deixar aqui só a função mais simples possível (capturar o texto) significa que ela
 * nunca mais precisa de deploy de novo, mesmo quando o resto do app evolui — diferente da
 * limpeza de "Pago a", que ficou presa numa versão antiga por estar no caminho errado.
 * Toda a extração de valor/tipo/contraparte/categoria roda em listPendingEntries/
 * purgeJunkPendingEntries_ (ver deriveNotificationFields_ abaixo), sob o deployment
 * principal — esse sim atualizado a cada deploy.
 */
function captureRawNotification_(app, title, text) {
  return appendRow_('PendingEntries', { App: app, Title: title, Text: text });
}

/**
 * Deriva valor/tipo/contraparte de uma pendência a partir do texto bruto — usado quando a
 * fonte não gravou isso estruturado na criação (é sempre o caso do MacroDroid agora, ver
 * captureRawNotification_ acima). Fontes que já sabem esses valores com certeza (email,
 * importação de documento) gravam estruturado desde a criação — aqui só entra em ação
 * quando falta, então nunca reprocessa o que já veio confiável.
 */
function deriveNotificationFields_(row) {
  var combinedText = ((row.Title || '') + ' ' + (row.Text || '')).trim();
  var amount = row.SuggestedAmount !== '' ? Number(row.SuggestedAmount) : extractAmountFromText_(combinedText);
  var kind = row.SuggestedKind || guessKindFromText_(combinedText);
  var counterparty = row.Counterparty || guessCounterpartyFromNotificationText_(combinedText);
  return { combinedText: combinedText, amount: amount || null, kind: kind || '', counterparty: counterparty || '' };
}

/**
 * Categoria lembrada pra essa pessoa/estabelecimento, se alguma vez foi salva.
 * mappingsCache (opcional): evita reler CounterpartyMappings a cada pendência da lista
 * (ver listPendingEntries) — mesmo motivo de transactionsCache/pendingEntriesCache.
 */
function findCategoryIdForCounterparty_(name, mappingsCache) {
  if (!name) return null;
  var key = counterpartyKey_(name);
  var mappings = mappingsCache || readAll_('CounterpartyMappings');
  var mapping = mappings.find(function (m) { return counterpartyKey_(m.CounterpartyName) === key; });
  return mapping ? mapping.CategoryID : null;
}

/** Expõe as categorias lembradas por contraparte pro cliente (ver findCategoryIdForCounterpartyClient_, JavaScript.html). */
function listCounterpartyMappings() {
  return readAll_('CounterpartyMappings').map(function (m) {
    return { counterparty: m.CounterpartyName, categoryId: m.CategoryID };
  });
}

function saveCounterpartyMapping_(name, categoryId) {
  if (!name || !categoryId) return;
  var key = counterpartyKey_(name);
  var existing = readAll_('CounterpartyMappings').find(function (m) { return counterpartyKey_(m.CounterpartyName) === key; });
  if (existing) {
    updateRowById_('CounterpartyMappings', existing.ID, { CategoryID: categoryId });
  } else {
    appendRow_('CounterpartyMappings', { CounterpartyName: name.trim(), CategoryID: categoryId });
  }
}

/**
 * Nome de estabelecimento lembrado pro texto CRU de "Pago a"/"Recebido de" que chega
 * automático (ex: "Comerciald" -> "Planeta Doce") — mesmo princípio de
 * findCategoryIdForCounterparty_, só que pra REESCREVER o texto sugerido, não escolher
 * categoria. aliasesCache (opcional): evita reler CounterpartyAliases a cada pendência
 * da lista (ver listPendingEntries), mesmo motivo de mappingsCache ali.
 */
function findEstablishmentForCounterparty_(name, aliasesCache) {
  if (!name) return null;
  var key = counterpartyKey_(name);
  var aliases = aliasesCache || readAll_('CounterpartyAliases');
  var alias = aliases.find(function (a) { return counterpartyKey_(a.RawText) === key; });
  return alias ? alias.EstablishmentName : null;
}

/** Expõe os nomes lembrados pro cliente (ver forecastCounterpartyNames_/JavaScript.html
 *  -- sugestão de vincular lançamento a recorrente por nome de "Pago a"/"Recebido de"). */
function listCounterpartyAliases() {
  return readAll_('CounterpartyAliases').map(function (a) {
    return { rawText: a.RawText, establishmentName: a.EstablishmentName };
  });
}

function saveCounterpartyAlias_(rawText, establishmentName) {
  if (!rawText || !establishmentName) return;
  var key = counterpartyKey_(rawText);
  var existing = readAll_('CounterpartyAliases').find(function (a) { return counterpartyKey_(a.RawText) === key; });
  if (existing) {
    updateRowById_('CounterpartyAliases', existing.ID, { EstablishmentName: establishmentName.trim() });
  } else {
    appendRow_('CounterpartyAliases', { RawText: rawText.trim(), EstablishmentName: establishmentName.trim() });
  }
}

/**
 * Se essa pendência for uma "Corrida 99" (débito) ou um reembolso ("um valor foi
 * adicionado"), procura entre as pendências JÁ na fila a contraparte (o outro lado do
 * par) com o mesmo valor. Achando, remove a contraparte da planilha e sinaliza pra não
 * criar essa nova também — as duas se cancelam (corrida debitada e depois cancelada,
 * gerando o reembolso).
 */
function findAndClearRefundPair_(combinedText, amount) {
  var lower = combinedText.toLowerCase();
  var isCorrida = lower.indexOf(CORRIDA_99_TEXT) !== -1;
  var isReembolso = lower.indexOf(REEMBOLSO_99_TEXT) !== -1;
  if (!isCorrida && !isReembolso) return false;

  var counterpartPhrase = isCorrida ? REEMBOLSO_99_TEXT : CORRIDA_99_TEXT;
  var counterpart = readAll_('PendingEntries').find(function (row) {
    var rowText = ((row.Title || '') + ' ' + (row.Text || '')).toLowerCase();
    if (rowText.indexOf(counterpartPhrase) === -1) return false;
    var rowAmount = row.SuggestedAmount === '' ? 0 : Number(row.SuggestedAmount);
    return Math.abs(rowAmount - amount) < 0.005;
  });
  if (!counterpart) return false;
  console.error(
    'findAndClearRefundPair_: cancelando par corrida/reembolso — texto atual: "' + combinedText +
    '" (R$' + amount + '), contraparte removida: "' + ((counterpart.Title || '') + ' ' + (counterpart.Text || '')) +
    '" (R$' + counterpart.SuggestedAmount + ', ID ' + counterpart.ID + ')'
  );
  deleteRowById_('PendingEntries', counterpart.ID);
  return true;
}

/**
 * Dedup genérico: evita criar uma pendência pra algo que já virou Transaction de
 * verdade — vale pra qualquer fonte (MacroDroid, email, importação de documento),
 * não só fatura. Casos reais: lançamento manual feito antes da notificação chegar,
 * ou reprocessamento de um email/documento já confirmado antes.
 *
 * Parcelas (installmentSeq/installmentTotal informados) comparam por Data+Parcela,
 * não por valor: quando a parcela 1 de uma compra é confirmada, createTransaction já
 * cria TODAS as parcelas de uma vez (mesma Date original, BillMonth avançando —
 * ver Transactions.js), então a parcela 2 "aparecer" numa fatura do mês seguinte não
 * é uma pendência nova, é a mesma compra já lançada. Valor não entra na comparação
 * porque splitAmount_ pode variar 1-2 centavos entre parcelas por arredondamento.
 *
 * Lançamentos avulsos comparam por Data+Valor (tolerância de 1 centavo).
 *
 * accountId (opcional): quando informado, restringe a busca a essa conta — evita que
 * uma coincidência de Data+Valor em OUTRA conta (ex: uma assinatura de valor comum)
 * esconda uma pendência legítima.
 *
 * transactionsCache (opcional): evita reler a aba Transactions inteira a cada chamada —
 * a importação de documento chama createPendingEntry dezenas de vezes seguidas (uma por
 * linha da fatura); sem isso, cada uma relia a aba inteira do zero (lento em contas com
 * histórico grande, e foi o que causou o cliente perder o rastro de uma confirmação que
 * na prática já tinha terminado no servidor). MacroDroid/email continuam sem passar isso,
 * já que ali é sempre UMA pendência por vez.
 */
function transactionExistsForPending_(date, amount, installmentSeq, installmentTotal, accountId, transactionsCache) {
  if (!date) return false;
  var transactions = (transactionsCache || readAll_('Transactions')).filter(function (t) {
    return !accountId || t.AccountID === accountId;
  });
  if (installmentSeq && installmentTotal) {
    return transactions.some(function (t) {
      return t.Date === date &&
        Number(t.InstallmentSeq) === Number(installmentSeq) &&
        Number(t.InstallmentTotal) === Number(installmentTotal);
    });
  }
  if (!amount) return false;
  return transactions.some(function (t) {
    return t.Date === date && Math.abs(Number(t.Amount) - Number(amount)) < 0.01;
  });
}

/**
 * Dedup contra pendências que JÁ ESTÃO na fila de revisão (diferente de
 * transactionExistsForPending_, que compara contra Transaction CONFIRMADA) — evita
 * pendência duplicada quando o MESMO documento é importado mais de uma vez (reimportar
 * a mesma fatura por engano, ou repetir um teste antes de revisar a fila da vez
 * anterior). Só ativa quando pendingEntriesCache é passado — MacroDroid/email não
 * passam, de propósito: lá já existe markDuplicatePendingEntries_ (aviso visual, não
 * bloqueio), porque duas notificações do mesmo valor no mesmo dia podem ser duas
 * transações reais distintas (ex: duas corridas de mesmo preço). Pra um documento, a
 * data de cada linha é determinística (vem do próprio PDF) — reimportar reproduz a
 * MESMA data pro mesmo item, então Data+Valor é uma comparação segura aqui.
 *
 * Escopado por App (só compara contra pendências da MESMA fonte) e por conta quando
 * disponível, pelos mesmos motivos de transactionExistsForPending_.
 */
function pendingEntryExists_(app, accountId, date, amount, pendingEntriesCache) {
  if (!date || !amount) return false;
  return pendingEntriesCache.some(function (p) {
    if (p.App !== app) return false;
    if (accountId && p.SuggestedAccountID !== accountId) return false;
    var pDate = (p.CreatedAt || '').slice(0, 10);
    return pDate === date && Math.abs(Number(p.SuggestedAmount) - Number(amount)) < 0.01;
  });
}

/**
 * Chamado pela leitura de email (processEmailInbox) e pela importação de documentos
 * (DocumentImport.js) — fontes que sabem o valor/tipo/contraparte com uma estrutura
 * confiável e não sofrem do problema de deployment travado (rodam via trigger/deployment
 * principal, não atrás de uma URL fixa externa). O MacroDroid NÃO passa mais por aqui —
 * ver captureRawNotification_/doPost: grava só o texto cru, e toda essa mesma decisão
 * (frase ignorada, valor, par de corrida/reembolso) roda de novo em
 * purgeJunkPendingEntries_/listPendingEntries quando falta.
 * input.amount/kind/capturedAt/counterparty são opcionais — quando a fonte já sabe o
 * valor/tipo/contraparte com certeza (ex: email do 99Pay, que tem "Valor:"/"Para:"/"De:"
 * explícitos, ou uma linha de fatura já estruturada), passa aqui em vez de deixar pro
 * palpite genérico de regex/palavra-chave usado pra notificações de celular. Retorna null
 * (sem criar nada) se o texto bater com IGNORED_PENDING_PHRASES, se nenhum valor em R$ for
 * identificado (sem valor, quase sempre é propaganda, lembrete ou outra notificação que não
 * é uma transação de verdade), se essa pendência fechar um par corrida cancelada/reembolso
 * (ver findAndClearRefundPair_), ou se já existir uma Transaction correspondente (ver
 * transactionExistsForPending_).
 *
 * input.accountId (opcional): força SuggestedAccountID direto, sem depender de
 * BankAppMappings — usado pela importação de documento, que já sabe a conta de destino.
 * input.toAccountId (opcional): SuggestedToAccountID, pra pendência de transferência (ex:
 * pagamento de fatura, onde o destino — a própria conta de cartão — já é conhecido).
 * input.installments (opcional): total de parcelas (SuggestedInstallments), repassado pro
 * createTransaction ao confirmar. input.installmentSeq (opcional, só influencia o dedup
 * acima): qual parcela esta linha representa — não é armazenado, já que ao confirmar
 * createTransaction sempre recria a série inteira a partir de installments+date.
 */
function createPendingEntry(input) {
  var row = buildPendingEntryRow_(input);
  if (!row) return null;
  return appendRow_('PendingEntries', row);
}

/**
 * Mesma lógica de decisão de createPendingEntry (frases ignoradas, valor, par corrida/
 * reembolso, dedup contra Transaction, palpite de tipo/categoria), só que devolve a
 * LINHA pronta pra gravar em vez de já gravar — permite montar várias de uma vez e
 * gravar todas juntas (ver createPendingEntriesBatch_), sem duplicar essa lógica.
 */
function buildPendingEntryRow_(input) {
  var combinedText = ((input.title || '') + ' ' + (input.text || '')).trim();
  if (shouldIgnorePendingText_(combinedText)) return null;

  var accountMapping = readAll_('BankAppMappings').find(function (m) { return m.AppIdentifier === input.app; });
  var suggestedAmount = (input.amount !== undefined && input.amount !== null) ? input.amount : extractAmountFromText_(combinedText);
  if (!suggestedAmount) return null;
  if (findAndClearRefundPair_(combinedText, suggestedAmount)) return null;

  var resolvedAccountId = input.accountId || (accountMapping ? accountMapping.AccountID : '');
  // Pra dedup de uma pendência de transferência (sem accountId, só toAccountId — ex:
  // pagamento de fatura), o lado que importa comparar é o destino: createTransfer grava
  // a perna "transfer_in" com AccountID = toAccountId (ver Transactions.js).
  var dedupAccountId = resolvedAccountId || input.toAccountId || '';
  var entryDate = input.capturedAt ? input.capturedAt.slice(0, 10) : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (transactionExistsForPending_(entryDate, suggestedAmount, input.installmentSeq, input.installments, dedupAccountId, input.transactionsCache)) return null;
  // Só ativa quando input.pendingEntriesCache é passado (importação de documento) —
  // ver pendingEntryExists_ pro motivo de não valer pra MacroDroid/email.
  if (input.pendingEntriesCache && pendingEntryExists_(input.app, dedupAccountId, entryDate, suggestedAmount, input.pendingEntriesCache)) return null;

  var suggestedKind = input.kind || guessKindFromText_(combinedText);
  var suggestedCategoryId = findCategoryIdForCounterparty_(input.counterparty);

  var row = {
    App: input.app || '',
    Title: input.title || '',
    Text: input.text || '',
    SuggestedAmount: suggestedAmount || '',
    SuggestedKind: suggestedKind || '',
    SuggestedAccountID: resolvedAccountId,
    Counterparty: input.counterparty || '',
    SuggestedCategoryID: suggestedCategoryId || '',
    SourceMessageId: input.sourceMessageId || '',
    SuggestedToAccountID: input.toAccountId || '',
    SuggestedInstallments: input.installments || ''
  };
  if (input.capturedAt) row.CreatedAt = input.capturedAt;
  return row;
}

/**
 * Cria várias pendências de uma vez com UMA escrita só na planilha (ver
 * appendRowsBatch_, SheetService.js) — usado pela importação de documento, que pode
 * ter dezenas de itens num lote. Cada input passa pela MESMA decisão de
 * buildPendingEntryRow_; só a gravação final é diferente.
 */
function createPendingEntriesBatch_(inputs) {
  var rows = [];
  var skipped = 0;
  inputs.forEach(function (input) {
    var row = buildPendingEntryRow_(input);
    if (row) rows.push(row); else skipped++;
  });
  if (rows.length) appendRowsBatch_('PendingEntries', rows);
  return { created: rows.length, skipped: skipped };
}

/**
 * Marca como possível duplicata qualquer pendência que compartilhe valor, texto e
 * horário (minuto) com outra — comum quando o mesmo Pix dispara notificação no
 * celular E email ao mesmo tempo.
 */
function markDuplicatePendingEntries_(entries) {
  var seen = {};
  entries.forEach(function (e) {
    var minuteKey = e.capturedAt ? e.capturedAt.slice(0, 16) : ''; // yyyy-MM-ddTHH:mm
    var key = (e.suggestedAmount || '') + '|' + (e.text || '').trim().toLowerCase() + '|' + minuteKey;
    if (!seen[key]) seen[key] = [];
    seen[key].push(e);
  });
  Object.keys(seen).forEach(function (key) {
    if (seen[key].length > 1) seen[key].forEach(function (e) { e.isDuplicate = true; });
  });
  return entries;
}

/**
 * Roda sozinha a cada listagem — autocorretiva, igual normalizeTextColumns_. Faz DUAS
 * coisas, ambas agora sobre os valores DERIVADOS (ver deriveNotificationFields_), já que
 * o MacroDroid não grava mais nada pronto (só o texto cru — ver captureRawNotification_):
 *   1. Descarta frase ignorada (spam/aviso) ou sem valor em R$ identificável no texto.
 *   2. Cancela pares de corrida 99 debitada + reembolsada (mesmo valor).
 * Repete a mesma decisão de buildPendingEntryRow_ (que continua rodando essas duas coisas
 * na ESCRITA pra email/importação de documento, fontes que não sofrem do problema de
 * deployment travado) — aqui só de novo, pra pendências do MacroDroid que nunca passaram
 * por lá. NÃO inclui o dedup contra Transaction confirmada (transactionExistsForPending_):
 * rodar isso a cada listagem, indefinidamente, arriscaria descartar em silêncio uma
 * pendência legítima que só por coincidência bate Data+Valor com uma transação
 * confirmada BEM DEPOIS dela já estar na fila — pior que ocasionalmente sobrar uma
 * pendência duplicada (visível, fácil de descartar na mão).
 *
 * Reescreve a aba inteira de uma vez (1 leitura + no máximo 1 clear + 1 write) em vez
 * de apagar linha por linha — com uma fila de centenas de pendências, deleteRowById_
 * em loop levava minutos (cada exclusão re-lê a coluna de ID e desloca as linhas
 * seguintes) e travava o app em "Carregando…" o tempo todo.
 */
function purgeJunkPendingEntries_() {
  withSheetLock_(function () {
    var sheet = getOrCreateSheet_('PendingEntries');
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var headers = SCHEMA.PendingEntries;
    var rows = readAll_('PendingEntries');

    var withDerived = rows.map(function (row) {
      return { row: row, derived: deriveNotificationFields_(row) };
    }).filter(function (x) {
      return x.derived.amount && !shouldIgnorePendingText_(x.derived.combinedText);
    });

    var cancelledIds = {};
    withDerived.forEach(function (x) {
      if (cancelledIds[x.row.ID]) return;
      var lower = x.derived.combinedText.toLowerCase();
      var isCorrida = lower.indexOf(CORRIDA_99_TEXT) !== -1;
      var isReembolso = lower.indexOf(REEMBOLSO_99_TEXT) !== -1;
      if (!isCorrida && !isReembolso) return;
      var counterpartPhrase = isCorrida ? REEMBOLSO_99_TEXT : CORRIDA_99_TEXT;
      var pair = withDerived.find(function (y) {
        if (y.row.ID === x.row.ID || cancelledIds[y.row.ID]) return false;
        if (y.derived.combinedText.toLowerCase().indexOf(counterpartPhrase) === -1) return false;
        return Math.abs(y.derived.amount - x.derived.amount) < 0.005;
      });
      if (pair) { cancelledIds[x.row.ID] = true; cancelledIds[pair.row.ID] = true; }
    });

    var kept = withDerived.filter(function (x) { return !cancelledIds[x.row.ID]; }).map(function (x) { return x.row; });
    if (kept.length === rows.length) return; // nada pra remover, evita reescrever à toa

    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
    if (kept.length) {
      sheet.getRange(2, 1, kept.length, headers.length).setValues(
        kept.map(function (row) { return objectToRow_(headers, row); })
      );
    }
  });
}

function listPendingEntries() {
  purgeJunkPendingEntries_();

  var accountsById = {};
  readAll_('Accounts').forEach(function (a) { accountsById[a.ID] = a; });
  var allCategories = readAll_('Categories');
  var categoriesById = {};
  allCategories.forEach(function (c) { categoriesById[c.ID] = c; });
  var allMappings = readAll_('CounterpartyMappings');
  var allAliases = readAll_('CounterpartyAliases');
  var allBankAppMappings = readAll_('BankAppMappings');

  var entries = readAll_('PendingEntries').map(function (p) {
    // Valor/tipo/contraparte: usa o que a fonte já gravou estruturado (email, importação
    // de documento); deriva do texto cru quando falta — sempre o caso do MacroDroid (ver
    // deriveNotificationFields_/captureRawNotification_) — recalculado a cada listagem,
    // então uma melhoria nessa lógica já vale pra pendências antigas também, sem precisar
    // recriar nada nem depender de qual versão o webhook do celular está rodando.
    var derived = deriveNotificationFields_(p);
    // Mesma ideia pra conta lembrada (BankAppMappings): se não veio gravada, procura na
    // hora — sem isso, pendência crua de um app já mapeado nunca sugeriria a conta certa.
    var suggestedAccountId = p.SuggestedAccountID || (function () {
      var mapping = allBankAppMappings.find(function (m) { return m.AppIdentifier === p.App; });
      return mapping ? mapping.AccountID : '';
    })();
    var acc = accountsById[suggestedAccountId];
    var toAcc = accountsById[p.SuggestedToAccountID];
    // Contraparte exata (CounterpartyMappings) e palavra-chave (Categories.Keywords) são
    // recalculadas NA HORA, não confiam só no SuggestedCategoryID gravado na criação —
    // assim marcar "lembrar categoria" ou editar/adicionar uma palavra-chave já vale pra
    // pendências que já estão na fila, sem precisar recriar nada. Contraparte exata vem
    // primeiro (sinal mais forte — já foi confirmado antes pra esse mesmo nome); palpite
    // gravado na criação (ex: BankAppMappings) e palavra-chave são fallback nessa ordem.
    var suggestedCategoryId = findCategoryIdForCounterparty_(derived.counterparty, allMappings) ||
      p.SuggestedCategoryID ||
      findCategoryIdByKeyword_(derived.counterparty || p.Text, allCategories);
    var cat = categoriesById[suggestedCategoryId];
    // Mesma ideia do suggestedCategoryId acima: recalculado NA HORA, não gravado na
    // criação — criar/editar uma associação já vale pras pendências que já estão na fila.
    var establishmentName = findEstablishmentForCounterparty_(derived.counterparty, allAliases);
    var capturedDateObj = new Date(p.CreatedAt);
    var tz = Session.getScriptTimeZone();
    return {
      id: p.ID,
      app: p.App,
      title: p.Title,
      text: p.Text,
      capturedAt: p.CreatedAt,
      date: Utilities.formatDate(capturedDateObj, tz, 'yyyy-MM-dd'),
      time: Utilities.formatDate(capturedDateObj, tz, 'HH:mm'),
      suggestedAmount: derived.amount,
      suggestedKind: derived.kind || null,
      suggestedAccountId: suggestedAccountId || null,
      suggestedAccountName: acc ? acc.Name : null,
      counterparty: derived.counterparty || null,
      establishmentName: establishmentName || null,
      suggestedCategoryId: suggestedCategoryId || null,
      suggestedCategoryName: cat ? cat.Name : null,
      suggestedToAccountId: p.SuggestedToAccountID || null,
      suggestedToAccountName: toAcc ? toAcc.Name : null,
      suggestedInstallments: p.SuggestedInstallments === '' ? null : Number(p.SuggestedInstallments),
      isDuplicate: false
    };
  }).sort(function (a, b) { return a.capturedAt < b.capturedAt ? 1 : -1; });

  return markDuplicatePendingEntries_(entries);
}

/**
 * Confirma a pendência: cria a transação de verdade (data/horário default pra quando
 * a notificação chegou, mas o usuário pode ter ajustado na revisão), opcionalmente
 * memoriza a conta pra esse app de banco e/ou a categoria pra essa contraparte, e
 * remove a pendência.
 */
function confirmPendingEntry(input) {
  if (!input.id) throw new Error('ID da pendência é obrigatório.');
  var pending = findById_('PendingEntries', input.id);
  if (!pending) throw new Error('Pendência não encontrada (talvez já processada).');

  var tz = Session.getScriptTimeZone();
  var capturedDateObj = new Date(pending.CreatedAt);
  var date = input.date || Utilities.formatDate(capturedDateObj, tz, 'yyyy-MM-dd');
  var time = input.time || Utilities.formatDate(capturedDateObj, tz, 'HH:mm');
  var counterparty = input.counterparty !== undefined ? input.counterparty : pending.Counterparty;

  var createdTransactions;
  if (input.kind === 'transfer') {
    var legs = createTransfer({
      date: date, time: time, fromAccountId: input.fromAccountId, toAccountId: input.toAccountId,
      amount: input.amount, description: input.description, billMonth: input.billMonth
    });
    createdTransactions = [legs.out, legs.in];
  } else {
    createdTransactions = createTransaction({
      kind: input.kind, date: date, time: time, accountId: input.accountId, categoryId: input.categoryId,
      amount: input.amount, description: input.description, counterparty: counterparty,
      installments: input.installments || Number(pending.SuggestedInstallments) || 1, forecastId: input.forecastId,
      forecastCycleMonth: input.forecastCycleMonth,
      excludeFromAnalysis: input.excludeFromAnalysis
    });
  }

  // Pra transferência, memoriza a conta de ORIGEM: é a que corresponde à notificação
  // desse app (o destino é só pra onde o dinheiro foi, não "a conta desse app").
  var accountToRemember = input.kind === 'transfer' ? input.fromAccountId : input.accountId;
  if (input.rememberAccountForApp && pending.App && accountToRemember) {
    var existingMapping = readAll_('BankAppMappings').find(function (m) { return m.AppIdentifier === pending.App; });
    if (existingMapping) {
      updateRowById_('BankAppMappings', existingMapping.ID, { AccountID: accountToRemember });
    } else {
      appendRow_('BankAppMappings', { AppIdentifier: pending.App, AccountID: accountToRemember });
    }
  }

  if (input.rememberCategoryForCounterparty && counterparty && input.kind !== 'transfer' && input.categoryId) {
    saveCounterpartyMapping_(counterparty, input.categoryId);
  }
  if (input.rememberCounterpartyAlias) {
    saveCounterpartyAlias_(input.rememberCounterpartyAlias.rawText, input.rememberCounterpartyAlias.establishmentName);
  }
  recordNinetyNineConfirmedSample_(pending);

  deleteRowById_('PendingEntries', input.id);
  return createdTransactions;
}

function discardPendingEntry(id) {
  var removed = deleteRowById_('PendingEntries', id);
  if (!removed) throw new Error('Pendência não encontrada.');
  return true;
}

/**
 * Manutenção manual (menu Finanças): força uma rodada de purgeJunkPendingEntries_ fora do
 * fluxo normal de listagem — frase ignorada, sem valor em R$ identificável (nem gravado
 * nem derivável do texto), ou par de corrida 99 cancelada/reembolso já cobrem tudo que
 * essa limpeza faz (a mesma lógica que roda sozinha a cada listagem). Não toca em nenhuma
 * pendência legítima.
 */
function removeJunkPendingEntries() {
  var before = readAll_('PendingEntries').length;
  purgeJunkPendingEntries_();
  var removedCount = before - readAll_('PendingEntries').length;
  SpreadsheetApp.getUi().alert(removedCount + ' pendência(s) removida(s) (frase ignorada, sem valor real, ou par de corrida cancelada/reembolso).');
}

// ===== Mode.js =====
/**
 * Modo de execução do backend.
 *
 * false = o app do dono (script preso à planilha dele, executa como o dono, MacroDroid,
 * Gmail, Drive, OCR). É o que está em src/ e o que roda na produção.
 * true  = beta multiusuário (script solto, executa como quem acessa, cada pessoa com a
 * própria planilha). Só o pacote gerado por scripts/build-multiuser.js troca esta linha
 * — não editar aqui.
 */
var MULTIUSER_MODE = true;

// ===== Plans.js =====
/**
 * CRUD de "Planos" (ver Planos em Planejamento) — eventos financeiros únicos e futuros
 * (ex: "Trocar impressora", R$1200, despesa) que alimentam só o gráfico de tendência de
 * saldo (ver computeBalanceTrend_, JavaScript.html): deslocam o saldo projetado a partir
 * do mês em que acontecem, sem afetar a tendência de crescimento/queda calculada. Não
 * têm conta nem categoria porque nunca geram uma Transaction de verdade — são puramente
 * um input do gráfico. Por isso a remoção aqui é definitiva (sem Archived), diferente de
 * Budgets/MonthlyForecasts: nada referencia um Plano por ID, então não há o que preservar.
 */

/**
 * Ignora (em vez de deixar quebrar a leitura inteira) qualquer linha com Month fora do
 * formato esperado — sem isso, uma única linha corrompida faria .sort() lançar exceção
 * e getBootstrapData() engolir o erro (ver try/catch em Code.js), escondendo TODOS os
 * planos, não só o de linha ruim.
 */
function listPlans() {
  return readAll_('Plans')
    .filter(function (p) { return /^\d{4}-\d{2}$/.test(p.Month || ''); })
    .map(function (p) {
      return { id: p.ID, label: p.Label, amount: Number(p.Amount), kind: p.Kind, month: p.Month };
    })
    .sort(function (a, b) { return a.month.localeCompare(b.month); });
}

function validatePlanInput_(input, requireAll) {
  var updates = {};
  if (requireAll || input.label !== undefined) {
    if (!input.label) throw new Error('Nome é obrigatório.');
    updates.Label = input.label;
  }
  if (requireAll || input.amount !== undefined) {
    var amount = Number(input.amount);
    if (!amount || amount <= 0) throw new Error('Valor precisa ser maior que zero.');
    updates.Amount = amount;
  }
  if (requireAll || input.kind !== undefined) {
    if (input.kind !== 'income' && input.kind !== 'expense') throw new Error('Tipo inválido.');
    updates.Kind = input.kind;
  }
  if (requireAll || input.month !== undefined) {
    var monthMatch = /^\d{4}-(\d{2})$/.exec(input.month || '');
    if (!monthMatch || Number(monthMatch[1]) < 1 || Number(monthMatch[1]) > 12) throw new Error('Mês inválido.');
    if (input.month < currentYm_()) throw new Error('Não é possível usar um mês passado — meses passados já têm saldo real.');
    updates.Month = input.month;
  }
  return updates;
}

function createPlan(input) {
  return appendRow_('Plans', validatePlanInput_(input, true));
}

function updatePlan(input) {
  if (!input.id) throw new Error('ID do plano é obrigatório.');
  var updates = validatePlanInput_(input, false);
  var updated = updateRowById_('Plans', input.id, updates);
  if (!updated) throw new Error('Plano não encontrado.');
  return updated;
}

function removePlan(id) {
  var removed = deleteRowById_('Plans', id);
  if (!removed) throw new Error('Plano não encontrado.');
  return true;
}

/**
 * "Meta" de saldo do gráfico de projeção (ver drawBalanceGoalLine_/JavaScript.html): uma
 * linha horizontal no valor digitado. No máximo UMA (aba BalanceGoal com 0 ou 1 linha) e
 * só o valor, sem data. null quando não há meta. Se já foi "atingida" ou não é derivado
 * no cliente a partir do saldo de hoje -- nada disso é gravado aqui.
 */
function getBalanceGoal() {
  var rows = readAll_('BalanceGoal');
  if (!rows.length) return null;
  var amount = Number(rows[0].Amount);
  return amount > 0 ? amount : null;
}

/** Cria a meta ou troca o valor da que já existe (upsert -- nunca acumula linha). */
function setBalanceGoal(input) {
  var amount = Number(input && input.amount);
  if (!amount || amount <= 0) throw new Error('Meta precisa ser maior que zero.');
  var rows = readAll_('BalanceGoal');
  if (rows.length) {
    updateRowById_('BalanceGoal', rows[0].ID, { Amount: amount, UpdatedAt: nowIso_() });
  } else {
    appendRow_('BalanceGoal', { Amount: amount, UpdatedAt: nowIso_() });
  }
  return amount;
}

/** Apaga a meta. Sem efeito (não é erro) se já não havia nenhuma. */
function removeBalanceGoal() {
  readAll_('BalanceGoal').forEach(function (r) { deleteRowById_('BalanceGoal', r.ID); });
  return true;
}

// ===== Recurring.js =====
/**
 * Transações recorrentes: receita, despesa ou transferência que se repete todo mês
 * num dia fixo, até ser cancelada. O motor (processRecurringRules) roda via gatilho
 * diário (ver ensureDailyTrigger em Setup.js) e também pode ser disparado manualmente
 * pela aba "Recorrências" no app.
 */

function effectiveDayForMonth_(day, year, monthIndex0) {
  var lastDay = new Date(year, monthIndex0 + 1, 0).getDate();
  return Math.min(Number(day), lastDay);
}

function nextRecurringDate_(dayOfMonth, lastCreatedYm) {
  var now = new Date();
  var year = now.getFullYear(), month = now.getMonth();
  if (lastCreatedYm === currentYm_()) {
    month += 1;
    if (month > 11) { month = 0; year += 1; }
  }
  var day = effectiveDayForMonth_(dayOfMonth, year, month);
  return Utilities.formatDate(new Date(year, month, day), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function listRecurringRules(includeArchived) {
  var rules = readAll_('RecurringRules').filter(function (r) { return includeArchived || !isArchived_(r.Archived); });
  var accountsById = {};
  readAll_('Accounts').forEach(function (a) { accountsById[a.ID] = a; });
  var categoriesById = {};
  readAll_('Categories').forEach(function (c) { categoriesById[c.ID] = c; });

  return rules.map(function (r) {
    var acc = accountsById[r.AccountID];
    var cat = categoriesById[r.CategoryID];
    var fromAcc = accountsById[r.FromAccountID];
    var toAcc = accountsById[r.ToAccountID];
    return {
      id: r.ID,
      kind: r.Kind,
      accountId: r.AccountID || null,
      accountName: acc ? acc.Name : null,
      categoryId: r.CategoryID || null,
      categoryName: cat ? cat.Name : null,
      fromAccountId: r.FromAccountID || null,
      fromAccountName: fromAcc ? fromAcc.Name : null,
      toAccountId: r.ToAccountID || null,
      toAccountName: toAcc ? toAcc.Name : null,
      amount: Number(r.Amount),
      description: r.Description || '',
      dayOfMonth: Number(r.DayOfMonth),
      archived: isArchived_(r.Archived),
      nextDate: nextRecurringDate_(r.DayOfMonth, r.LastCreatedYm)
    };
  });
}

function createRecurringRule(input) {
  if (input.kind !== 'income' && input.kind !== 'expense' && input.kind !== 'transfer') {
    throw new Error('Tipo de recorrência inválido: ' + input.kind);
  }
  var amount = Number(input.amount);
  if (!amount || amount <= 0) throw new Error('Valor precisa ser maior que zero.');
  var dayOfMonth = Number(input.dayOfMonth);
  if (!dayOfMonth || dayOfMonth < 1 || dayOfMonth > 31) throw new Error('Dia do mês precisa ser entre 1 e 31.');

  var row = {
    Kind: input.kind, Amount: amount, Description: input.description || '', DayOfMonth: dayOfMonth,
    Archived: false, AccountID: '', CategoryID: '', FromAccountID: '', ToAccountID: ''
  };

  if (input.kind === 'transfer') {
    if (!input.fromAccountId || !input.toAccountId) throw new Error('Contas de origem e destino são obrigatórias.');
    if (input.fromAccountId === input.toAccountId) throw new Error('Conta de origem e destino não podem ser a mesma.');
    if (!findById_('Accounts', input.fromAccountId) || !findById_('Accounts', input.toAccountId)) {
      throw new Error('Conta não encontrada.');
    }
    row.FromAccountID = input.fromAccountId;
    row.ToAccountID = input.toAccountId;
  } else {
    var account = findById_('Accounts', input.accountId);
    if (!account) throw new Error('Conta não encontrada.');
    if (account.Type === 'investment') throw new Error('Conta de investimento só aceita transferências.');
    assertKindMatchesCategory_(input.kind, input.categoryId);
    row.AccountID = input.accountId;
    row.CategoryID = input.categoryId;
  }

  // Se o dia já passou neste mês, só começa a valer no mês seguinte (sem lançar retroativo).
  var today = new Date();
  var effDay = effectiveDayForMonth_(dayOfMonth, today.getFullYear(), today.getMonth());
  row.LastCreatedYm = today.getDate() >= effDay ? currentYm_() : '';

  return appendRow_('RecurringRules', row);
}

function updateRecurringRule(input) {
  if (!input.id) throw new Error('ID da recorrência é obrigatório.');
  var updates = {};
  if (input.amount !== undefined) {
    var amount = Number(input.amount);
    if (!amount || amount <= 0) throw new Error('Valor precisa ser maior que zero.');
    updates.Amount = amount;
  }
  if (input.description !== undefined) updates.Description = input.description;
  if (input.dayOfMonth !== undefined) {
    var dayOfMonth = Number(input.dayOfMonth);
    if (!dayOfMonth || dayOfMonth < 1 || dayOfMonth > 31) throw new Error('Dia do mês precisa ser entre 1 e 31.');
    updates.DayOfMonth = dayOfMonth;
  }
  var updated = updateRowById_('RecurringRules', input.id, updates);
  if (!updated) throw new Error('Recorrência não encontrada.');
  return updated;
}

function archiveRecurringRule(id) {
  var updated = updateRowById_('RecurringRules', id, { Archived: true });
  if (!updated) throw new Error('Recorrência não encontrada.');
  return updated;
}

/**
 * Cria os lançamentos de todas as recorrências que estão em dia (ainda não
 * lançadas neste mês, e cujo dia configurado já chegou). Roda no gatilho diário
 * e também pode ser chamada manualmente pela aba Recorrências. Um erro numa
 * regra (ex: conta arquivada) não impede as demais de processarem.
 */
function processRecurringRules() {
  var today = new Date();
  var currentYm = currentYm_();
  var todayDay = today.getDate();
  var tz = Session.getScriptTimeZone();

  var rules = readAll_('RecurringRules').filter(function (r) { return !isArchived_(r.Archived); });
  var createdCount = 0;

  rules.forEach(function (r) {
    try {
      if (r.LastCreatedYm === currentYm) return;
      var effDay = effectiveDayForMonth_(r.DayOfMonth, today.getFullYear(), today.getMonth());
      if (todayDay < effDay) return;

      var dateStr = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth(), effDay), tz, 'yyyy-MM-dd');
      if (r.Kind === 'transfer') {
        createTransfer({
          date: dateStr, fromAccountId: r.FromAccountID, toAccountId: r.ToAccountID,
          amount: Number(r.Amount), description: r.Description
        });
      } else {
        createTransaction({
          kind: r.Kind, date: dateStr, accountId: r.AccountID, categoryId: r.CategoryID,
          amount: Number(r.Amount), description: r.Description, installments: 1
        });
      }
      updateRowById_('RecurringRules', r.ID, { LastCreatedYm: currentYm });
      createdCount++;
    } catch (err) {
      // Regra com problema (ex: conta/categoria arquivada) não deve travar as demais.
    }
  });

  return createdCount;
}

// ===== Refunds.js =====
/**
 * Reembolso de despesa compartilhada: associar um lançamento de RECEITA (o Pix que
 * chegou) à parte que uma PESSOA específica deve de uma despesa (o lançamento
 * "passivo"). N:N nos dois sentidos — uma despesa pode ter várias pessoas te devendo
 * (RefundShares), e um único reembolso pode cobrir a parte de várias pessoas ou várias
 * despesas de uma vez (RefundLinks) — por isso são DUAS tabelas de junção, não uma
 * coluna a mais em Transactions (que continua sem nenhuma mudança de schema).
 *
 * "Completo" nunca é uma flag: é sempre DERIVADO (soma dos RefundLinks de uma Share >=
 * ExpectedAmount dela). O valor líquido mostrado no Extrato/Dashboard também é derivado
 * (Amount original − o que já foi de fato vinculado, não a meta inteira) — a coluna
 * Transactions.Amount nunca é reescrita. Esse cálculo roda no CLIENTE (JavaScript.html)
 * a partir do que getBootstrapData expõe cru aqui: listPeople/listRefundShares/
 * listRefundLinks.
 *
 * Escopo v1: só despesa NÃO parcelada pode virar "passivo" (ver assertPassiveEligible_);
 * só receita pode ser "ativo" (ver assertActiveEligible_). O par de cancelamento
 * automático 99/Amazon (checkAndRemoveRefundPair_, Inbox.js) é um sistema totalmente
 * separado — aquele é "essa despesa nunca deveria ter existido" (cancela e some); este
 * é "essa despesa existiu de verdade, só que vou ser reembolsado por ela" (fica
 * visível, reduzida).
 *
 * UI (ver JavaScript.html): os dois popups (nomear pessoas / vincular reembolso) são
 * "staged até confirmar" — nada toca o servidor enquanto o popup está aberto; só o
 * clique em Concluído/Vincular manda a lista FINAL de uma vez (setRefundSharesFor
 * PassiveTransaction / setRefundLinksForActiveTransaction), que servem tanto pra criar
 * do zero (despesa/receita nova) quanto pra editar (nada existente = só cria; o que já
 * existia e não está mais na lista final é removido). Chegou a existir uma versão "ao
 * vivo" (cada edição salvando na hora, sozinha) — foi abandonada de propósito depois
 * de gerar a maioria dos bugs de corrida das primeiras rodadas; não recriar.
 */

function personNamesMatch_(a, b) {
  return (a || '').trim().toLowerCase() === (b || '').trim().toLowerCase();
}

/**
 * Carência antes de uma pessoa sem nenhuma Share poder ser apagada como órfã (ver
 * findOrphanPeopleIds_). A pessoa é criada (createPerson) um instante ANTES de
 * setRefundSharesForPassiveTransaction dar a primeira Share a ela -- uma leitura no meio
 * desses dois passos (ex: o refresh de outra mutação em voo) a veria sem Share nenhuma,
 * e apagá-la ali quebraria o "Concluído" com "Pessoa não encontrada". 1 hora cobre isso
 * com folga (a janela real é de segundos) sem deixar lixo de teste pra sempre.
 */
var ORPHAN_PERSON_GRACE_MS = 60 * 60 * 1000;

/**
 * IDs de pessoas (linhas cruas de People) que não têm NENHUMA Share (RefundShares.PersonID)
 * e já passaram da carência desde CreatedAt (ver ORPHAN_PERSON_GRACE_MS). Pura -- recebe as
 * linhas e "agora" -- pra dar pra testar sem planilha. CreatedAt vazio/ilegível conta como
 * antigo (linha de antes da coluna existir, sem como saber a idade).
 */
function findOrphanPeopleIds_(peopleRows, shareRows, nowMs) {
  var hasShare = {};
  shareRows.forEach(function (s) { hasShare[s.PersonID] = true; });
  return peopleRows.filter(function (p) {
    if (hasShare[p.ID]) return false;
    var createdMs = new Date(p.CreatedAt).getTime();
    return isNaN(createdMs) || nowMs - createdMs >= ORPHAN_PERSON_GRACE_MS;
  }).map(function (p) { return p.ID; });
}

/**
 * Pessoas de reembolso. Autocorretiva (mesmo espírito de listBudgetSuggestionDismissals,
 * Budgets.js): apaga daqui as pessoas órfãs (ver findOrphanPeopleIds_) -- People só é
 * criada de dentro do popup "Reembolso?" e NUNCA havia um caminho pra remover uma pessoa
 * quando a última Share dela sai (o popup só remove Shares, ver
 * setRefundSharesForPassiveTransaction), então nomes de teste/desistência ficavam na lista
 * pra sempre (pedido explícito do usuário: "nem deveriam existir mais"). Melhor esforço:
 * se apagar falhar, ainda devolve a lista (o erro vai pro log, não derruba o bootstrap).
 */
function listPeople(includeArchived) {
  var rows = readAll_('People');
  try {
    var orphanIds = findOrphanPeopleIds_(rows, readAll_('RefundShares'), Date.now());
    if (orphanIds.length) {
      orphanIds.forEach(function (id) { deleteRowById_('People', id); });
      rows = rows.filter(function (p) { return orphanIds.indexOf(p.ID) === -1; });
    }
  } catch (err) {
    console.error('Não foi possível limpar pessoas de reembolso órfãs: ' + err.message);
  }
  return rows
    .filter(function (p) { return includeArchived || !isArchived_(p.Archived); })
    .map(function (p) {
      return { id: p.ID, name: p.Name, archived: isArchived_(p.Archived) };
    });
}

function createPerson(input) {
  if (!input.name) throw new Error('Nome da pessoa é obrigatório.');
  var duplicate = readAll_('People').some(function (p) { return !isArchived_(p.Archived) && personNamesMatch_(p.Name, input.name); });
  if (duplicate) throw new Error('Já existe uma pessoa com esse nome.');
  return appendRow_('People', { Name: input.name, Archived: false });
}

function updatePerson(input) {
  if (!input.id) throw new Error('ID da pessoa é obrigatório.');
  if (!input.name) throw new Error('Nome da pessoa é obrigatório.');
  var duplicate = readAll_('People').some(function (p) {
    return p.ID !== input.id && !isArchived_(p.Archived) && personNamesMatch_(p.Name, input.name);
  });
  if (duplicate) throw new Error('Já existe uma pessoa com esse nome.');
  var updated = updateRowById_('People', input.id, { Name: input.name });
  if (!updated) throw new Error('Pessoa não encontrada.');
  return updated;
}

function archivePerson(id) {
  var updated = updateRowById_('People', id, { Archived: true });
  if (!updated) throw new Error('Pessoa não encontrada.');
  return updated;
}

function listRefundShares() {
  return readAll_('RefundShares').map(function (s) {
    return { id: s.ID, passiveTransactionId: s.PassiveTransactionID, personId: s.PersonID, expectedAmount: Number(s.ExpectedAmount) };
  });
}

function listRefundLinks() {
  return readAll_('RefundLinks').map(function (l) {
    return { id: l.ID, shareId: l.ShareID, activeTransactionId: l.ActiveTransactionID, allocatedAmount: Number(l.AllocatedAmount) };
  });
}

/** Só despesa NÃO parcelada pode virar "passivo" — parcelamento já espalha a compra em
 *  N linhas (InstallmentGroupID), e ligar reembolso a "qual das parcelas" é uma
 *  ambiguidade que não apareceu em nenhum caso de uso real ainda; melhor recusar cedo
 *  do que adivinhar (fica pra decidir se aparecer a necessidade de verdade). */
function assertPassiveEligible_(tx) {
  if (tx.Kind !== 'expense') throw new Error('Só uma despesa pode ficar aguardando reembolso.');
  if (tx.InstallmentGroupID) throw new Error('Compra parcelada ainda não pode ficar aguardando reembolso.');
}

/** Só receita pode ser "ativo" — uma transferência é entre CONTAS SUAS, nunca dinheiro
 *  vindo de outra pessoa, então não faz sentido como reembolso. */
function assertActiveEligible_(tx) {
  if (tx.Kind !== 'income') throw new Error('Só uma receita pode ser vinculada como reembolso.');
}

/** Remove uma Share e os RefundLinks dela junto — "desistir de cobrar" essa parte não
 *  pode deixar link órfão apontando pra uma Share que não existe mais. A(s) transação
 *  ativa(s) envolvida(s) simplesmente volta(m) a não ter aquela alocação — o valor dela
 *  volta a contar como receita normal (tudo derivado, nada mais a ajustar). */
function removeRefundShare_(shareId) {
  readAll_('RefundLinks').filter(function (l) { return l.ShareID === shareId; })
    .forEach(function (l) { deleteRowById_('RefundLinks', l.ID); });
  deleteRowById_('RefundShares', shareId);
}

/**
 * Substitui as Shares de uma despesa pelo conjunto final escolhido no popup — serve
 * tanto pra CRIAR (despesa nova, nenhuma Share existente ainda — everything vira
 * criação) quanto pra EDITAR (despesa já tinha Shares, ver popup "staged"). shares:
 * [{shareId (opcional — presente só se já existia de verdade), personId,
 * expectedAmount}]. Faz um DIFF por shareId, não um apaga-tudo-recria: uma Share
 * mantida (mesmo shareId na lista final) só tem o valor esperado atualizado, PRESER-
 * VANDO os RefundLinks dela — um apaga-e-recria ingênuo destruiria reembolsos já
 * recebidos contra uma Share que o usuário só queria editar o valor. Toda Share
 * existente que NÃO está na lista final é removida (removeRefundShare_, cascateando
 * os links dela) — é o "desmarquei essa pessoa" do popup.
 */
function setRefundSharesForPassiveTransaction(input) {
  if (!input.passiveTransactionId) throw new Error('Lançamento passivo é obrigatório.');
  var tx = findById_('Transactions', input.passiveTransactionId);
  if (!tx) throw new Error('Lançamento não encontrado.');
  assertPassiveEligible_(tx);
  var shares = input.shares || [];
  shares.forEach(function (s) {
    if (!s.personId) throw new Error('Pessoa é obrigatória em cada parte.');
    if (!(Number(s.expectedAmount) > 0)) throw new Error('Valor esperado precisa ser maior que zero.');
  });
  var totalExpected = shares.reduce(function (sum, s) { return sum + Number(s.expectedAmount); }, 0);
  if (Math.round(totalExpected * 100) > Math.round(Number(tx.Amount) * 100)) {
    throw new Error('A soma dos valores esperados não pode passar do valor da despesa.');
  }

  // keepIds precisa ter o ID de TODA linha do resultado final, criada ou mantida —
  // não só as que tinham shareId — senão uma Share recém-CRIADA aqui mesmo (sem
  // shareId ainda) não entraria em keepIds, e a limpeza abaixo a apagaria de volta no
  // mesmíssimo lugar em que acabou de nascer.
  var keepIds = {};
  var result = shares.map(function (s) {
    var row;
    if (s.shareId) {
      row = updateRowById_('RefundShares', s.shareId, { ExpectedAmount: Number(s.expectedAmount) });
      if (!row) throw new Error('Parte não encontrada.');
    } else {
      if (!findById_('People', s.personId)) throw new Error('Pessoa não encontrada.');
      row = appendRow_('RefundShares', {
        PassiveTransactionID: input.passiveTransactionId,
        PersonID: s.personId,
        ExpectedAmount: Number(s.expectedAmount)
      });
    }
    keepIds[row.ID] = true;
    return row;
  });
  readAll_('RefundShares').filter(function (s) { return s.PassiveTransactionID === input.passiveTransactionId; })
    .forEach(function (s) { if (!keepIds[s.ID]) removeRefundShare_(s.ID); });
  return result;
}

function assertActiveTransactionEligible_(activeTransactionId) {
  if (!activeTransactionId) throw new Error('Lançamento de receita é obrigatório.');
  var tx = findById_('Transactions', activeTransactionId);
  if (!tx) throw new Error('Lançamento não encontrado.');
  assertActiveEligible_(tx);
  return tx;
}
function validateAllocations_(allocations) {
  allocations = allocations || [];
  if (!allocations.length) throw new Error('Informe ao menos uma parte a cobrir.');
  allocations.forEach(function (a) {
    if (!a.shareId) throw new Error('Parte é obrigatória em cada alocação.');
    if (!findById_('RefundShares', a.shareId)) throw new Error('Parte não encontrada.');
    if (!(Number(a.amount) > 0)) throw new Error('Valor alocado precisa ser maior que zero.');
  });
  return allocations;
}

/** Pura: a soma das alocações passa do valor da receita? Compara em centavos (evita
 *  0,1 + 0,2 > 0,3 por ponto flutuante). Sem isso o líquido da receita ficava negativo
 *  (ex: receita 19,00 vinculada a 25,03 -> "+ -R$ 6,03*" no Extrato). */
function allocationsExceedAmount_(allocations, amount) {
  var cents = (allocations || []).reduce(function (sum, a) { return sum + Math.round(Number(a.amount) * 100); }, 0);
  return cents > Math.round(Number(amount) * 100);
}

/**
 * Substitui TODOS os vínculos de uma receita pelo conjunto final escolhido no popup —
 * serve pra criar (receita nova, nenhum vínculo existente ainda) e pra editar (receita
 * já tinha vínculos). Ao contrário de setRefundSharesForPassiveTransaction, aqui um
 * apaga-e-recria é seguro: nada mais referencia o ID de um RefundLink, então não tem
 * histórico pra perder ao trocar o ID de um vínculo mantido.
 */
function setRefundLinksForActiveTransaction(input) {
  var activeTx = assertActiveTransactionEligible_(input.activeTransactionId);
  var allocations = validateAllocations_(input.allocations);
  if (allocationsExceedAmount_(allocations, activeTx.Amount)) {
    throw new Error('A soma dos vínculos não pode passar do valor da receita. Se a pessoa pagou menos que o combinado, ajuste o valor devido dela na despesa.');
  }
  readAll_('RefundLinks').filter(function (l) { return l.ActiveTransactionID === input.activeTransactionId; })
    .forEach(function (l) { deleteRowById_('RefundLinks', l.ID); });
  return allocations.map(function (a) {
    return appendRow_('RefundLinks', {
      ShareID: a.shareId,
      ActiveTransactionID: input.activeTransactionId,
      AllocatedAmount: Number(a.amount)
    });
  });
}

/**
 * Ganchos chamados por updateTransaction/deleteTransaction (Transactions.js) — mantêm
 * Shares/Links consistentes quando o lançamento que eles referenciam muda por baixo,
 * sem exigir que o usuário desvincule manualmente antes (ver Q7 do grilling: "ajusta
 * sozinho" foi a opção escolhida).
 */

/** Lançamento (passivo OU ativo) removido: cascateia pras Shares/Links dele. */
function cascadeRefundsOnTransactionDelete_(transactionId) {
  readAll_('RefundShares').filter(function (s) { return s.PassiveTransactionID === transactionId; })
    .forEach(function (s) { removeRefundShare_(s.ID); });
  readAll_('RefundLinks').filter(function (l) { return l.ActiveTransactionID === transactionId; })
    .forEach(function (l) { deleteRowById_('RefundLinks', l.ID); });
}

/**
 * Lançamento ATIVO editado: se o novo valor ficar MENOR que a soma já alocada nos links
 * dele, remove todos os links (o usuário re-associa depois com o valor certo).
 *
 * O lado PASSIVO não reage mais a isso — chegou a reescalar a meta de cada pessoa
 * proporcionalmente (mantendo o % de cada uma), mas isso rodava em TODO Salvar do
 * lançamento (o cliente sempre manda o campo amount, mesmo quando o valor não mudou de
 * verdade — ver o filtro em updateTransaction, Transactions.js, que agora só chama
 * aqui quando o valor realmente é diferente), e mesmo só nesse caso o usuário reportou
 * o débito de alguém mudando "sozinho" ao reabrir o lançamento, sem ter mexido em nada.
 * A meta de cada pessoa só muda quando o PRÓPRIO usuário edita ela no popup de
 * reembolso (ver setRefundSharesForPassiveTransaction) — nunca como efeito colateral
 * de outra edição.
 */
function cascadeRefundsOnTransactionAmountChange_(transactionId, newAmount) {
  var links = readAll_('RefundLinks').filter(function (l) { return l.ActiveTransactionID === transactionId; });
  if (links.length) {
    var totalAllocated = links.reduce(function (sum, l) { return sum + Number(l.AllocatedAmount); }, 0);
    if (Math.round(totalAllocated * 100) > Math.round(newAmount * 100)) {
      links.forEach(function (l) { deleteRowById_('RefundLinks', l.ID); });
    }
  }
}

// ===== Setup.js =====
/**
 * Inicialização da planilha: cria as abas com cabeçalho e semeia categorias padrão.
 * Rode setupSpreadsheet() uma vez (menu "Finanças > Inicializar planilha") antes de usar o app.
 */

var DEFAULT_CATEGORIES = [
  { name: 'Salário', kind: 'income', children: [] },
  { name: 'Outras receitas', kind: 'income', children: ['Reembolso', 'Presente', 'Rendimento'] },
  { name: 'Alimentação', kind: 'expense', children: ['Mercado', 'Restaurante', 'Delivery'] },
  { name: 'Moradia', kind: 'expense', children: ['Aluguel', 'Condomínio', 'Energia', 'Água', 'Internet'] },
  { name: 'Transporte', kind: 'expense', children: ['Combustível', 'Transporte por app', 'Transporte público'] },
  { name: 'Saúde', kind: 'expense', children: ['Farmácia', 'Consulta', 'Plano de saúde'] },
  { name: 'Lazer', kind: 'expense', children: ['Assinaturas', 'Passeio'] },
  { name: 'Educação', kind: 'expense', children: [] },
  { name: 'Compras', kind: 'expense', children: [] },
  { name: 'Outros', kind: 'expense', children: [] }
];

/** Cria as abas e semeia categorias/conta padrão (idempotente). Sem UI: serve ao menu e ao onboarding. */
function seedSpreadsheet_() {
  Object.keys(SCHEMA).forEach(function (sheetName) { getOrCreateSheet_(sheetName); });

  if (readAll_('Categories').length === 0) {
    DEFAULT_CATEGORIES.forEach(function (cat) {
      var parent = appendRow_('Categories', {
        Name: cat.name, ParentID: '', Kind: cat.kind, Archived: false
      });
      cat.children.forEach(function (childName) {
        appendRow_('Categories', {
          Name: childName, ParentID: parent.ID, Kind: cat.kind, Archived: false
        });
      });
    });
  }

  if (readAll_('Accounts').length === 0) {
    appendRow_('Accounts', {
      Name: 'Carteira', Type: 'cash', InitialBalance: 0,
      ClosingDay: '', DueDay: '', Archived: false
    });
  }
}

function setupSpreadsheet() {
  seedSpreadsheet_();
  SpreadsheetApp.getUi().alert('Planilha inicializada com sucesso.');
}

/**
 * Onboarding do beta multiusuário: no primeiro acesso cria a planilha ("Cacau") no Drive de quem
 * está acessando, guarda o ID nas propriedades do usuário e semeia abas/categorias.
 * Devolve o ID. Chamadas seguintes só devolvem o ID guardado.
 *
 * Usa o lock de SCRIPT (não o de usuário) de propósito: a semeadura chama appendRow_, que
 * pega o lock de usuário (withSheetLock_) — pegar o mesmo lock duas vezes na mesma
 * execução trava. O custo é só serializar o PRIMEIRO acesso de usuários diferentes.
 */
function ensureUserSpreadsheet_() {
  var props = PropertiesService.getUserProperties();
  var existing = props.getProperty(USER_SHEET_ID_KEY);
  if (existing) return existing;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    existing = props.getProperty(USER_SHEET_ID_KEY);
    if (existing) return existing;
    var ss = SpreadsheetApp.create(appName_());
    var starterSheets = ss.getSheets();
    spreadsheetCache_ = ss;
    seedSpreadsheet_();
    // A aba que o Google cria junto ("Página1"/"Sheet1") não faz parte do app.
    starterSheets.forEach(function (sheet) {
      if (!SCHEMA[sheet.getName()]) ss.deleteSheet(sheet);
    });
    // Só grava o ID depois de semear: se algo falhar no meio, o próximo acesso refaz
    // tudo (deixando no máximo uma planilha órfã) em vez de ficar com uma pela metade.
    props.setProperty(USER_SHEET_ID_KEY, ss.getId());
    return ss.getId();
  } finally {
    lock.releaseLock();
  }
}

var PREVIOUS_SHEETS_KEY = 'PREVIOUS_SPREADSHEET_IDS';

function requireMultiUser_() {
  if (!MULTIUSER_MODE) throw new Error('Disponível só no beta multiusuário.');
}

/** Guarda o ID da planilha atual (se houver) na lista de anteriores, pra nunca "perder" uma planilha. */
function rememberPreviousSheet_() {
  var props = PropertiesService.getUserProperties();
  var current = props.getProperty(USER_SHEET_ID_KEY);
  if (!current) return;
  var previous = [];
  try { previous = JSON.parse(props.getProperty(PREVIOUS_SHEETS_KEY) || '[]'); } catch (e) { previous = []; }
  if (previous.indexOf(current) === -1) previous.push(current);
  props.setProperty(PREVIOUS_SHEETS_KEY, JSON.stringify(previous));
}

/** Aceita o ID puro ou o link completo da planilha. */
function extractSpreadsheetId_(raw) {
  var text = String(raw || '').trim();
  var fromUrl = text.match(/\/d\/([A-Za-z0-9_-]{20,})/);
  if (fromUrl) return fromUrl[1];
  return /^[A-Za-z0-9_-]{20,}$/.test(text) ? text : '';
}

function recoveryResult_() {
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; }
  return { ok: true, url: url };
}

/**
 * Recuperação quando a planilha da pessoa sumiu ou perdeu o acesso: cria uma NOVA (com as
 * categorias padrão). A antiga não é apagada — o ID dela fica em PREVIOUS_SPREADSHEET_IDS,
 * e a pessoa pode voltar a ela com reconnectUserSpreadsheet. Chamada pela tela de recuperação.
 */
function recreateUserSpreadsheet() {
  requireMultiUser_();
  rememberPreviousSheet_();
  PropertiesService.getUserProperties().deleteProperty(USER_SHEET_ID_KEY);
  spreadsheetCache_ = null;
  ensureUserSpreadsheet_();
  return recoveryResult_();
}

/**
 * Reconecta o app a uma planilha que a pessoa já tem (ID ou link). Só aceita se a conta
 * consegue abri-la E ela parece do app (tem as abas Transactions e Categories) — assim um
 * ID colado errado não faz o app escrever em planilha alheia.
 */
function reconnectUserSpreadsheet(rawId) {
  requireMultiUser_();
  var id = extractSpreadsheetId_(rawId);
  if (!id) throw new Error('Cole o código ou o link da planilha.');
  var ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (err) {
    throw new Error('Não consegui abrir essa planilha com esta conta. Confira o código e se ela está compartilhada com você.');
  }
  if (!ss.getSheetByName('Transactions') || !ss.getSheetByName('Categories')) {
    throw new Error('Essa planilha não parece ser do app de finanças (faltam as abas Transactions e Categories).');
  }
  rememberPreviousSheet_();
  PropertiesService.getUserProperties().setProperty(USER_SHEET_ID_KEY, id);
  spreadsheetCache_ = ss;
  return recoveryResult_();
}

// ===== SheetService.js =====
/**
 * Camada genérica de acesso às abas da planilha.
 * Cada aba tem uma linha de cabeçalho e cada linha vira um objeto {header: valor}.
 */

var SCHEMA = {
  // Description/InvestmentDate/MaturityDate ficam no final de propósito: se o
  // schema crescer de novo, planilhas já existentes não têm suas colunas antigas
  // reinterpretadas (ver getOrCreateSheet_).
  // ColorIndex: índice (0-9) na paleta fixa de cores de conta (ver ACCOUNT_COLOR_PALETTE_
  // em Accounts.js/JavaScript.html), não mais um matiz livre — renomeado de ColorHue.
  Accounts: [
    'ID', 'Name', 'Type', 'InitialBalance', 'ClosingDay', 'DueDay', 'Archived', 'CreatedAt',
    'Description', 'InvestmentDate', 'MaturityDate', 'ColorIndex'
  ],
  // Keywords (ver Categories.js): lista separada por vírgula, comparada como substring
  // (case-insensitive) contra o texto de um lançamento/pendência pra sugerir essa
  // categoria automaticamente — ex: "airbnb, booking.com" na subcategoria Hospedagem.
  // Description (no final, ver comentário do Accounts): texto livre opcional respondendo
  // "quais tipos de lançamento devem ser feitos nessa categoria?" — mostrado como tooltip
  // em itálico ao passar o mouse na categoria (ver JavaScript.html).
  Categories: ['ID', 'Name', 'ParentID', 'Kind', 'Archived', 'CreatedAt', 'Keywords', 'Description', 'Icon'],
  // Time (HH:mm), Counterparty ("Recebido de"/"Pago a") e ForecastID (vínculo com uma
  // Previsão mensal, ver Forecasts.js) são opcionais e ficam no final pelo mesmo
  // motivo do Accounts acima. Description agora é "Referente a".
  // ExcludeFromAnalysis (no final, ver comentário do Accounts acima): marca um
  // lançamento real mas "fora da curva" (ex: quitação pontual de dívida, ajuste de
  // conciliação de saldo) que deve continuar contando pro saldo real (Extrato/
  // Dashboard) mas sair de qualquer média/tendência (cor do calendário, orçamentos
  // Automático, projeção de saldo, sugestão de valor de recorrente) -- ver
  // shouldExcludeFromAnalysis_/JavaScript.html. Nunca ligada automaticamente sem o
  // usuário confirmar (só uma sugestão), grillado com o usuário.
  // ForecastCycleMonth (no final, mesmo motivo): 'YYYY-MM' do ciclo de um Lançamento
  // recorrente MENSAL ao qual esse lançamento está vinculado, quando é DIFERENTE do mês da
  // própria data dele (ex: pagou em setembro o aluguel de agosto). Vazio = o ciclo é o do
  // mês da data (o que sempre foi). Só faz sentido com ForecastID preenchido e recorrente
  // mensal (ciclo = mês de calendário); nos demais os ciclos são janelas móveis, sem
  // identidade estável pra ancorar (ver forecastCycleWindows_/JavaScript.html).
  Transactions: [
    'ID', 'Date', 'AccountID', 'CategoryID', 'Kind', 'Amount', 'Description',
    'BillMonth', 'InstallmentGroupID', 'InstallmentSeq', 'InstallmentTotal',
    'TransferGroupID', 'CreatedAt', 'Time', 'Counterparty', 'ForecastID', 'ExcludeFromAnalysis',
    'ForecastCycleMonth'
  ],
  // MonthlyLimit é, apesar do nome legado, o valor esperado a cada IntervalDays dias
  // (ex: R$100 a cada 7 dias) — não mais um valor fixo mensal. Mantive o nome da coluna
  // pra não precisar migrar dados existentes; linhas antigas sem IntervalDays são
  // tratadas como 30 dias na leitura (ver Budgets.js). AutoIncludeParentId (vazio na
  // maioria): quando um orçamento foi criado/editado selecionando uma categoria de topo
  // + TODAS as suas subcategorias atuais, guarda o ID dela — a partir daí CategoryID
  // deixa de ser a fonte de verdade pra esse orçamento; categoryIds é recalculado do
  // zero a cada leitura (categoria + subcategorias ATUAIS), pra uma subcategoria nova
  // criada depois entrar automaticamente sem precisar editar o orçamento (ver
  // classifyBudgetCategorySelection_/effectiveCategoryIdsForBudgetRow_, Budgets.js).
  // IsUnbudgetedCatchAll marca a ÚNICA linha especial (auto-criada por listBudgets,
  // Budgets.js) que rastreia despesas fora de qualquer orçamento -- CategoryID fica
  // vazio nela (categorias de verdade são recalculadas ao vivo, ver JavaScript.html),
  // só CreatedAt importa (âncora do período de 30 dias que ela mesma controla).
  // PeriodMode ('manual'|'auto', linha antiga sem essa coluna = 'manual'): 'auto' faz
  // o orçamento usar o período/teto derivados de Y (data de engajamento real) em vez
  // de IntervalDays/MonthlyLimit -- ver budgetAutoWindowDays_/computeBudgetAutoProgress_,
  // JavaScript.html, grillado com o usuário. IntervalDays/MonthlyLimit continuam
  // gravados mesmo com PeriodMode='auto' (só ficam sem uso enquanto isso durar) --
  // voltar pra outro período restaura o valor de antes sozinho, sem precisar de
  // colunas extras: o formulário simplesmente não manda um valor novo pra eles
  // enquanto o modo for 'auto' (ver openBudgetModal).
  Budgets: ['ID', 'CategoryID', 'MonthlyLimit', 'Archived', 'CreatedAt', 'IntervalDays', 'AutoIncludeParentId', 'IsUnbudgetedCatchAll', 'PeriodMode'],
  // Lançamento recorrente (ver Forecasts.js): um gasto/receita que se repete a cada
  // IntervalDays dias (ex: "Diarista" a cada 7 dias, "Salário" a cada 30) mas não faz
  // sentido como um limite de Orçamento. CategoryID pode ser categoria OU subcategoria
  // (sem a restrição de topo que Budgets tem). Repete automaticamente até ser editado
  // ou removido; o vínculo com o lançamento real de cada ocorrência fica na própria
  // Transaction (campo ForecastID), não aqui. IntervalDays no final de propósito (ver
  // Accounts acima) — linhas antigas sem essa coluna são tratadas como 30 dias na
  // leitura (mesma convenção de Budgets.IntervalDays).
  // IntervalUnit ('days'|'months', linha antiga sem essa coluna = 'days') diz como
  // IntervalDays deve ser lido: dias direto, ou quantidade de MESES de verdade (não
  // multiplicado por 30) — ver validateForecastInterval_/Forecasts.js.
  MonthlyForecasts: ['ID', 'Kind', 'CategoryID', 'Label', 'ExpectedAmount', 'Archived', 'CreatedAt', 'IntervalDays', 'IntervalUnit', 'DayOfMonth'],
  // AccountID/CategoryID (receita/despesa) ou FromAccountID/ToAccountID (transferência)
  // são preenchidos conforme o Kind; os outros ficam vazios.
  RecurringRules: [
    'ID', 'Kind', 'AccountID', 'CategoryID', 'FromAccountID', 'ToAccountID',
    'Amount', 'Description', 'DayOfMonth', 'Archived', 'LastCreatedYm', 'CreatedAt'
  ],
  // Notificações capturadas pelo MacroDroid, aguardando revisão antes de virar
  // Transaction de verdade (ver Inbox.js). CreatedAt aqui é a hora que a notificação chegou.
  // SourceMessageId (só pra pendências de email) é o ID único da mensagem no Gmail —
  // usado pra deduplicar num reprocessamento sem depender de Título+Texto, que pra
  // notificações genéricas (ex: toda "Corrida 99" tem o mesmo texto) não distingue
  // uma transação real de outra.
  // SuggestedToAccountID/SuggestedInstallments (ver DocumentImport.js): só usados por
  // pendências de transferência (conta de destino já conhecida, ex: fatura de cartão
  // sempre paga na mesma conta) ou de compra parcelada importada de um documento.
  PendingEntries: [
    'ID', 'App', 'Title', 'Text', 'SuggestedAmount', 'SuggestedKind', 'SuggestedAccountID', 'CreatedAt',
    'Counterparty', 'SuggestedCategoryID', 'SourceMessageId', 'SuggestedToAccountID', 'SuggestedInstallments'
  ],
  // Registro permanente de "esse email do 99Pay já foi avaliado" (ver EmailInbox.js).
  // Existe separado de PendingEntries porque essa aba tem linhas apagadas o tempo todo
  // (confirmar ou descartar uma pendência remove a linha) — usar só PendingEntries pra
  // saber "isso já foi visto" faria um email já confirmado ser reprocessado de novo e
  // virar pendência duplicada. O rótulo do Gmail não serve pra isso porque só existe
  // por thread inteira, não por mensagem individual.
  ProcessedEmailMessages: ['ID', 'MessageId', 'CreatedAt'],
  // Lembra qual conta usar por padrão pra notificações de um determinado app de banco.
  BankAppMappings: ['ID', 'AppIdentifier', 'AccountID', 'CreatedAt'],
  // Lembra qual categoria usar por padrão pra lançamentos de uma mesma pessoa/estabelecimento.
  CounterpartyMappings: ['ID', 'CounterpartyName', 'CategoryID', 'CreatedAt'],
  // Lembra o nome de verdade de um estabelecimento pro texto CRU que chega automático (ex:
  // "Comerciald" -> "Planeta Doce") — só afeta o que é SUGERIDO em pendências novas
  // (findEstablishmentForCounterparty_, Inbox.js), nunca reescreve lançamentos já confirmados.
  CounterpartyAliases: ['ID', 'RawText', 'EstablishmentName', 'CreatedAt'],
  // Textos de notificação do 99 que JÁ VIRARAM lançamento confirmado (ver
  // recordNinetyNineConfirmedSample_, Inbox.js) -- pedido explícito do usuário pra
  // substituir a lista IGNORED_PENDING_PHRASES (que só crescia pra tentar adivinhar o
  // que NÃO deveria entrar) por uma coleta do que DEVE entrar: olhando essa aba, o
  // usuário mesmo identifica o padrão de notificação legítima do 99, sem precisar de
  // mais nenhuma frase nova na lista de ignoradas. Uma linha por texto ÚNICO (dedupe
  // por Text, não por notificação individual) -- o objetivo é ver os FORMATOS
  // diferentes que já apareceram, não contar quantas vezes cada um se repetiu.
  Notification99Samples: ['ID', 'Text', 'App', 'Title', 'CreatedAt'],
  // "Descartar" uma sugestão de orçamento (novo OU aumento de período — ver
  // computeBudgetSuggestions_/computeBudgetPeriodIncreaseSuggestions_, JavaScript.html)
  // esconde ela por 30 dias, não pra sempre — depois disso volta a aparecer se ainda for
  // válida. Upsert em dismissBudgetSuggestion_/Budgets.js; linhas com mais de 30 dias são
  // purgadas sozinhas na leitura, igual purgeJunkPendingEntries_. SuggestionType/BudgetID
  // no final de propósito (ver comentário do Accounts acima) — linha antiga (de antes de
  // existir sugestão de aumento) não tem esses 2 campos; SuggestionType vazio é tratado
  // como 'new' na leitura (ver listBudgetSuggestionDismissals). Uma sugestão de NOVO
  // orçamento é identificada por CategoryID (ainda não existe orçamento, só a categoria);
  // uma de AUMENTO é identificada por BudgetID (o orçamento já existe, pode cobrir várias
  // categorias de uma vez — CategoryID sozinho não seria uma chave única nesse caso).
  BudgetSuggestionDismissals: ['ID', 'CategoryID', 'DismissedAt', 'SuggestionType', 'BudgetID'],
  // "Plano" (ver Plans.js): evento financeiro único e futuro (ex: "Trocar impressora",
  // R$1200, despesa) usado só pelo gráfico de tendência de saldo (ver
  // computeBalanceTrend_, JavaScript.html) — nunca gera Transaction, por isso não tem
  // conta nem categoria. Month é sempre o mês corrente ou futuro (validado em
  // createPlan/updatePlan); um mês que já passou simplesmente para de contar pro
  // gráfico (que usa saldo real pra meses passados), mas a linha só some daqui se
  // removida manualmente.
  Plans: ['ID', 'Label', 'Amount', 'Kind', 'Month'],
  // Reembolsos (ver Refunds.js) — quem te deve uma parte de uma despesa, e os
  // lançamentos de receita que já cobriram parte dela. Sem tab própria no app: gerido
  // de dentro do lançamento (botão "Reembolso?") e da aba Reembolsos.
  People: ['ID', 'Name', 'Archived', 'CreatedAt'],
  // Uma linha por (despesa compartilhada, pessoa que deve uma parte dela). ExpectedAmount
  // é a meta dessa pessoa — "completo" é derivado (soma dos RefundLinks dessa Share).
  RefundShares: ['ID', 'PassiveTransactionID', 'PersonID', 'ExpectedAmount', 'CreatedAt'],
  // Uma linha por (lançamento de receita, Share que ele cobre) — é o que permite 1
  // reembolso cobrir 2 despesas, ou 1 despesa ser coberta por vários reembolsos.
  RefundLinks: ['ID', 'ShareID', 'ActiveTransactionID', 'AllocatedAmount', 'CreatedAt'],
  // "Ignorar mês" de um ciclo não pago de um Lançamento recorrente (ver
  // computeForecastCycleStatus_/JavaScript.html, grillado com o usuário) -- marca uma
  // bolinha vermelha como definitivamente descartada (vira cinza), sem apagar o
  // histórico. CycleStart identifica o ciclo (a data de início da janela daquela
  // bolinha, YYYY-MM-DD) — junto com ForecastID formam a chave única (upsert em
  // ignoreForecastCycle/Forecasts.js). Reversível (unignoreForecastCycle apaga a linha).
  ForecastCycleIgnores: ['ID', 'ForecastID', 'CycleStart', 'IgnoredAt'],
  // Descarte (por 30 dias, mesmo mecanismo/janela de BudgetSuggestionDismissals) do
  // "Sugestão: R$X Aplicar" de um Lançamento recorrente (ver forecastSuggestedAmount/
  // forecastRowHtml, JavaScript.html) — pedido explícito do usuário, que não tinha como
  // descartar antes (só "Aplicar"). Upsert por ForecastID (dismissForecastSuggestion/
  // Forecasts.js) -- não guarda o valor sugerido no momento do descarte de propósito:
  // se a sugestão recalcular pra outro número enquanto ainda dentro dos 30 dias, continua
  // descartada (o usuário já disse "não agora" pra essa previsão, não pra um valor
  // específico).
  ForecastSuggestionDismissals: ['ID', 'ForecastID', 'DismissedAt'],
  // "Meta" de saldo do gráfico de projeção (ver setBalanceGoal/Plans.js e
  // drawBalanceGoalLine_/JavaScript.html, grillado com o usuário): no máximo UMA linha
  // (uma meta só, só o valor -- sem data). Vive na planilha (não em localStorage) pra
  // sincronizar entre celular e computador. Se "atingida" ou não é sempre DERIVADO do
  // saldo de hoje, nunca gravado aqui.
  BalanceGoal: ['ID', 'Amount', 'UpdatedAt']
};

// Colunas que guardam datas como texto (ex: '2026-08-15', '2026-08') para
// evitar que o Sheets as autoconverta em valores de data/hora.
//
// PendingEntries.CreatedAt entrou aqui depois de achar a causa real de "a pendência
// nunca aparece pra revisar depois de importar um documento": createPendingEntry
// (MacroDroid/email) sempre escreveu CreatedAt como new Date().toISOString() (ex:
// "2026-08-27T14:32:10.123Z"), formato que o Sheets nunca converteu sozinho. A
// importação de documento escreve um formato diferente (ex: "2026-04-22T12:00:00",
// sem milissegundos/Z — é a data da linha da fatura, não "agora") pra várias linhas de
// uma vez via appendRowsBatch_, e ESSE formato o Sheets reconhece como data/hora de
// verdade e autoconverte a célula. Isso faz p.CreatedAt voltar como um objeto Date (não
// string) na leitura, e markDuplicatePendingEntries_ chama .slice() nele — quebra com
// "e.capturedAt.slice is not a function", capturado (silenciosamente) pelo try/catch
// de listPendingEntries em getBootstrapData, que degrada pendingEntries pra [] sem
// avisar ninguém. Com a coluna aqui, normalizeTextColumns_ desfaz a autoconversão na
// leitura, não importa como a célula foi escrita.
var TEXT_COLUMNS = {
  Accounts: ['InvestmentDate', 'MaturityDate'],
  Transactions: ['Date', 'BillMonth', 'Time'],
  RecurringRules: ['LastCreatedYm'],
  PendingEntries: ['CreatedAt'],
  Plans: ['Month'],
  ForecastCycleIgnores: ['CycleStart']
};

// Padrão de formatação usado para normalizar de volta a texto caso o Sheets
// tenha convertido a célula em Date de qualquer forma (ex: colagem, edição manual).
var TEXT_COLUMN_DATE_PATTERNS = {
  Date: 'yyyy-MM-dd',
  InvestmentDate: 'yyyy-MM-dd',
  MaturityDate: 'yyyy-MM-dd',
  BillMonth: 'yyyy-MM',
  LastCreatedYm: 'yyyy-MM',
  Time: 'HH:mm',
  CreatedAt: "yyyy-MM-dd'T'HH:mm:ss",
  Month: 'yyyy-MM',
  CycleStart: 'yyyy-MM-dd'
};

/**
 * Planilha desta execução. App do dono: a planilha à qual o script está preso. Multiusuário
 * (MULTIUSER_MODE, ver Mode.js): a que o onboarding criou no Drive de quem está acessando,
 * guardada por ID nas propriedades do USUÁRIO (cada pessoa só enxerga as próprias).
 * Guardada em variável global — vale só durante a execução, evita reabrir a cada leitura.
 */
var USER_SHEET_ID_KEY = 'SPREADSHEET_ID';
var spreadsheetCache_ = null;

function getSpreadsheet_() {
  if (spreadsheetCache_) return spreadsheetCache_;
  if (MULTIUSER_MODE) {
    var id = PropertiesService.getUserProperties().getProperty(USER_SHEET_ID_KEY);
    if (!id) throw new Error('A planilha desta conta ainda não foi criada (abra o app de novo).');
    spreadsheetCache_ = SpreadsheetApp.openById(id);
  } else {
    spreadsheetCache_ = SpreadsheetApp.getActiveSpreadsheet();
  }
  return spreadsheetCache_;
}

/**
 * Devolve a aba, criando/completando o cabeçalho quando faz falta.
 *
 * CUIDADO: isso roda em TODA leitura (readAll_ chama aqui), então precisa ser um no-op
 * barato no caso normal. Já tem duas cicatrizes aqui:
 *
 *  1. `getRange(1, 1, 1, headers.length)` estourava ("out of bounds") quando a aba tinha
 *     MENOS colunas que o schema — situação que acontece exatamente quando o schema
 *     cresce (foi o caso do Description em Categories). E como isso acontecia ANTES de
 *     escrever o cabeçalho novo, a aba nunca se consertava: toda leitura daquela aba
 *     jogava exceção pra sempre. Em getBootstrapData, que engole erro por seção, o
 *     sintoma era silencioso e enganoso — a seção voltava vazia, e com Transactions
 *     vazio TODA conta passava a mostrar só o InitialBalance.
 *  2. Escrever aqui coloca o caminho de ESCRITA (que pega lock) dentro do caminho de
 *     LEITURA do bootstrap. Quando um schema cresce, a primeira leitura de cada execução
 *     concorrente tenta gravar o cabeçalho ao mesmo tempo, e as escritas de verdade do
 *     usuário (ex: o lançamento de consolidação de saldo) ficam esperando o lock e
 *     estouram "Lock timeout" — desfazendo a edição dele. Mesma classe de problema do
 *     rename ColorHue -> ColorIndex (ver defaultColorIndexForId_, Accounts.js).
 *
 * Por isso agora: garante a LARGURA da aba antes de qualquer getRange; lê só o que
 * existe; e só escreve quando o cabeçalho realmente está diferente — nunca de novo
 * depois de arrumado.
 */
function getOrCreateSheet_(name) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  var headers = SCHEMA[name];
  // Alarga a aba primeiro: sem isso o getRange abaixo estoura quando o schema cresceu.
  var maxCols = sheet.getMaxColumns();
  if (maxCols < headers.length) {
    sheet.insertColumnsAfter(maxCols, headers.length - maxCols);
  }
  var firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var hasHeaders = headers.every(function (h, i) { return firstRow[i] === h; });
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    var textCols = TEXT_COLUMNS[name] || [];
    textCols.forEach(function (colName) {
      var colIndex = headers.indexOf(colName) + 1;
      sheet.getRange(1, colIndex, sheet.getMaxRows(), 1).setNumberFormat('@');
    });
  }
  return sheet;
}

function generateId_() {
  return Utilities.getUuid();
}

function nowIso_() {
  return new Date().toISOString();
}

/**
 * Data/hora de uma célula como texto ISO, pra devolver ao cliente. O Sheets converte em
 * Date uma string ISO gravada por setValues (ex: DismissedAt), e google.script.run NÃO
 * serializa Date: se qualquer pedaço do retorno tiver um, a chamada inteira resolve com
 * null no cliente (foi o "Cannot read properties of null (reading 'errors')" de
 * getBootstrapData depois de descartar uma sugestão). Texto passa direto.
 */
function isoCell_(value) {
  if (value instanceof Date) return isNaN(value.getTime()) ? '' : value.toISOString();
  return value;
}

function rowToObject_(headers, row) {
  var obj = {};
  headers.forEach(function (h, i) { obj[h] = row[i]; });
  return obj;
}

function objectToRow_(headers, obj) {
  return headers.map(function (h) {
    return obj[h] !== undefined && obj[h] !== null ? obj[h] : '';
  });
}

/**
 * Corrige em memória colunas de texto que o Sheets tenha convertido para Date
 * (acontece às vezes com colagem/edição manual mesmo com a coluna formatada
 * como texto simples). Sem isso, código como `data.indexOf(...)` quebra
 * porque `data` deixa de ser string.
 */
/**
 * "Está arquivado?" tolerante ao FORMATO da célula. A checagem antiga era `Archived !==
 * true`, que só reconhece o booleano nativo do Sheets — uma célula com o TEXTO "TRUE"
 * (jeito que uma edição manual, uma colagem ou uma versão antiga do app podia deixar)
 * passava batido e o item arquivado voltava a aparecer nos seletores. Só o booleano
 * false, vazio ou "FALSE" contam como não-arquivado.
 */
function isArchived_(value) {
  if (value === true) return true;
  if (value === false || value === '' || value === undefined || value === null) return false;
  var text = String(value).trim().toLowerCase();
  return text === 'true' || text === 'verdadeiro' || text === '1' || text === 'sim';
}

/** Mesma tolerância de formato de isArchived_, só com um nome que não confunde quem
 *  ler uma célula booleana qualquer que não seja "arquivado" (ex: IsUnbudgetedCatchAll,
 *  Budgets.js). */
function parseBoolCell_(value) {
  return isArchived_(value);
}

function normalizeTextColumns_(sheetName, obj) {
  var textCols = TEXT_COLUMNS[sheetName] || [];
  textCols.forEach(function (col) {
    var value = obj[col];
    if (value instanceof Date) {
      var pattern = TEXT_COLUMN_DATE_PATTERNS[col] || 'yyyy-MM-dd';
      obj[col] = Utilities.formatDate(value, Session.getScriptTimeZone(), pattern);
    }
  });
  return obj;
}

/** Lê todas as linhas de uma aba como array de objetos. */
function readAll_(sheetName) {
  var sheet = getOrCreateSheet_(sheetName);
  var headers = SCHEMA[sheetName];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values
    .map(function (row) { return normalizeTextColumns_(sheetName, rowToObject_(headers, row)); })
    .filter(function (obj) { return obj.ID; });
}

/**
 * google.script.run pode disparar várias execuções concorrentes de verdade (ex: o
 * usuário clica em "arquivar"/"descartar" várias vezes rápido) — e updateRowById_/
 * deleteRowById_ leem o índice da linha ANTES de escrever. Sem serialização, duas
 * execuções concorrentes podem ler o mesmo índice e uma delete/edite a linha errada
 * depois que a outra já deslocou as linhas (isso causava "Pendência não encontrada"
 * ao descartar/confirmar várias pendências rápido). LockService serializa: só uma
 * execução mexe na planilha por vez, as outras esperam a vez em vez de corromper dado.
 */
/*
 * NÃO chamar SpreadsheetApp.flush() aqui dentro: já foi tentado (na hipótese de que a
 * escrita ficava em buffer e o refresh seguinte não a enxergava) e o efeito real foi
 * cada escrita segurar o lock por MUITO mais tempo — createTransaction parcelado faz uma
 * chamada dessas por parcela — o que passou a estourar "Lock timeout: another process was
 * holding the lock for too long" ao editar/criar vários itens em sequência, desfazendo a
 * edição do usuário. A causa de verdade do atraso no Extrato era outra (o
 * confirmPendingEntry não devolvia o lançamento criado pro cliente — ver Inbox.js).
 */
function withSheetLock_(fn) {
  // No multiusuário o lock de SCRIPT é compartilhado por todo mundo (uma pessoa travaria
  // as outras); o de USUÁRIO serializa só as execuções da mesma pessoa/planilha.
  var lock = MULTIUSER_MODE ? LockService.getUserLock() : LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/** Adiciona uma linha; preenche ID e CreatedAt se ausentes. Retorna o objeto salvo. */
function appendRow_(sheetName, obj) {
  return withSheetLock_(function () {
    var sheet = getOrCreateSheet_(sheetName);
    var headers = SCHEMA[sheetName];
    if (!obj.ID) obj.ID = generateId_();
    if (headers.indexOf('CreatedAt') !== -1 && !obj.CreatedAt) obj.CreatedAt = nowIso_();
    sheet.appendRow(objectToRow_(headers, obj));
    return obj;
  });
}

/**
 * Adiciona VÁRIAS linhas de uma vez (um setValues só, dentro de UM lock só) — usado
 * quando dezenas de linhas precisam ser criadas juntas (ex: importação de documento).
 * appendRow_ chamada uma vez por linha paga o custo de adquirir/soltar o lock e gravar
 * a cada chamada; pra dezenas de linhas isso se acumula rápido e foi a causa real por
 * trás de "Criar pendências" ficar lento mesmo depois de já ter parado de reler
 * Transactions a cada item. Retorna os objetos salvos (com ID/CreatedAt preenchidos).
 */
function appendRowsBatch_(sheetName, objs) {
  if (!objs.length) return [];
  return withSheetLock_(function () {
    var sheet = getOrCreateSheet_(sheetName);
    var headers = SCHEMA[sheetName];
    var hasCreatedAt = headers.indexOf('CreatedAt') !== -1;
    objs.forEach(function (obj) {
      if (!obj.ID) obj.ID = generateId_();
      if (hasCreatedAt && !obj.CreatedAt) obj.CreatedAt = nowIso_();
    });
    var startRow = sheet.getLastRow() + 1;
    var values = objs.map(function (obj) { return objectToRow_(headers, obj); });
    sheet.getRange(startRow, 1, values.length, headers.length).setValues(values);
    return objs;
  });
}

/** Atualiza a linha cujo ID bate, mesclando os campos passados. Retorna o objeto atualizado ou null. */
function updateRowById_(sheetName, id, updates) {
  return withSheetLock_(function () {
    var sheet = getOrCreateSheet_(sheetName);
    var headers = SCHEMA[sheetName];
    var idCol = headers.indexOf('ID') + 1;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    var ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] === id) {
        var rowIndex = i + 2;
        var currentRow = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
        var obj = normalizeTextColumns_(sheetName, rowToObject_(headers, currentRow));
        Object.keys(updates).forEach(function (k) { obj[k] = updates[k]; });
        sheet.getRange(rowIndex, 1, 1, headers.length).setValues([objectToRow_(headers, obj)]);
        return obj;
      }
    }
    return null;
  });
}

/** Remove a linha cujo ID bate. Retorna true se removeu. */
function deleteRowById_(sheetName, id) {
  return withSheetLock_(function () {
    var sheet = getOrCreateSheet_(sheetName);
    var headers = SCHEMA[sheetName];
    var idCol = headers.indexOf('ID') + 1;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    var ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] === id) {
        sheet.deleteRow(i + 2);
        return true;
      }
    }
    return false;
  });
}

function findById_(sheetName, id) {
  var all = readAll_(sheetName);
  for (var i = 0; i < all.length; i++) {
    if (all[i].ID === id) return all[i];
  }
  return null;
}

// ===== Transactions.js =====
/**
 * Lançamentos: receitas, despesas (com fatura de cartão e parcelamento) e transferências.
 *
 * Convenção de sinal (ver amountSignForKind_ em Accounts.js):
 *   income: +   expense: -   transfer_out: -   transfer_in: +
 * Amount é sempre armazenado positivo; o Kind define o efeito no saldo.
 */

function pad2_(n) { return (n < 10 ? '0' : '') + n; }

function parseYmd_(ymd) {
  var parts = ymd.split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function formatYm_(year, monthIndex0) {
  // Normaliza monthIndex0 (0-11) que pode ter estourado a faixa ao somar meses.
  var d = new Date(year, monthIndex0, 1);
  return d.getFullYear() + '-' + pad2_(d.getMonth() + 1);
}

/**
 * Fatura a que uma compra pertence: se o dia da compra é <= dia de fechamento,
 * a compra entra na fatura que fecha no mesmo mês; senão, na do mês seguinte.
 */
function billMonthForDate_(dateYmd, closingDay) {
  var d = parseYmd_(dateYmd);
  var day = d.getDate();
  if (day <= closingDay) {
    return formatYm_(d.getFullYear(), d.getMonth());
  }
  return formatYm_(d.getFullYear(), d.getMonth() + 1);
}

function addMonthsToYm_(ym, months) {
  var parts = ym.split('-').map(Number);
  return formatYm_(parts[0], parts[1] - 1 + months);
}

/**
 * Data de vencimento de uma fatura. Cartões brasileiros tipicamente fecham num dia
 * e vencem num dia posterior — se dueDay < closingDay, o vencimento cai no mês
 * seguinte ao mês de fechamento (ex.: fecha dia 25, vence dia 5 do mês seguinte).
 */
function dueDateForBillMonth_(billMonth, closingDay, dueDay) {
  var parts = billMonth.split('-').map(Number);
  var year = parts[0], monthIndex0 = parts[1] - 1;
  if (dueDay < closingDay) monthIndex0 += 1;
  return new Date(year, monthIndex0, dueDay);
}

function splitAmount_(totalAmount, installments) {
  var cents = Math.round(totalAmount * 100);
  var base = Math.floor(cents / installments);
  var remainder = cents - base * installments;
  var parts = [];
  for (var i = 0; i < installments; i++) {
    var partCents = base + (i < remainder ? 1 : 0);
    parts.push(partCents / 100);
  }
  return parts;
}

/**
 * Linha bruta da aba (schema PascalCase) -> formato que o cliente usa em toda parte
 * (STATE.allTransactions). Compartilhado entre listTransactions e os create* — sem isso,
 * create* devolvia a linha crua, e o cliente só via o lançamento de verdade quando o
 * refresh coalescido (uma leitura SEPARADA da planilha, sujeita a atraso logo depois de
 * um appendRow_) chegasse — com o create* já devolvendo mapeado, o cliente pode confiar
 * direto na PRÓPRIA resposta da criação e mostrar na hora, sem esperar nem arriscar ler
 * uma planilha ainda desatualizada.
 */
function mapTransactionRow_(t) {
  return {
    id: t.ID, date: t.Date, time: t.Time || null, accountId: t.AccountID || null, categoryId: t.CategoryID || null,
    kind: t.Kind, amount: Number(t.Amount), description: t.Description, counterparty: t.Counterparty || null,
    billMonth: t.BillMonth || null,
    installmentSeq: t.InstallmentSeq === '' ? null : Number(t.InstallmentSeq),
    installmentTotal: t.InstallmentTotal === '' ? null : Number(t.InstallmentTotal),
    installmentGroupId: t.InstallmentGroupID || null,
    transferGroupId: t.TransferGroupID || null,
    forecastId: t.ForecastID || null,
    // 'YYYY-MM' (ciclo mensal do recorrente vinculado, quando diferente do mês da data) ou
    // null -- ver SCHEMA.Transactions.
    forecastCycleMonth: t.ForecastCycleMonth || null,
    excludeFromAnalysis: parseBoolCell_(t.ExcludeFromAnalysis)
  };
}

/** categoryId vazio é permitido (lançamento "Sem categoria") — só valida quando informado. */
function assertKindMatchesCategory_(kind, categoryId) {
  if (!categoryId) return;
  var category = findById_('Categories', categoryId);
  if (!category) throw new Error('Categoria não encontrada.');
  if (category.Kind !== kind) {
    throw new Error('A categoria "' + category.Name + '" é de ' + category.Kind + ', incompatível com lançamento de ' + kind + '.');
  }
}

/**
 * Cria um lançamento de receita ou despesa. Conta e categoria são opcionais — um
 * lançamento sem conta não afeta saldo de nenhuma conta (fica "Sem conta" até ser
 * ajustado depois). Se a conta for cartão de crédito e installments > 1, gera N
 * linhas (uma por parcela) distribuídas nas faturas seguintes.
 */
function createTransaction(input) {
  if (input.kind !== 'income' && input.kind !== 'expense') {
    throw new Error('Kind inválido para lançamento: ' + input.kind);
  }
  if (!input.date) throw new Error('Data é obrigatória.');
  var amount = Number(input.amount);
  if (!amount || amount <= 0) throw new Error('Valor precisa ser maior que zero.');

  assertKindMatchesCategory_(input.kind, input.categoryId);

  if (!input.accountId) throw new Error('Selecione uma conta.');
  var account = findById_('Accounts', input.accountId);
  if (!account) throw new Error('Conta não encontrada.');
  if (account.Type === 'investment') {
    throw new Error('Conta de investimento só aceita transferências (depósito/resgate), não receita ou despesa direta.');
  }

  var installments = Math.max(1, Number(input.installments || 1));
  var isCreditCard = !!account && account.Type === 'credit_card';
  if (installments > 1 && !isCreditCard) {
    throw new Error('Parcelamento só é permitido em contas de cartão de crédito.');
  }

  var baseBillMonth = isCreditCard ? billMonthForDate_(input.date, Number(account.ClosingDay)) : '';
  var groupId = installments > 1 ? generateId_() : '';
  var amounts = splitAmount_(amount, installments);

  var created = [];
  for (var i = 0; i < installments; i++) {
    var billMonth = isCreditCard ? addMonthsToYm_(baseBillMonth, i) : '';
    var description = input.description || '';
    if (installments > 1) {
      description += (description ? ' ' : '') + '(' + (i + 1) + '/' + installments + ')';
    }
    created.push(appendRow_('Transactions', {
      Date: input.date,
      Time: input.time || '',
      AccountID: input.accountId || '',
      CategoryID: input.categoryId || '',
      Kind: input.kind,
      Amount: amounts[i],
      Description: description,
      Counterparty: input.counterparty || '',
      BillMonth: billMonth,
      InstallmentGroupID: groupId,
      InstallmentSeq: installments > 1 ? i + 1 : '',
      InstallmentTotal: installments > 1 ? installments : '',
      TransferGroupID: '',
      ForecastID: input.forecastId || '',
      ForecastCycleMonth: input.forecastId ? (input.forecastCycleMonth || '') : '',
      ExcludeFromAnalysis: !!input.excludeFromAnalysis
    }));
  }

  if (input.rememberCategoryForCounterparty && input.counterparty && input.categoryId) {
    saveCounterpartyMapping_(input.counterparty, input.categoryId);
  }

  return created.map(mapTransactionRow_);
}

function updateTransaction(input) {
  if (!input.id) throw new Error('ID do lançamento é obrigatório.');
  var existing = findById_('Transactions', input.id);
  if (!existing) throw new Error('Lançamento não encontrado.');
  if (existing.Kind === 'transfer_in' || existing.Kind === 'transfer_out') {
    throw new Error('Transferências têm campos diferentes (contas de origem/destino, sem categoria) — use updateTransfer.');
  }
  var updates = {};
  if (input.date !== undefined) updates.Date = input.date;
  if (input.time !== undefined) updates.Time = input.time || '';
  if (input.categoryId !== undefined) {
    assertKindMatchesCategory_(existing.Kind, input.categoryId);
    updates.CategoryID = input.categoryId || '';
  }
  if (input.accountId !== undefined) {
    if (!input.accountId) throw new Error('Selecione uma conta.');
    var newAccount = findById_('Accounts', input.accountId);
    if (!newAccount) throw new Error('Conta não encontrada.');
    if (newAccount.Type === 'investment') throw new Error('Conta de investimento só aceita transferências.');
    updates.AccountID = input.accountId;
  }
  if (input.amount !== undefined) updates.Amount = Number(input.amount);
  if (input.description !== undefined) updates.Description = input.description;
  if (input.counterparty !== undefined) updates.Counterparty = input.counterparty || '';
  if (input.forecastId !== undefined) updates.ForecastID = input.forecastId || '';
  // O mês do ciclo acompanha o vínculo: sem recorrente vinculado não sobra ciclo nenhum
  // (desvincular zera os dois juntos); trocar de recorrente sem informar o mês também
  // zera (o mês guardado era do recorrente ANTERIOR); só mantém o guardado quando o
  // recorrente continua o mesmo e o mês não veio no pedido.
  if (input.forecastCycleMonth !== undefined || input.forecastId !== undefined) {
    var effectiveForecastId = updates.ForecastID !== undefined ? updates.ForecastID : existing.ForecastID;
    var cycleMonth;
    if (!effectiveForecastId) cycleMonth = '';
    else if (input.forecastCycleMonth !== undefined) cycleMonth = input.forecastCycleMonth || '';
    else cycleMonth = effectiveForecastId === existing.ForecastID ? (existing.ForecastCycleMonth || '') : '';
    updates.ForecastCycleMonth = cycleMonth;
  }
  if (input.excludeFromAnalysis !== undefined) updates.ExcludeFromAnalysis = !!input.excludeFromAnalysis;

  var effectiveAccountId = updates.AccountID !== undefined ? updates.AccountID : existing.AccountID;
  var effectiveAccount = effectiveAccountId ? findById_('Accounts', effectiveAccountId) : null;
  if (updates.Date && effectiveAccount && effectiveAccount.Type === 'credit_card' && !existing.InstallmentGroupID) {
    updates.BillMonth = billMonthForDate_(updates.Date, Number(effectiveAccount.ClosingDay));
  }
  var updated = updateRowById_('Transactions', input.id, updates);
  // Ver cascadeRefundsOnTransactionAmountChange_ (Refunds.js): desfaz vínculos de
  // reembolso que não cabem mais no novo valor (só do lado ATIVO). Compara com o valor
  // ANTERIOR (não só "updates.Amount !== undefined") porque o cliente manda amount em
  // TODO Salvar, tenha mudado ou não — sem essa comparação, isso rodava (inofensivo pro
  // ativo, mas achou-se um bug real quando isso também mexia no lado passivo) a cada
  // edição de qualquer campo, não só quando o valor de fato muda.
  if (updates.Amount !== undefined && Math.round(Number(updates.Amount) * 100) !== Math.round(Number(existing.Amount) * 100)) {
    cascadeRefundsOnTransactionAmountChange_(input.id, updates.Amount);
  }
  // Mesma ideia de createTransaction (ver ali) -- usa o contraparte/categoria EFETIVOS
  // (o que acabou de ser salvo, não o que existia antes), já que editar e marcar
  // "lembrar" na mesma tela é o caso comum.
  if (input.rememberCategoryForCounterparty) {
    var effectiveCounterparty = updates.Counterparty !== undefined ? updates.Counterparty : existing.Counterparty;
    var effectiveCategoryId = updates.CategoryID !== undefined ? updates.CategoryID : existing.CategoryID;
    if (effectiveCounterparty && effectiveCategoryId) saveCounterpartyMapping_(effectiveCounterparty, effectiveCategoryId);
  }
  return updated;
}

function deleteTransaction(id) {
  var existing = findById_('Transactions', id);
  if (!existing) throw new Error('Lançamento não encontrado.');
  if (existing.Kind === 'transfer_in' || existing.Kind === 'transfer_out') {
    return deleteTransfer(existing.TransferGroupID);
  }
  cascadeRefundsOnTransactionDelete_(id);
  return deleteRowById_('Transactions', id);
}

/** Exclui todas as parcelas futuras (>= seq atual) de um grupo de parcelamento. */
function deleteInstallmentGroupFrom(groupId, fromSeq) {
  var txs = readAll_('Transactions').filter(function (t) {
    return t.InstallmentGroupID === groupId && Number(t.InstallmentSeq) >= Number(fromSeq);
  });
  txs.forEach(function (t) { deleteRowById_('Transactions', t.ID); });
  return txs.length;
}

/**
 * Exclui TODAS as parcelas de um grupo, não só uma — a ação padrão ao remover um
 * lançamento parcelado (ver openTransactionDetailModal, JavaScript.html), já que todas
 * representam a MESMA compra. "Remover só esta parcela" continua disponível como opção
 * secundária, chamando deleteTransaction numa parcela só, sem tocar aqui.
 */
function deleteInstallmentGroup(groupId) {
  if (!groupId) throw new Error('InstallmentGroupID é obrigatório.');
  var txs = readAll_('Transactions').filter(function (t) { return t.InstallmentGroupID === groupId; });
  if (!txs.length) throw new Error('Parcelamento não encontrado.');
  txs.forEach(function (t) { deleteRowById_('Transactions', t.ID); });
  return txs.length;
}

/**
 * Transferência entre contas (inclui pagamento de fatura de cartão).
 * billMonth (opcional): quando o destino é um cartão de crédito, marca a que
 * fatura esse pagamento se refere, para o resumo de fatura bater certo.
 */
function createTransfer(input) {
  if (!input.date) throw new Error('Data é obrigatória.');
  if (!input.fromAccountId || !input.toAccountId) throw new Error('Contas de origem e destino são obrigatórias.');
  if (input.fromAccountId === input.toAccountId) throw new Error('Conta de origem e destino não podem ser a mesma.');
  var amount = Number(input.amount);
  if (!amount || amount <= 0) throw new Error('Valor precisa ser maior que zero.');

  var fromAccount = findById_('Accounts', input.fromAccountId);
  var toAccount = findById_('Accounts', input.toAccountId);
  if (!fromAccount || !toAccount) throw new Error('Conta não encontrada.');

  var groupId = generateId_();
  var description = input.description || '';
  var billMonth = (toAccount.Type === 'credit_card' && input.billMonth) ? input.billMonth : '';

  var out = appendRow_('Transactions', {
    Date: input.date, Time: input.time || '', AccountID: input.fromAccountId, CategoryID: '',
    Kind: 'transfer_out', Amount: amount, Description: description,
    BillMonth: '', InstallmentGroupID: '', InstallmentSeq: '', InstallmentTotal: '',
    TransferGroupID: groupId
  });
  var incoming = appendRow_('Transactions', {
    Date: input.date, Time: input.time || '', AccountID: input.toAccountId, CategoryID: '',
    Kind: 'transfer_in', Amount: amount, Description: description,
    BillMonth: billMonth, InstallmentGroupID: '', InstallmentSeq: '', InstallmentTotal: '',
    TransferGroupID: groupId
  });
  return { out: mapTransactionRow_(out), in: mapTransactionRow_(incoming) };
}

/**
 * Edita uma transferência já criada — ao contrário de updateTransaction (que bloqueia
 * transferência de propósito, ver lá), aqui SEMPRE atualiza as DUAS pernas
 * (transfer_out/transfer_in) juntas, então não corre o risco de uma ficar
 * dessincronizada da outra. Qualquer campo não informado preserva o valor atual —
 * mesma convenção de updateTransaction — exceto billMonth: como ele só existe pra
 * pernas de entrada em cartão de crédito, é recalculado a cada chamada a partir do
 * tipo ATUAL da conta de destino (mesma regra de createTransfer), não só preservado.
 */
function updateTransfer(input) {
  if (!input.transferGroupId) throw new Error('TransferGroupID é obrigatório.');
  var legs = readAll_('Transactions').filter(function (t) { return t.TransferGroupID === input.transferGroupId; });
  var outLeg = legs.find(function (t) { return t.Kind === 'transfer_out'; });
  var inLeg = legs.find(function (t) { return t.Kind === 'transfer_in'; });
  if (!outLeg || !inLeg) throw new Error('Transferência não encontrada.');

  var fromAccountId = input.fromAccountId !== undefined ? input.fromAccountId : outLeg.AccountID;
  var toAccountId = input.toAccountId !== undefined ? input.toAccountId : inLeg.AccountID;
  if (!fromAccountId || !toAccountId) throw new Error('Contas de origem e destino são obrigatórias.');
  if (fromAccountId === toAccountId) throw new Error('Conta de origem e destino não podem ser a mesma.');
  var fromAccount = findById_('Accounts', fromAccountId);
  var toAccount = findById_('Accounts', toAccountId);
  if (!fromAccount || !toAccount) throw new Error('Conta não encontrada.');

  var amount = input.amount !== undefined ? Number(input.amount) : Number(outLeg.Amount);
  if (!amount || amount <= 0) throw new Error('Valor precisa ser maior que zero.');

  var date = input.date !== undefined ? input.date : outLeg.Date;
  var time = input.time !== undefined ? (input.time || '') : (outLeg.Time || '');
  var description = input.description !== undefined ? (input.description || '') : (outLeg.Description || '');

  // Só existe billMonth se o destino é (continuar sendo) cartão de crédito — mesma
  // regra de createTransfer. Se o destino deixou de ser cartão, limpa; se continua
  // cartão, usa o que foi mandado (mesmo vazio, pra "desvincular" funcionar).
  var billMonth = (toAccount.Type === 'credit_card' && input.billMonth) ? input.billMonth : '';

  updateRowById_('Transactions', outLeg.ID, {
    Date: date, Time: time, AccountID: fromAccountId, Amount: amount, Description: description
  });
  updateRowById_('Transactions', inLeg.ID, {
    Date: date, Time: time, AccountID: toAccountId, Amount: amount, Description: description, BillMonth: billMonth
  });
  return true;
}

function deleteTransfer(transferGroupId) {
  if (!transferGroupId) throw new Error('TransferGroupID é obrigatório.');
  var txs = readAll_('Transactions').filter(function (t) { return t.TransferGroupID === transferGroupId; });
  txs.forEach(function (t) { deleteRowById_('Transactions', t.ID); });
  return txs.length === 2;
}

/** Lista lançamentos com filtros opcionais, mais recentes primeiro. */
function listTransactions(filters, allTransactions) {
  filters = filters || {};
  var txs = (allTransactions || readAll_('Transactions')).filter(function (t) {
    if (filters.accountId && t.AccountID !== filters.accountId) return false;
    if (filters.categoryId && t.CategoryID !== filters.categoryId) return false;
    if (filters.kind && t.Kind !== filters.kind) return false;
    if (filters.billMonth && t.BillMonth !== filters.billMonth) return false;
    if (filters.dateFrom && t.Date < filters.dateFrom) return false;
    if (filters.dateTo && t.Date > filters.dateTo) return false;
    return true;
  });
  // Desempata por Time (tratando "sem horário" como vazio, pra não cair direto no
  // CreatedAt quando só um dos dois lançamentos tem horário informado) e por fim por
  // CreatedAt, senão a ordem dentro do mesmo dia/horário ficaria na ordem de inserção
  // da planilha (mais antigo primeiro).
  txs.sort(function (a, b) {
    if (a.Date !== b.Date) return a.Date < b.Date ? 1 : -1;
    var aTime = a.Time || '', bTime = b.Time || '';
    if (aTime !== bTime) return aTime < bTime ? 1 : -1;
    return a.CreatedAt < b.CreatedAt ? 1 : (a.CreatedAt > b.CreatedAt ? -1 : 0);
  });
  return txs.map(mapTransactionRow_);
}

/** Resumo das faturas (abertas, fechadas, pagas) de uma conta de cartão de crédito. */
function getCreditCardBillSummary(accountId, allTransactions) {
  var account = findById_('Accounts', accountId);
  if (!account || account.Type !== 'credit_card') throw new Error('Conta não é cartão de crédito.');
  var closingDay = Number(account.ClosingDay), dueDay = Number(account.DueDay);
  var currentBillMonth = billMonthForDate_(todayYmd_(), closingDay);

  var txs = (allTransactions || readAll_('Transactions')).filter(function (t) { return t.AccountID === accountId; });
  var byMonth = {};
  txs.forEach(function (t) {
    if (!t.BillMonth) return;
    if (!byMonth[t.BillMonth]) byMonth[t.BillMonth] = { spent: 0, paid: 0 };
    if (t.Kind === 'expense') byMonth[t.BillMonth].spent += Number(t.Amount);
    if (t.Kind === 'transfer_in') byMonth[t.BillMonth].paid += Number(t.Amount);
  });

  var today = parseYmd_(todayYmd_());
  return Object.keys(byMonth).sort().map(function (billMonth) {
    var data = byMonth[billMonth];
    var dueDate = dueDateForBillMonth_(billMonth, closingDay, dueDay);
    var isOpen = billMonth === currentBillMonth;
    var isPaid = data.paid >= data.spent && data.spent > 0;
    var status = isOpen ? 'aberta' : (isPaid ? 'paga' : (dueDate < today ? 'atrasada' : 'fechada'));
    return {
      billMonth: billMonth, spent: data.spent, paid: data.paid,
      remaining: Math.max(0, data.spent - data.paid),
      dueDate: Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      status: status
    };
  });
}

function todayYmd_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** Hora atual em HH:mm — usada pelos lançamentos que o próprio sistema cria (saldo
 *  inicial e consolidação de saldo, ver Accounts.js), pra ficarem rastreáveis no
 *  horário exato em que a conta foi criada/editada. */
function nowHm_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
}

return {
  archiveAccount: (typeof archiveAccount === 'function' ? archiveAccount : function () { throw new Error('"archiveAccount" não existe na versão web.'); }),
  archiveCategory: (typeof archiveCategory === 'function' ? archiveCategory : function () { throw new Error('"archiveCategory" não existe na versão web.'); }),
  confirmDocumentImportBatch: (typeof confirmDocumentImportBatch === 'function' ? confirmDocumentImportBatch : function () { throw new Error('"confirmDocumentImportBatch" não existe na versão web.'); }),
  confirmPendingEntry: (typeof confirmPendingEntry === 'function' ? confirmPendingEntry : function () { throw new Error('"confirmPendingEntry" não existe na versão web.'); }),
  createAccount: (typeof createAccount === 'function' ? createAccount : function () { throw new Error('"createAccount" não existe na versão web.'); }),
  createBudget: (typeof createBudget === 'function' ? createBudget : function () { throw new Error('"createBudget" não existe na versão web.'); }),
  createCategory: (typeof createCategory === 'function' ? createCategory : function () { throw new Error('"createCategory" não existe na versão web.'); }),
  createForecast: (typeof createForecast === 'function' ? createForecast : function () { throw new Error('"createForecast" não existe na versão web.'); }),
  createPerson: (typeof createPerson === 'function' ? createPerson : function () { throw new Error('"createPerson" não existe na versão web.'); }),
  createPlan: (typeof createPlan === 'function' ? createPlan : function () { throw new Error('"createPlan" não existe na versão web.'); }),
  createTransaction: (typeof createTransaction === 'function' ? createTransaction : function () { throw new Error('"createTransaction" não existe na versão web.'); }),
  createTransfer: (typeof createTransfer === 'function' ? createTransfer : function () { throw new Error('"createTransfer" não existe na versão web.'); }),
  deleteInstallmentGroup: (typeof deleteInstallmentGroup === 'function' ? deleteInstallmentGroup : function () { throw new Error('"deleteInstallmentGroup" não existe na versão web.'); }),
  deleteTransaction: (typeof deleteTransaction === 'function' ? deleteTransaction : function () { throw new Error('"deleteTransaction" não existe na versão web.'); }),
  discardPendingEntry: (typeof discardPendingEntry === 'function' ? discardPendingEntry : function () { throw new Error('"discardPendingEntry" não existe na versão web.'); }),
  dismissBudgetModeSwitchSuggestion: (typeof dismissBudgetModeSwitchSuggestion === 'function' ? dismissBudgetModeSwitchSuggestion : function () { throw new Error('"dismissBudgetModeSwitchSuggestion" não existe na versão web.'); }),
  dismissBudgetSuggestion: (typeof dismissBudgetSuggestion === 'function' ? dismissBudgetSuggestion : function () { throw new Error('"dismissBudgetSuggestion" não existe na versão web.'); }),
  dismissForecastSuggestion: (typeof dismissForecastSuggestion === 'function' ? dismissForecastSuggestion : function () { throw new Error('"dismissForecastSuggestion" não existe na versão web.'); }),
  exportTransactionsCsv: (typeof exportTransactionsCsv === 'function' ? exportTransactionsCsv : function () { throw new Error('"exportTransactionsCsv" não existe na versão web.'); }),
  getBootstrapData: (typeof getBootstrapData === 'function' ? getBootstrapData : function () { throw new Error('"getBootstrapData" não existe na versão web.'); }),
  getPhoneConnectionInfo: (typeof getPhoneConnectionInfo === 'function' ? getPhoneConnectionInfo : function () { throw new Error('"getPhoneConnectionInfo" não existe na versão web.'); }),
  ignoreForecastCycle: (typeof ignoreForecastCycle === 'function' ? ignoreForecastCycle : function () { throw new Error('"ignoreForecastCycle" não existe na versão web.'); }),
  previewDocumentImport: (typeof previewDocumentImport === 'function' ? previewDocumentImport : function () { throw new Error('"previewDocumentImport" não existe na versão web.'); }),
  reconnectUserSpreadsheet: (typeof reconnectUserSpreadsheet === 'function' ? reconnectUserSpreadsheet : function () { throw new Error('"reconnectUserSpreadsheet" não existe na versão web.'); }),
  recreateUserSpreadsheet: (typeof recreateUserSpreadsheet === 'function' ? recreateUserSpreadsheet : function () { throw new Error('"recreateUserSpreadsheet" não existe na versão web.'); }),
  removeBalanceGoal: (typeof removeBalanceGoal === 'function' ? removeBalanceGoal : function () { throw new Error('"removeBalanceGoal" não existe na versão web.'); }),
  removeBudget: (typeof removeBudget === 'function' ? removeBudget : function () { throw new Error('"removeBudget" não existe na versão web.'); }),
  removeForecast: (typeof removeForecast === 'function' ? removeForecast : function () { throw new Error('"removeForecast" não existe na versão web.'); }),
  removePlan: (typeof removePlan === 'function' ? removePlan : function () { throw new Error('"removePlan" não existe na versão web.'); }),
  setBalanceGoal: (typeof setBalanceGoal === 'function' ? setBalanceGoal : function () { throw new Error('"setBalanceGoal" não existe na versão web.'); }),
  setBudgetViewPreference: (typeof setBudgetViewPreference === 'function' ? setBudgetViewPreference : function () { throw new Error('"setBudgetViewPreference" não existe na versão web.'); }),
  setRefundLinksForActiveTransaction: (typeof setRefundLinksForActiveTransaction === 'function' ? setRefundLinksForActiveTransaction : function () { throw new Error('"setRefundLinksForActiveTransaction" não existe na versão web.'); }),
  setRefundSharesForPassiveTransaction: (typeof setRefundSharesForPassiveTransaction === 'function' ? setRefundSharesForPassiveTransaction : function () { throw new Error('"setRefundSharesForPassiveTransaction" não existe na versão web.'); }),
  unignoreForecastCycle: (typeof unignoreForecastCycle === 'function' ? unignoreForecastCycle : function () { throw new Error('"unignoreForecastCycle" não existe na versão web.'); }),
  updateAccount: (typeof updateAccount === 'function' ? updateAccount : function () { throw new Error('"updateAccount" não existe na versão web.'); }),
  updateBudget: (typeof updateBudget === 'function' ? updateBudget : function () { throw new Error('"updateBudget" não existe na versão web.'); }),
  updateCategory: (typeof updateCategory === 'function' ? updateCategory : function () { throw new Error('"updateCategory" não existe na versão web.'); }),
  updateForecast: (typeof updateForecast === 'function' ? updateForecast : function () { throw new Error('"updateForecast" não existe na versão web.'); }),
  updatePlan: (typeof updatePlan === 'function' ? updatePlan : function () { throw new Error('"updatePlan" não existe na versão web.'); }),
  updateTransaction: (typeof updateTransaction === 'function' ? updateTransaction : function () { throw new Error('"updateTransaction" não existe na versão web.'); }),
  updateTransfer: (typeof updateTransfer === 'function' ? updateTransfer : function () { throw new Error('"updateTransfer" não existe na versão web.'); }),
  seedSpreadsheet_: seedSpreadsheet_,
  __SCHEMA: SCHEMA,
  __TEXT_COLUMNS: TEXT_COLUMNS
};
})(CacauGAS.Utilities, CacauGAS.Session, CacauGAS.LockService, CacauGAS.PropertiesService, CacauGAS.SpreadsheetApp, CacauGAS.Logger);

/**
 * Backend do Cacau no navegador (protótipo, etapa 1): faz o papel do google.script.run.
 *
 *  1. Login com Google (Google Identity Services) pedindo só o escopo de Planilhas + e-mail.
 *  2. Antes de cada chamada, baixa um snapshot da planilha pela Sheets API (CacauGAS o
 *     expõe como SpreadsheetApp) e roda a MESMA função do servidor (CacauServer) em cima dele.
 *  3. Mede tudo (tempo, chamadas à API, validade do token) num painel de diagnóstico.
 *
 * Duas permissões (escolhidas na tela de entrada, guardadas no navegador):
 *   file  (padrão) escopo "drive.file": só enxerga planilhas que o próprio app criou ou que a pessoa escolheu
 *         no seletor do Google (Picker). Decidido: a mínima é a padrão.
 *   full  escopo "spreadsheets": enxerga qualquer planilha da conta (sensível: exige verificação do Google)
 *
 * Duas modalidades (interruptor "modo gravação", guardado no navegador):
 *   leitura   (padrão) gravações do servidor são recusadas; leituras que "arrumam" a planilha
 *             valem só na memória
 *   gravação  o que o servidor fez na cópia em memória (CacauGAS.pendingOps) é enviado à Sheets
 *             API por CacauWriter, depois de conferir que a planilha não mudou em outro lugar
 * As chamadas rodam uma de cada vez (fila), como o LockService fazia no Apps Script.
 * Quem usa isto é o run() do cliente (JavaScript.html), que só delega para cá quando
 * window.CacauBackend existe.
 */
var CacauBackend = (function () {
  var cfg = window.CACAU_WEB_CONFIG || {};
  var SCOPES = {
    full: 'openid email https://www.googleapis.com/auth/spreadsheets',
    file: 'openid email https://www.googleapis.com/auth/drive.file'
  };
  var SCOPE_KEY = 'cacau-web-scope';
  var SHEET_KEY = 'cacau-web-sheet-id';
  var DIAG_KEY = 'cacau-web-diag-hidden';
  var META_TTL_MS = 5 * 60 * 1000;
  var SILENT_TIMEOUT_MS = 8000;
  var TOKEN_MARGIN_MS = 60 * 1000;
  var NATIVE_TOKEN_MS = 30 * 60 * 1000; // o app não informa a validade; o Google dá ~1 h, então 30 min é conservador
  var NATIVE_ASK_TIMEOUT_MS = 3 * 60 * 1000;
  var SHEETS = 'https://sheets.googleapis.com/v4/spreadsheets/';
  var TOKEN_KEY = 'cacau-web-token';
  var HINT_KEY = 'cacau-web-hint'; // { email, mode }: quem já entrou aqui (pra não perguntar a conta de novo)
  var WRITE_KEY = 'cacau-web-write';
  var CONFLICT_MESSAGE = 'A planilha foi alterada em outro lugar (outro aparelho ou aba) enquanto você editava. Nada foi gravado: recarregue e tente de novo.';

  var tokenClient = null;
  var tokenClientMode = '';
  var pendingToken = null; // { resolve, reject }
  var token = { value: '', expiresAt: 0, grantedAt: 0 };
  var everSignedIn = false;
  var meta = null;
  var metaAt = 0;
  var userGate = null;     // promessa que só resolve quando a pessoa clica em "Entrar" de novo
  var gateDone = null;
  var stats = { apiCalls: 0, calls: [], events: [] };
  var queue = Promise.resolve();

  /** Dentro do app Android o login é do próprio app (Google bloqueia login em WebView): ele entrega o token pela ponte. */
  function nativeLogin_() {
    return !!(window.FinancasApp && typeof window.FinancasApp.requestWebToken === 'function');
  }

  function scopeMode_() {
    if (nativeLogin_()) return 'file';
    try { return localStorage.getItem(SCOPE_KEY) === 'full' ? 'full' : 'file'; } catch (e) { return 'file'; }
  }
  /** Trocar de permissão invalida o token atual (foi emitido para outro escopo). */
  function setScopeMode_(mode) {
    if (mode === scopeMode_()) return;
    try { if (mode === 'full') localStorage.setItem(SCOPE_KEY, 'full'); else localStorage.removeItem(SCOPE_KEY); } catch (e) { /* segue */ }
    tokenClient = null;
    token = { value: '', expiresAt: 0, grantedAt: 0 };
    everSignedIn = false;
    forgetToken_();
  }

  function writeMode_() {
    try { return localStorage.getItem(WRITE_KEY) !== '0'; } catch (e) { return true; }
  }
  function setWriteMode_(on) {
    try { localStorage.setItem(WRITE_KEY, on ? '1' : '0'); } catch (e) { /* segue */ }
  }

  var ready = new Promise(function (resolve) { gateDone = resolve; });

  // ---------- diagnóstico ----------
  function event_(message) {
    stats.events.unshift({ at: Date.now(), message: message });
    if (stats.events.length > 12) stats.events.length = 12;
    renderDiag_();
  }
  function hhmmss_(ms) {
    var d = new Date(ms);
    return [d.getHours(), d.getMinutes(), d.getSeconds()].map(function (n) { return n < 10 ? '0' + n : String(n); }).join(':');
  }
  function renderDiag_() {
    var box = document.getElementById('cacau-diag');
    if (!box) return;
    var left = token.expiresAt ? Math.max(0, Math.round((token.expiresAt + TOKEN_MARGIN_MS - Date.now()) / 1000)) : 0;
    var lines = ['<b>Diagnóstico (protótipo)</b> — modo ' + (writeMode_() ? '<b style="color:#fc6">GRAVAÇÃO</b>' : 'leitura') + ' · permissão ' + (scopeMode_() === 'file' ? '<b style="color:#9f9">mínima (drive.file)</b>' : 'completa') + ' · token vale mais ' + left + ' s · ' + stats.apiCalls + ' chamadas à API'];
    stats.calls.slice(0, 5).forEach(function (c) {
      lines.push(hhmmss_(c.at) + ' ' + c.fn + ' · ' + c.ms + ' ms (rede ' + c.netMs + ' ms, ' + c.api + ' req, ' + Math.round(c.bytes / 1024) + ' KB)' + (c.writes ? ' · ' + c.writes + ' gravação(ões)' : '') + (c.error ? ' · ERRO ' + c.error : ''));
    });
    stats.events.slice(0, 5).forEach(function (e) { lines.push(hhmmss_(e.at) + ' ' + e.message); });
    box.querySelector('.d-body').innerHTML = lines.join('<br>');
  }
  /** Painel de diagnóstico + um botão discreto (sempre presente) que mostra/oculta o painel; a escolha fica guardada. */
  function mountDiag_() {
    var hidden = false;
    hidden = true;
    try { hidden = localStorage.getItem(DIAG_KEY) !== '0'; } catch (e) { /* segue */ }
    var box = document.createElement('div');
    box.id = 'cacau-diag';
    box.setAttribute('style', 'position:fixed;left:8px;bottom:40px;z-index:200;max-width:92vw;background:rgba(0,0,0,.78);color:#fff;font:11px/1.35 monospace;padding:6px 8px;border-radius:8px');
    box.innerHTML = '<div class="d-body"></div><a href="#" class="d-x" style="color:#9cf">ocultar</a> · <a href="#" class="d-mode" style="color:#9cf">trocar modo</a> · <a href="#" class="d-sheet" style="color:#9cf">trocar planilha</a> · <a href="#" class="d-out" style="color:#9cf">sair da conta</a>';
    var toggle = document.createElement('button');
    toggle.id = 'cacau-diag-toggle';
    toggle.type = 'button';
    toggle.title = 'Mostrar/ocultar o diagnóstico';
    toggle.setAttribute('aria-label', 'Mostrar ou ocultar o diagnóstico');
    toggle.textContent = 'i';
    toggle.setAttribute('style', 'position:fixed;left:8px;bottom:8px;z-index:201;width:26px;height:26px;padding:0;border-radius:50%;border:1px solid rgba(0,0,0,.25);background:rgba(0,0,0,.35);color:#fff;font:italic 600 14px/1 Georgia,serif;opacity:.5;cursor:pointer');
    function show_(on) {
      if (on) { if (!box.parentNode) document.body.appendChild(box); renderDiag_(); } else if (box.parentNode) box.parentNode.removeChild(box);
      toggle.style.opacity = on ? '.9' : '.5';
      try { localStorage.setItem(DIAG_KEY, on ? '0' : '1'); } catch (e) { /* segue */ }
    }
    toggle.onclick = function () { show_(!box.parentNode); };
    box.querySelector('.d-mode').onclick = function (e) {
      e.preventDefault();
      if (!writeMode_() && !confirm('Ligar o MODO GRAVAÇÃO? Lançamentos, edições e exclusões passam a valer na planilha escolhida. Use uma CÓPIA para testar.')) return;
      setWriteMode_(!writeMode_());
      location.reload();
    };
    box.querySelector('.d-sheet').onclick = function (e) {
      e.preventDefault();
      if (!confirm('Trocar de planilha? Você vai escolher ou criar outra na próxima tela (a atual não é apagada).')) return;
      try { localStorage.removeItem(SHEET_KEY); } catch (err) { /* segue */ }
      location.reload();
    };
    box.querySelector('.d-x').onclick = function (e) { e.preventDefault(); show_(false); };
    box.querySelector('.d-out').onclick = function (e) { e.preventDefault(); signOut(); };
    document.body.appendChild(toggle);
    show_(!hidden);
    setInterval(renderDiag_, 5000);
  }

  // ---------- login ----------
  function whenGsi_() {
    return new Promise(function (resolve, reject) {
      var tries = 0;
      (function check() {
        if (window.google && google.accounts && google.accounts.oauth2) return resolve();
        if (++tries > 100) return reject(new Error('Não consegui carregar o login do Google (accounts.google.com/gsi/client).'));
        setTimeout(check, 100);
      })();
    });
  }

  function initClient_() {
    if (tokenClient && tokenClientMode === scopeMode_()) return Promise.resolve();
    if (!cfg.clientId) return Promise.reject(new Error('Falta o ID do cliente OAuth (web/config.js).'));
    return whenGsi_().then(function () {
      tokenClientMode = scopeMode_();
      tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: cfg.clientId,
        scope: SCOPES[tokenClientMode],
        // sem isto o Google soma o que a pessoa já autorizou antes (ex: a permissão completa) e o
        // experimento da permissão mínima não provaria nada
        include_granted_scopes: false,
        callback: function (resp) {
          var p = pendingToken; pendingToken = null;
          if (resp.error) {
            event_('login recusado: ' + resp.error);
            if (p) p.reject(new Error(resp.error_description || resp.error));
            return;
          }
          token = { value: resp.access_token, expiresAt: Date.now() + Number(resp.expires_in || 3600) * 1000 - TOKEN_MARGIN_MS, grantedAt: Date.now(), mode: tokenClientMode };
          everSignedIn = true;
          saveToken_();
          event_('token novo (vale ' + Math.round(Number(resp.expires_in || 3600) / 60) + ' min) · permissões: ' + String(resp.scope || '?').replace(/https:\/\/www\.googleapis\.com\/auth\//g, ''));
          if (p) p.resolve();
        },
        error_callback: function (err) {
          var p = pendingToken; pendingToken = null;
          event_('login falhou: ' + (err && (err.type || err.message) || 'erro'));
          if (p) p.reject(new Error((err && (err.type || err.message)) || 'login falhou'));
        }
      });
    });
  }

  /** Pede o token ao app Android (ponte window.FinancasApp); ele mostra o consentimento do Google se precisar. */
  function requestNativeToken_() {
    return new Promise(function (resolve, reject) {
      var timer = setTimeout(function () {
        window.__cacauNativeToken = null;
        reject(new Error('o app não respondeu ao pedido de login'));
      }, NATIVE_ASK_TIMEOUT_MS);
      window.__cacauNativeToken = function (ok, value) {
        clearTimeout(timer);
        window.__cacauNativeToken = null;
        if (!ok) { event_('login pelo app recusado: ' + value); return reject(new Error(String(value || 'login recusado'))); }
        token = { value: String(value), expiresAt: Date.now() + NATIVE_TOKEN_MS - TOKEN_MARGIN_MS, grantedAt: Date.now(), mode: 'file' };
        everSignedIn = true;
        saveToken_();
        event_('token novo entregue pelo app Android (assumido por 30 min)');
        resolve();
      };
      window.FinancasApp.requestWebToken();
    });
  }

  function requestToken_(prompt) {
    if (nativeLogin_()) return requestNativeToken_();
    return initClient_().then(function () {
      return new Promise(function (resolve, reject) {
        pendingToken = { resolve: resolve, reject: reject };
        var opts = { prompt: prompt };
        var hint = savedEmail_();
        if (hint) opts.hint = hint;
        tokenClient.requestAccessToken(opts);
      });
    });
  }

  /** Guarda o token na aba (sessionStorage) e no navegador (localStorage): sobrevive ao F5 e, enquanto ainda
   *  valer (o Google dá ~1 h), a fechar e reabrir o app sem pedir login de novo. */
  function saveToken_() {
    var raw = JSON.stringify(token);
    try { sessionStorage.setItem(TOKEN_KEY, raw); } catch (e) { /* segue */ }
    try { localStorage.setItem(TOKEN_KEY, raw); } catch (e) { /* segue */ }
  }
  function forgetToken_() {
    try { sessionStorage.removeItem(TOKEN_KEY); } catch (e) { /* segue */ }
    try { localStorage.removeItem(TOKEN_KEY); } catch (e) { /* segue */ }
  }
  function readSavedToken_(store) {
    try { return JSON.parse(store.getItem(TOKEN_KEY) || 'null'); } catch (e) { return null; }
  }
  function restoreToken_() {
    var stores = [];
    try { stores.push(sessionStorage); } catch (e) { /* segue */ }
    try { stores.push(localStorage); } catch (e) { /* segue */ }
    for (var i = 0; i < stores.length; i++) {
      var saved = readSavedToken_(stores[i]);
      if (saved && saved.value && Date.now() < saved.expiresAt && (saved.mode || 'file') === scopeMode_()) { token = saved; everSignedIn = true; return true; }
    }
    return false;
  }

  /** E-mail de quem já entrou neste navegador (com a permissão atual), ou ''. Com ele o Google não pergunta a conta
   *  de novo (prompt vazio + dica) e o login vira só um clique. */
  function savedEmail_() {
    try {
      var h = JSON.parse(localStorage.getItem(HINT_KEY) || 'null');
      return h && h.email && h.mode === scopeMode_() ? h.email : '';
    } catch (e) { return ''; }
  }
  function rememberEmail_(email) {
    try { if (email) localStorage.setItem(HINT_KEY, JSON.stringify({ email: email, mode: scopeMode_() })); } catch (e) { /* segue */ }
  }

  function tokenValid_() { return token.value && Date.now() < token.expiresAt; }

  /** Garante um token. Renovação silenciosa primeiro; se não der, pede um clique na tela. */
  function ensureToken_() {
    if (tokenValid_()) return Promise.resolve();
    if (nativeLogin_() && everSignedIn) {
      event_('token venceu: pedindo outro ao app');
      return requestNativeToken_().catch(function (err) {
        event_('renovação pelo app falhou: ' + err.message);
        return askUser_('Sua sessão com o Google expirou. Entre de novo para continuar.');
      });
    }
    if (everSignedIn) {
      event_('token venceu: tentando renovar sem pedir nada');
      var timer;
      var silent = requestToken_('').then(function () { clearTimeout(timer); });
      var timeout = new Promise(function (_, reject) {
        timer = setTimeout(function () { reject(new Error('renovação silenciosa demorou mais de ' + SILENT_TIMEOUT_MS / 1000 + ' s')); }, SILENT_TIMEOUT_MS);
      });
      return Promise.race([silent, timeout]).catch(function (err) {
        event_('renovação silenciosa falhou: ' + err.message);
        return askUser_('Sua sessão com o Google expirou. Entre de novo para continuar.');
      });
    }
    return askUser_('');
  }

  function askUser_(message) {
    if (!userGate) {
      userGate = new Promise(function (resolve) {
        showGate_(message, function () { userGate = null; resolve(); });
      });
    }
    return userGate;
  }

  function fetchEmail_() {
    return fetch('https://www.googleapis.com/oauth2/v3/userinfo', { headers: { Authorization: 'Bearer ' + token.value } })
      .then(function (r) { return r.json(); })
      .then(function (info) { CacauGAS.setEmail(info.email || ''); rememberEmail_(info.email); return info.email || ''; })
      .catch(function () { return ''; });
  }

  function signOut() {
    if (nativeLogin_() && typeof window.FinancasApp.signOut === 'function') { window.FinancasApp.signOut(); return; }
    try { if (token.value && window.google) google.accounts.oauth2.revoke(token.value, function () {}); } catch (e) { /* segue */ }
    token = { value: '', expiresAt: 0, grantedAt: 0 };
    everSignedIn = false;
    forgetToken_();
    try { localStorage.removeItem(HINT_KEY); } catch (e) { /* segue */ }
    location.reload();
  }

  // ---------- tela de entrada ----------
  var NEW_SHEET = {
    properties: { title: 'Cacau', locale: 'pt_BR', timeZone: 'America/Sao_Paulo' },
    // a aba que o Google cria junto já nasce com um nome do app: assim não sobra "Página1" para apagar
    sheets: [{ properties: { title: 'Accounts' } }]
  };

  function loadScript_(src) {
    return new Promise(function (resolve, reject) {
      var el = document.createElement('script');
      el.src = src; el.async = true;
      el.onload = resolve;
      el.onerror = function () { reject(new Error('Não consegui carregar ' + src)); };
      document.head.appendChild(el);
    });
  }

  /** Seletor de planilhas do Google (Picker): é ele que dá ao app acesso a uma planilha existente no modo mínimo. */
  function pickSheet_() {
    if (!cfg.apiKey) return Promise.reject(new Error('Falta a chave de API do seletor (apiKey em web/config.js).'));
    return (window.gapi ? Promise.resolve() : loadScript_('https://apis.google.com/js/api.js')).then(function () {
      return new Promise(function (resolve) { gapi.load('picker', resolve); });
    }).then(function () {
      return new Promise(function (resolve) {
        var view = new google.picker.DocsView(google.picker.ViewId.SPREADSHEETS).setMode(google.picker.DocsViewMode.LIST);
        new google.picker.PickerBuilder()
          .addView(view)
          .setOAuthToken(token.value)
          .setDeveloperKey(cfg.apiKey)
          .setAppId(cfg.projectNumber || '')
          .setTitle('Escolha a planilha do Cacau')
          .setCallback(function (data) {
            if (data.action === google.picker.Action.PICKED) resolve(data.docs[0].id);
            else if (data.action === google.picker.Action.CANCEL) resolve('');
          })
          .build().setVisible(true);
      });
    });
  }

  function useSheet_(id) {
    try { localStorage.setItem(SHEET_KEY, id); } catch (e) { /* segue */ }
    CacauGAS.PropertiesService.getUserProperties().setProperty('SPREADSHEET_ID', id);
  }

  /** Planilha nova na conta da pessoa, já com abas, categorias e a conta "Carteira" (como o onboarding do Apps Script). */
  function createSheet_() {
    return api_('', { method: 'POST', body: NEW_SHEET }).then(function (r) {
      var id = r.json.spreadsheetId;
      event_('planilha criada: ' + id);
      useSheet_(id);
      setWriteMode_(true); // planilha nova e vazia: gravar é seguro (e é preciso para semear)
      meta = null;
      return runCall_('seedSpreadsheet_', [], { setup: true }).then(function () {
        event_('planilha semeada (abas, categorias, conta Carteira)');
        return id;
      });
    });
  }

  /**
   * Tela de entrada. Quem já tem planilha guardada só vê "Entrar"; quem ainda não tem escolhe entre criar uma
   * ou escolher uma existente. O que é técnico (permissão, código da planilha, modo gravação) fica em "Opções avançadas".
   */
  function showGate_(message, onDone) {
    var saved = '';
    try { saved = localStorage.getItem(SHEET_KEY) || ''; } catch (e) { /* segue */ }
    var gate = document.getElementById('cacau-gate');
    if (gate) gate.remove();
    gate = document.createElement('div');
    gate.id = 'cacau-gate';
    gate.setAttribute('style', 'position:fixed;inset:0;z-index:150;overflow:auto;background:var(--bg,#FFF4DC);color:var(--text,#3b2314);display:flex;align-items:center;justify-content:center;padding:16px');
    var btn = 'font-size:16px;padding:12px;border-radius:10px;cursor:pointer';
    var primary = btn + ';background:var(--accent,#8a4b25);color:#fff;border:0';
    var secondary = btn + ';background:transparent;color:inherit;border:1px solid rgba(0,0,0,.25)';
    var newPickHtml =
      '<button id="cacau-new" style="' + primary + '">Entrar e criar minha planilha</button>' +
      '<button id="cacau-pick" style="' + secondary + '">Entrar e escolher uma planilha que já tenho</button>';
    gate.innerHTML =
      '<div style="max-width:380px;width:100%;display:flex;flex-direction:column;gap:12px;font-family:system-ui,sans-serif">' +
      '<h1 style="margin:0;font-size:30px">Cacau</h1>' +
      '<p style="margin:0;opacity:.8">Suas finanças numa planilha só sua, no seu Google Drive.</p>' +
      (message ? '<p style="margin:0;color:#b00">' + message + '</p>' : '') +
      (saved ? '<button id="cacau-go" style="' + primary + '">Entrar</button>' : newPickHtml) +
      '<p id="cacau-gate-msg" style="margin:0;color:#b00;font-size:13px"></p>' +
      '</div>';
    document.body.appendChild(gate);
    var msg = gate.querySelector('#cacau-gate-msg');
    var buttons = gate.querySelectorAll('button');
    function busy_(on) { Array.prototype.forEach.call(buttons, function (b) { b.disabled = on; }); }
    function done_() { gate.remove(); onDone(); gateDone(); }
    function fail_(err) { busy_(false); msg.textContent = 'Não deu: ' + err.message; }
    function login_() {
      setScopeMode_('file');
      msg.textContent = 'Abrindo o login do Google…';
      busy_(true);
      return requestToken_(everSignedIn ? '' : 'select_account');
    }
    function enterWith_(id) {
      login_().then(function () { useSheet_(id); return fetchEmail_(); }).then(done_).catch(fail_);
    }

    var go = gate.querySelector('#cacau-go');
    if (go) go.onclick = function () { enterWith_(saved); };
    var nw = gate.querySelector('#cacau-new');
    if (nw) nw.onclick = function () {
      login_().then(fetchEmail_).then(function () {
        msg.textContent = 'Criando e preparando a planilha…';
        return createSheet_();
      }).then(done_).catch(fail_);
    };
    var pk = gate.querySelector('#cacau-pick');
    if (pk) pk.onclick = function () {
      login_().then(function () { msg.textContent = 'Abrindo o seletor de planilhas…'; return pickSheet_(); }).then(function (id) {
        if (!id) { busy_(false); msg.textContent = 'Nenhuma planilha escolhida.'; return null; }
        useSheet_(id);
        return fetchEmail_().then(done_);
      }).catch(fail_);
    };
  }

  // ---------- Sheets API ----------
  function api_(path, init, retried) {
    stats.apiCalls++;
    var options = { method: (init && init.method) || 'GET', headers: { Authorization: 'Bearer ' + token.value } };
    if (init && init.body !== undefined) {
      options.headers['Content-Type'] = 'application/json';
      options.body = JSON.stringify(init.body);
    }
    return fetch((SHEETS + path).replace(/\/$/, ''), options).then(function (resp) {
      if (resp.status === 401 && !retried) {
        token.expiresAt = 0;
        event_('API devolveu 401: renovando');
        return ensureToken_().then(function () { return api_(path, init, true); });
      }
      return resp.text().then(function (text) {
        if (!resp.ok) {
          var detail = text;
          try { detail = JSON.parse(text).error.message; } catch (e) { /* texto cru */ }
          throw new Error('Google Planilhas (' + resp.status + '): ' + detail);
        }
        return { json: JSON.parse(text), bytes: text.length };
      });
    });
  }

  function loadMeta_(sheetId, schemaNames) {
    if (meta && meta.id === sheetId && Date.now() - metaAt < META_TTL_MS) return Promise.resolve({ bytes: 0 });
    var fields = encodeURIComponent('spreadsheetUrl,properties.timeZone,sheets.properties(sheetId,title,gridProperties(rowCount,columnCount))');
    return api_(sheetId + '?fields=' + fields).then(function (r) {
      var tabs = (r.json.sheets || []).map(function (s) { return s.properties; }).filter(function (p) { return schemaNames.indexOf(p.title) !== -1; });
      var ids = {};
      tabs.forEach(function (p) { ids[p.title] = p.sheetId; });
      meta = { id: sheetId, url: r.json.spreadsheetUrl || '', tz: (r.json.properties || {}).timeZone || CacauGAS.scriptTimeZone, tabs: tabs, ids: ids };
      metaAt = Date.now();
      return { bytes: r.bytes };
    });
  }

  /** Colunas cujo valor numérico é uma data/hora (o Apps Script devolveria Date). */
  function dateColumns_(schema, textColumns) {
    var set = {};
    Object.keys(textColumns).forEach(function (sheet) { textColumns[sheet].forEach(function (c) { set[sheet + '.' + c] = true; }); });
    Object.keys(schema).forEach(function (sheet) {
      schema[sheet].forEach(function (c) { if (/(^|[a-z])At$/.test(c)) set[sheet + '.' + c] = true; });
    });
    return set;
  }

  function convertRows_(sheetName, values, tz, dateCols) {
    var header = values[0] || [];
    var isDate = header.map(function (h) { return !!dateCols[sheetName + '.' + h]; });
    return values.map(function (row, rowIndex) {
      if (rowIndex === 0) return row;
      return row.map(function (v, i) {
        return (isDate[i] && typeof v === 'number') ? CacauGAS.serialToDate(v, tz) : v;
      });
    });
  }

  function loadSnapshot_(sheetId, server, retried) {
    var schemaNames = Object.keys(server.__SCHEMA);
    var net = { bytes: 0, api: 0, ms: 0 };
    var t0 = Date.now();
    return loadMeta_(sheetId, schemaNames).then(function (m) {
      net.bytes += m.bytes;
      if (!meta.tabs.length) return { values: [] };
      var query = meta.tabs.map(function (p) { return 'ranges=' + encodeURIComponent("'" + p.title.replace(/'/g, "''") + "'"); }).join('&');
      return api_(sheetId + '/values:batchGet?' + query + '&majorDimension=ROWS&valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=SERIAL_NUMBER')
        .then(function (r) { net.bytes += r.bytes; return { values: r.json.valueRanges || [] }; });
    }).then(function (data) {
      net.ms = Date.now() - t0;
      var dateCols = dateColumns_(server.__SCHEMA, server.__TEXT_COLUMNS);
      var sheets = meta.tabs.map(function (p, i) {
        var values = (data.values[i] && data.values[i].values) || [];
        return {
          title: p.title,
          rowCount: p.gridProperties && p.gridProperties.rowCount,
          columnCount: p.gridProperties && p.gridProperties.columnCount,
          values: convertRows_(p.title, values, meta.tz, dateCols)
        };
      });
      return { snapshot: { id: sheetId, url: meta.url, tz: meta.tz, sheets: sheets }, net: net };
    }, function (err) {
      // aba apagada/criada por fora com a lista de abas ainda em cache: relê a lista e tenta uma vez
      if (!retried && /\(400\)/.test(err.message)) {
        event_('lista de abas desatualizada: relendo');
        meta = null;
        return loadSnapshot_(sheetId, server, true);
      }
      throw err;
    });
  }

  function currentSheetId_() {
    var id = CacauGAS.PropertiesService.getUserProperties().getProperty('SPREADSHEET_ID');
    if (!id) throw new Error('Nenhuma planilha escolhida.');
    return id;
  }

  /** Mesma regra de serialização do google.script.run: só dados puros voltam. */
  function clone_(value) {
    return value === undefined ? null : JSON.parse(JSON.stringify(value));
  }

  // ---------- gravação ----------
  /** Código numérico de cada aba que o passo usa; se faltar (aba recém-criada), relê a lista de abas. */
  function idsFor_(sheetId, names, server) {
    var known = function () { return names.every(function (n) { return meta && meta.ids && meta.ids[n] !== undefined; }); };
    if (known()) return Promise.resolve(meta.ids);
    meta = null;
    return loadMeta_(sheetId, Object.keys(server.__SCHEMA)).then(function () {
      if (!known()) throw new Error('Não achei a aba "' + names.filter(function (n) { return !(meta.ids[n] !== undefined); })[0] + '" na planilha.');
      return meta.ids;
    });
  }

  /** Confere que as linhas que o servidor leu (coluna ID) ainda são as mesmas na planilha. */
  function verifyUnchanged_(sheetId, ops, server) {
    var names = CacauWriter.sheetsToVerify(ops).filter(function (n) { return meta && meta.ids && meta.ids[n] !== undefined; });
    if (!names.length) return Promise.resolve();
    var cols = names.map(function (n) { return (server.__SCHEMA[n] || []).indexOf('ID') + 1 || 1; });
    var query = names.map(function (n, i) {
      var letter = CacauWriter.colLetter(cols[i]);
      var last = CacauGAS.originalColumn(n, cols[i]).length + 1;
      return 'ranges=' + encodeURIComponent("'" + n.replace(/'/g, "''") + "'!" + letter + '2:' + letter + Math.max(last, 2));
    }).join('&');
    return api_(sheetId + '/values:batchGet?' + query + '&majorDimension=COLUMNS&valueRenderOption=UNFORMATTED_VALUE').then(function (r) {
      (r.json.valueRanges || []).forEach(function (vr, i) {
        var before = CacauGAS.originalColumn(names[i], cols[i]);
        var now = (vr.values && vr.values[0]) || [];
        for (var k = 0; k < before.length; k++) {
          var was = before[k] === undefined ? '' : before[k];
          var is = now[k] === undefined ? '' : now[k];
          if (String(was) !== String(is)) { event_('conflito na aba ' + names[i] + ' (linha ' + (k + 2) + ')'); throw new Error(CONFLICT_MESSAGE); }
        }
      });
    });
  }

  function runSteps_(sheetId, steps, index, server) {
    if (index >= steps.length) return Promise.resolve();
    var step = steps[index];
    var names = CacauWriter.sheetsNeedingIds(step);
    var ready = names.length ? idsFor_(sheetId, names, server) : Promise.resolve(meta ? meta.ids : {});
    return ready.then(function (ids) {
      var req = CacauWriter.request(step, ids, meta ? meta.tz : CacauGAS.scriptTimeZone, CacauGAS.dateToSerial);
      return api_(sheetId + req.path, { method: req.method, body: req.body });
    }).then(function () {
      if (step.kind === 'structure' && CacauWriter.createsSheets(step.ops)) meta = null; // sheetIds mudaram
      return runSteps_(sheetId, steps, index + 1, server);
    });
  }

  function flush_(sheetId, server, record) {
    var ops = CacauGAS.pendingOps();
    if (!ops.length) return Promise.resolve();
    record.writes = ops.length;
    var steps = CacauWriter.plan(ops);
    return verifyUnchanged_(sheetId, ops, server).then(function () {
      return runSteps_(sheetId, steps, 0, server);
    });
  }

  // ---------- chamada ----------
  function runCall_(fnName, args, opts) {
    opts = opts || {};
    var started = Date.now();
    var record = { at: started, fn: fnName, ms: 0, netMs: 0, api: 0, bytes: 0 };
    var apiBefore = stats.apiCalls;
    var writing = opts.setup || writeMode_();
    // opts.setup: preparação feita pela própria tela de entrada, antes de liberar o app
    return (opts.setup ? Promise.resolve() : ready).then(function () { return ensureToken_(); }).then(function () {
      var server = window.CacauServer;
      if (typeof server[fnName] !== 'function') throw new Error('Função "' + fnName + '" não existe no pacote web.');
      var id = currentSheetId_();
      return loadSnapshot_(id, server).then(function (loaded) {
        record.netMs = loaded.net.ms;
        record.bytes = loaded.net.bytes;
        var mode = writing ? 'write' : (/^(get|list)[A-Z]/.test(fnName) ? 'discard' : 'readonly');
        CacauGAS.setSnapshot(loaded.snapshot, { mode: mode });
        var result = clone_(server[fnName].apply(null, args));
        if (mode !== 'write') return result;
        return flush_(id, server, record).then(function () { return result; });
      });
    }).then(function (result) {
      record.ms = Date.now() - started;
      record.api = stats.apiCalls - apiBefore;
      stats.calls.unshift(record); if (stats.calls.length > 20) stats.calls.length = 20;
      renderDiag_();
      return result;
    }, function (err) {
      record.ms = Date.now() - started;
      record.api = stats.apiCalls - apiBefore;
      record.error = (err && err.message) || String(err);
      stats.calls.unshift(record); if (stats.calls.length > 20) stats.calls.length = 20;
      renderDiag_();
      throw err;
    });
  }

  /** Uma chamada por vez (o papel do LockService no Apps Script). */
  function call(fnName, args) {
    var run = function () { return runCall_(fnName, args); };
    var next = queue.then(run, run);
    queue = next.then(function () {}, function () {});
    return next;
  }

  function start() {
    mountDiag_();
    var sheetId = '';
    try { sheetId = localStorage.getItem(SHEET_KEY) || ''; } catch (e) { /* segue */ }
    if (sheetId && restoreToken_()) {
      CacauGAS.PropertiesService.getUserProperties().setProperty('SPREADSHEET_ID', sheetId);
      event_('sessão retomada após recarregar (token ainda valia)');
      fetchEmail_().then(function () { gateDone(); });
      return;
    }
    if (savedEmail_()) everSignedIn = true; // já entrou aqui antes: login sem escolher conta
    if (sheetId && nativeLogin_() && everSignedIn) {
      // App Android: o login é do próprio app e a renovação não pede nada à pessoa -- tenta direto, sem a tela
      // de entrada; se o app recusar, cai nela como antes.
      event_('abrindo: pedindo o token ao app, sem tela de entrada');
      requestNativeToken_().then(function () {
        CacauGAS.PropertiesService.getUserProperties().setProperty('SPREADSHEET_ID', sheetId);
        return fetchEmail_();
      }).then(function () { gateDone(); }).catch(function (err) {
        event_('entrada automática pelo app falhou: ' + err.message);
        askUser_('');
      });
      return;
    }
    askUser_('');
  }
  document.addEventListener('DOMContentLoaded', start);

  /** Esquece a planilha guardada e volta à tela de entrada (a planilha antiga não é apagada). */
  function switchSheet() {
    try { localStorage.removeItem(SHEET_KEY); } catch (e) { /* segue */ }
    location.reload();
  }

  return { call: call, signOut: signOut, stats: stats, createSheet: createSheet_, switchSheet: switchSheet };
})();
