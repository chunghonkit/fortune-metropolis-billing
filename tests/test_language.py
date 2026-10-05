"""Dashboard language catalog and the remembered English / Traditional Chinese choice."""

import json
import re
import subprocess
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
HTML = (ROOT / 'static' / 'index.html').read_text(encoding='utf-8')


def _used_keys():
    keys = set(re.findall(r'data-i18n="([^"]+)"', HTML))
    keys.update(re.findall(r'data-i18n-placeholder="([^"]+)"', HTML))
    keys.update(re.findall(r'data-i18n-aria="([^"]+)"', HTML))
    keys.update(re.findall(r"\bt\(\s*'([^']+)'", HTML))
    return keys


def test_every_ui_string_has_both_languages_and_choice_persists():
    report = subprocess.check_output(
        [
            'node',
            '-e',
            r'''
const i18n = require("./static/i18n.js");
const en = i18n.UI.en;
const zh = i18n.UI.zh;
const enKeys = Object.keys(en).sort();
const zhKeys = Object.keys(zh).sort();
if (JSON.stringify(enKeys) !== JSON.stringify(zhKeys)) {
  console.error(JSON.stringify({ enOnly: enKeys.filter(k => !(k in zh)), zhOnly: zhKeys.filter(k => !(k in en)) }));
  process.exit(1);
}
const cjk = /[\u3400-\u9fff]/;
for (const key of enKeys) {
  if (!String(en[key]).trim() || !String(zh[key]).trim()) {
    console.error("empty " + key);
    process.exit(2);
  }
  if (en[key] !== zh[key] && !cjk.test(zh[key])) {
    console.error("Chinese missing for " + key);
    process.exit(3);
  }
  if (en[key] === zh[key] && !/^(?:YYYY-MM|Metropolis|DC 100%|SW 100%|C 100%|SA 100%)$/.test(en[key])) {
    console.error("same text is not a code: " + key + "=" + en[key]);
    process.exit(4);
  }
}
for (const rule of i18n.SERVER) {
  if (!rule.zh || !cjk.test(rule.zh) || !(rule.en || rule.pattern || rule.prefix || rule.suffix)) {
    console.error("server rule " + JSON.stringify(rule));
    process.exit(5);
  }
}
const memory = new Map();
const storage = {
  getItem: (key) => memory.has(key) ? memory.get(key) : null,
  setItem: (key, value) => memory.set(key, value),
};
if (i18n.savedLanguage(storage) !== "en") process.exit(6);
if (i18n.rememberLanguage(storage, "zh") !== "zh") process.exit(7);
if (storage.getItem("dashboard-lang") !== "zh") process.exit(8);
if (i18n.savedLanguage(storage) !== "zh") process.exit(9);
if (i18n.rememberLanguage(storage, "en") !== "en") process.exit(10);
if (i18n.savedLanguage(storage) !== "en") process.exit(11);
i18n.setActiveLanguage("zh");
const accounts = "55861-52267-1, 00776-78552-1";
const missing = i18n.translateServer("Missing 2 accounts: " + accounts);
if (!missing.includes(accounts) || !cjk.test(missing)) {
  console.error(missing);
  process.exit(12);
}
if (i18n.translateServer("Please set billing month first") !== "請先設定帳單月份") process.exit(13);
if (!i18n.translateServer("82805-94744-7.pdf: Not a PDF").startsWith("82805-94744-7.pdf")) process.exit(14);
i18n.setActiveLanguage("en");
if (i18n.translateServer("Please set billing month first") !== "Please set billing month first") process.exit(15);
if (i18n.t("electricityBills") !== "Electricity bills") process.exit(16);
i18n.setActiveLanguage("zh");
if (i18n.t("meters") !== "電錶" || i18n.t("costCentres") !== "成本中心" || i18n.t("kwh") !== "用電量") process.exit(17);
console.log(JSON.stringify(enKeys));
''',
        ],
        cwd=ROOT,
        text=True,
    )
    catalog = set(json.loads(report.strip().splitlines()[-1]))
    used = _used_keys()
    used.update(key for key in re.findall(r"'([A-Za-z0-9]+)'", HTML) if key in catalog)
    missing = sorted(used - catalog)
    assert missing == []
    unused = sorted(catalog - used)
    assert unused == []
    assert 'savedLanguage(localStorage)' in HTML
    assert "rememberLanguage(localStorage, lang)" in HTML
    assert 'data-lang="en"' in HTML
    assert 'data-lang="zh"' in HTML
    assert '繁體中文' in HTML
    assert 'id="i18n.js"' not in HTML
    assert 'src="/i18n.js"' in HTML
