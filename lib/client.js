/**
 * 深鲸文档批量处理 —— 客户端半边（统一面板）
 *
 * 一个入口管三种文档：Word / Excel / PDF（各自宿主 API 不同，面板里切页签）。
 * 设计取舍（2026-10-28 用户反馈）：不要为每种格式各放一个侧栏图标 ——
 * 用户只需要知道「我要批量处理文档」，不需要先分辨格式再找入口。
 *
 * 槽位：sidebar.panellist（一个图标）+ main（keyed 面板）
 * 数据：/dsh-docx-batch/api、/dsh-xlsx-batch/api、/dsh-pdf-kit/api（同源）
 */
window.__ModuleLoader__.load({
  id: '@deepwhale-cn/dsh-docx-batch',
  factory: (require) => {
    var module = { exports: {} };
    var exports = module.exports;
    const react = require('react');
    const h = react.createElement;

    const SLOT_ICON = 'sidebar.panellist';
    const SLOT_MAIN = 'main';
    const PANEL_ID = 'office-batch';

    const FORMATS = [
      { key: 'word', label: 'Word', api: '/dsh-docx-batch/api', tips: '需要 python3 + python-docx' },
      { key: 'excel', label: 'Excel', api: '/dsh-xlsx-batch/api', tips: '需要 python3 + openpyxl' },
      { key: 'pdf', label: 'PDF', api: '/dsh-pdf-kit/api', tips: '需要 python3 + pypdf' },
    ];

    async function call(api, path, options) {
      const res = await fetch(`${api}${path}`, {
        headers: { 'content-type': 'application/json' },
        ...options,
        body: options && options.body !== undefined ? JSON.stringify(options.body) : undefined,
      });
      const text = await res.text();
      let data = null;
      try { data = text ? JSON.parse(text) : null; } catch { data = { ok: false, error: text }; }
      if (!res.ok && data && !data.error) data.error = `HTTP ${res.status}`;
      return data;
    }

    const IcoDoc = () => h('svg', { width: 18, height: 18, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 1.8, strokeLinecap: 'round', strokeLinejoin: 'round' }, [
      h('path', { key: 'a', d: 'M4 4h10l6 6v10a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1z' }),
      h('path', { key: 'b', d: 'M14 4v6h6' }),
      h('path', { key: 'c', d: 'M8 13h8M8 17h5' }),
    ]);

    const CSS = `
      .dwhb-root { display:flex; flex-direction:column; gap:14px; padding:18px 22px 30px; height:100%; overflow:auto; }
      .dwhb-head { display:flex; align-items:center; gap:10px; flex-wrap:wrap; }
      .dwhb-title { font-size:16px; font-weight:600; }
      .dwhb-sub { color: var(--dsw-alias-text-secondary, #8b98a8); font-size:12px; }
      .dwhb-brand { margin-left:auto; display:inline-flex; align-items:center; gap:6px; font-size:11px; color:var(--dsw-alias-text-secondary,#8b98a8); border:1px solid var(--dsw-alias-border-l1, rgba(255,255,255,.12)); border-radius:999px; padding:2px 10px; }
      .dwhb-brand b { color: var(--dsw-alias-brand-primary, #7fe3f0); font-weight:600; }
      .dwhb-tabs { display:flex; gap:6px; }
      .dwhb-tab { padding:6px 16px; border-radius:8px; font-size:13px; cursor:pointer; border:1px solid transparent; background:transparent; color:inherit; }
      .dwhb-tab[data-on='true'] { background: var(--dsw-alias-bg-layer-1, rgba(255,255,255,.06)); border-color: var(--dsw-alias-border-l1, rgba(255,255,255,.14)); font-weight:600; }
      .dwhb-card { border:1px solid var(--dsw-alias-border-l1, rgba(255,255,255,.08)); border-radius:12px; background: var(--dsw-alias-bg-layer-1, rgba(255,255,255,.03)); padding:14px; display:flex; flex-direction:column; gap:10px; }
      .dwhb-row { display:flex; gap:8px; align-items:center; }
      .dwhb-label { font-size:12px; color: var(--dsw-alias-text-secondary, #8b98a8); width:70px; flex:none; }
      .dwhb-input, .dwhb-select, .dwhb-area { flex:1; min-width:0; padding:7px 9px; border-radius:8px; border:1px solid var(--dsw-alias-border-l1, rgba(255,255,255,.12)); background: var(--dsw-alias-bg-base, transparent); color:inherit; font-size:13px; font-family:inherit; }
      .dwhb-area { min-height:70px; resize:vertical; font-family: ui-monospace, Menlo, monospace; }
      .dwhb-btn { padding:7px 16px; border-radius:8px; border:1px solid var(--dsw-alias-border-l1, rgba(255,255,255,.14)); background: var(--dsw-alias-brand-primary, #2f6f9f); color:#fff; font-size:13px; cursor:pointer; }
      .dwhb-btn[disabled] { opacity:.55; cursor:default; }
      .dwhb-out { white-space:pre-wrap; word-break:break-word; font-family: ui-monospace, Menlo, monospace; font-size:12px; line-height:1.6; max-height:280px; overflow:auto; background: var(--dsw-alias-bg-layer-2, rgba(0,0,0,.18)); border-radius:8px; padding:10px; }
      .dwhb-err { color: var(--dsw-alias-state-error-primary, #e0645a); }
      .dwhb-ok { color: var(--dsw-alias-state-success-primary, #46b47a); }
      .dwhb-tip { font-size:12px; color: var(--dsw-alias-text-secondary, #8b98a8); }
    `;

    function 面板() {
      const [fmt, 设fmt] = react.useState('word');
      const [ops, 设ops] = react.useState({});
      const [action, 设action] = react.useState('');
      const [input, 设input] = react.useState('');
      const [out, 设out] = react.useState('');
      const [pairs, 设pairs] = react.useState('');
      const [text, 设text] = react.useState('');
      const [mode, 设mode] = react.useState('page');
      const [n, 设n] = react.useState('');
      const [跑, 设跑] = react.useState(false);
      const [结果, 设结果] = react.useState(null);

      const 当前格式 = FORMATS.find((f) => f.key === fmt) || FORMATS[0];

      react.useEffect(() => {
        let 活 = true;
        void (async () => {
          const r = await call(当前格式.api, '/ops');
          if (!活 || !r) return;
          const list = Array.isArray(r.ops) ? r.ops : [];
          设ops((prev) => ({ ...prev, [fmt]: list }));
          if (list.length) 设action(list[0].key);
        })();
        return () => { 活 = false; };
      }, [fmt]);

      const 列表 = ops[fmt] || [];
      const 当前操作 = 列表.find((o) => o.key === action);
      const 是替换 = fmt === 'word' && action === 'replace';
      const 是拆分 = fmt === 'pdf' && action === 'split';
      const 是水印 = fmt === 'pdf' && action === 'watermark';

      const 执行 = async () => {
        设跑(true); 设结果(null);
        try {
          const r = await call(当前格式.api, '/run', { method: 'POST', body: { action, input, out, pairs, text, mode, n, dpi: n } });
          设结果(r);
        } catch (e) {
          设结果({ ok: false, error: e instanceof Error ? e.message : String(e) });
        } finally { 设跑(false); }
      };

      return h('div', { className: 'dwhb-root' }, [
        h('style', { key: 'css' }, CSS),
        h('div', { key: 'head', className: 'dwhb-head' }, [
          h(IcoDoc, { key: 'i' }),
          h('div', { key: 't' }, [
            h('div', { key: 'a', className: 'dwhb-title' }, '深鲸文档批量处理'),
            h('div', { key: 'b', className: 'dwhb-sub' }, 'Word / Excel / PDF 一次处理整目录 —— 替换、提取、合并、拆分、导出'),
          ]),
          h('span', { key: 'brand', className: 'dwhb-brand' }, ['powered by ', h('b', { key: 'b' }, 'dwh'), ' 深鲸']),
        ]),
        h('div', { key: 'tabs', className: 'dwhb-tabs' }, FORMATS.map((f) => h('button', {
          key: f.key, className: 'dwhb-tab', 'data-on': f.key === fmt ? 'true' : 'false',
          onClick: () => { 设fmt(f.key); 设结果(null); },
        }, f.label))),

        h('div', { key: 'form', className: 'dwhb-card' }, [
          h('div', { key: 'op', className: 'dwhb-row' }, [
            h('span', { key: 'l', className: 'dwhb-label' }, '操作'),
            h('select', { key: 's', className: 'dwhb-select', value: action, onChange: (e) => 设action(e.target.value) },
              列表.map((o) => h('option', { key: o.key, value: o.key }, o.label))),
          ]),
          h('div', { key: 'in', className: 'dwhb-row' }, [
            h('span', { key: 'l', className: 'dwhb-label' }, '文件/目录'),
            h('input', { key: 'i', className: 'dwhb-input', placeholder: '把文件或文件夹路径粘进来（可从访达/资源管理器复制）', value: input, onChange: (e) => 设input(e.target.value) }),
          ]),
          h('div', { key: 'out', className: 'dwhb-row' }, [
            h('span', { key: 'l', className: 'dwhb-label' }, '输出目录'),
            h('input', { key: 'i', className: 'dwhb-input', placeholder: '留空 = <输入目录>/_输出', value: out, onChange: (e) => 设out(e.target.value) }),
          ]),
          是替换 ? h('div', { key: 'pairs', className: 'dwhb-row', style: { alignItems: 'flex-start' } }, [
            h('span', { key: 'l', className: 'dwhb-label', style: { paddingTop: 7 } }, '替换规则'),
            h('textarea', { key: 't', className: 'dwhb-area', placeholder: '深鲸科技 → 深海鲸\n旧公司名 → 新公司名', value: pairs, onChange: (e) => 设pairs(e.target.value) }),
          ]) : null,
          是拆分 ? h('div', { key: 'split', className: 'dwhb-row' }, [
            h('span', { key: 'l', className: 'dwhb-label' }, '拆分方式'),
            h('select', { key: 's', className: 'dwhb-select', value: mode, onChange: (e) => 设mode(e.target.value) }, [
              h('option', { key: 'p', value: 'page' }, '按页拆（每页一个）'),
              h('option', { key: 'e', value: 'every' }, '每 N 页一组'),
            ]),
            h('input', { key: 'n', className: 'dwhb-input', style: { maxWidth: 110 }, placeholder: 'N', value: n, onChange: (e) => 设n(e.target.value) }),
          ]) : null,
          是水印 ? h('div', { key: 'wm', className: 'dwhb-row' }, [
            h('span', { key: 'l', className: 'dwhb-label' }, '水印文字'),
            h('input', { key: 't', className: 'dwhb-input', placeholder: '例如 机密 / 内部资料', value: text, onChange: (e) => 设text(e.target.value) }),
          ]) : null,
          当前操作 && 当前操作.hint ? h('div', { key: 'hint', className: 'dwhb-tip' }, 当前操作.hint) : null,
          h('div', { key: 'act', className: 'dwhb-row', style: { justifyContent: 'flex-end' } }, [
            h('button', { key: 'b', className: 'dwhb-btn', disabled: 跑 || !action, onClick: 执行 }, 跑 ? '正在处理…' : '开始处理'),
          ]),
        ]),

        结果 ? h('div', { key: 'res', className: 'dwhb-card' }, [
          h('div', { key: 's', className: 结果.ok ? 'dwhb-ok' : 'dwhb-err' }, 结果.ok ? '完成' : `失败${结果.error ? '：' + 结果.error : ''}`),
          结果.output ? h('div', { key: 'o', className: 'dwhb-out' }, 结果.output) : null,
        ]) : null,

        h('div', { key: 'tip', className: 'dwhb-tip' },
          `本机处理、不上传文件。${当前格式.tips}。命令等价于 CLI：dsh-${fmt === 'word' ? 'docx' : fmt === 'excel' ? 'xlsx' : 'pdf'}-batch <操作> --input <路径>`),
      ]);
    }

    const inject = ['slots'];
    function apply(ctx) {
      const 试 = (名, fn) => {
        try { fn(); } catch (e) { try { console.warn('[office-batch] 插槽注册失败（已跳过）：' + 名, e && e.message ? e.message : e); } catch {} }
      };
      试('侧栏图标', () => ctx.slots.inject(SLOT_ICON, function* () {
        yield ctx.slots.register({ name: SLOT_ICON, id: PANEL_ID, order: 24, label: '深鲸文档批量' }, IcoDoc);
      }));
      试('主面板', () => ctx.slots.inject(SLOT_MAIN, function* () {
        yield ctx.slots.register({ name: SLOT_MAIN, key: PANEL_ID }, 面板);
      }));
    }

    exports.inject = inject;
    exports.apply = apply;
    return module.exports;
  },
});
