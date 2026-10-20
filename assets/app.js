/* 铁矿石发运·到港 季节性图谱
 * 数据来自 data/data.js（内联，file:// 可直接打开），缺失时回退 fetch data/*.json
 * URL 参数：?ds=global|aubra|arrival 指定数据集；?shot=1 截图模式（全量渲染，不懒加载）
 */
(function () {
  'use strict';

  var S = {
    data: null,
    dsId: null,
    items: [],
    io: null,
    shot: false,
    hidden: {},        // 年份 -> true 表示隐藏
    years: [],
    opt: { connect: true, zero: false, sync: false },
    expanded: false
  };

  // 各数据集「区块内一行几张图」（默认 5 张，见 style.css .grid.rgrid）
  var GRID_COLS = { psi_plan: 3, daiguan: 3, chugang: 3, export_variety: 3, export_country: 3, mill_order: 3, biaowai: 3, variety_demand: 3, variety_inventory: 3, lc_inv: 4 };

  // 嵌入式数据集：本站（碳酸锂）暂无
  var EMBED_DATASETS = [];

  // 外链数据集：本站（碳酸锂）暂无
  var LINK_DATASETS = [];

  // 横坐标按「1月…12月」显示（每月 1 号一个刻度）的数据集
  var MONTH_AXIS = { daiguan: 1, chugang: 1, hanguan: 1, juanluo_luowen: 1, juanluo_rejuan: 1, mill_order: 1, biaowai: 1, variety_demand: 1, variety_inventory: 1, lc_inv: 1 };

  // 纯月份数字轴（'1'..'12'，月度数据）：12 个月标签全显示、格式化为「M月」
  var MONTH_NUM_AXIS = { psi_plan: 1, export_variety: 1, export_country: 1 };

  var DASH = {
    solid: 'solid', dash: 'dashed', sysDash: 'dashed',
    dot: 'dotted', sysDot: 'dotted',
    dashDot: [6, 3, 2, 3], sysDashDot: [6, 3, 2, 3]
  };

  function $(id) { return document.getElementById(id); }

  function dashOf(v) { return DASH[v] || 'solid'; }

  /* ---------------- 数据加载 ---------------- */
  function loadData(cb) {
    if (window.CHART_DATA && window.CHART_DATA.datasets) {
      S.data = window.CHART_DATA;
      cb();
      return;
    }
    fetch('data/meta.json').then(function (r) { return r.json(); }).then(function (meta) {
      var jobs = meta.datasets.map(function (d) {
        return fetch('data/' + d.id + '.json').then(function (r) { return r.json(); });
      });
      return Promise.all(jobs).then(function (arr) {
        S.data = { updated: meta.updated, datasets: arr };
        cb();
      });
    }).catch(function (e) {
      $('loading').textContent = '数据加载失败：' + e.message + '（请先运行 scripts/extract.py）';
    });
  }

  /* ---------------- 图表配置 ---------------- */
  // 排序柱状图（环比增量 / 累计同比增量）：正红负绿，柱子顶端标数值
  function buildBarOption(chart, ds) {
    var b = chart.bar || {};
    var cats = b.categories || [];
    var vals = b.values || [];
    var pcts = b.pcts || [];
    var unit = b.unit || '万吨';
    var upColor = '#c00000', downColor = '#2e7d32';
    return {
      animation: false,
      title: {
        text: chart.title, left: 'center', top: 6,
        textStyle: { fontSize: 13, fontWeight: 600, color: '#1f2937' }
      },
      tooltip: {
        trigger: 'axis', confine: true, textStyle: { fontSize: 11 },
        axisPointer: { type: 'shadow' },
        formatter: function (ps) {
          var p = ps[0]; if (!p) return '';
          var i = p.dataIndex;
          var v = vals[i];
          var s = p.name + '<br/>' + (v > 0 ? '+' : '') + v + ' ' + unit;
          if (pcts[i] !== null && pcts[i] !== undefined) {
            s += '（' + (pcts[i] > 0 ? '+' : '') + pcts[i] + '%）';
          }
          return s;
        }
      },
      grid: { left: 48, right: 18, top: 44, bottom: 78 },
      xAxis: {
        type: 'category', data: cats,
        axisLine: { lineStyle: { color: '#d5dbe6' } },
        axisTick: { show: false },
        axisLabel: {
          fontSize: 8, color: '#6b7280', interval: 0,
          rotate: cats.length > 10 ? 40 : 0
        }
      },
      yAxis: {
        type: 'value',
        splitLine: { lineStyle: { color: '#eef1f6' } },
        axisLine: { show: false }, axisTick: { show: false },
        axisLabel: { fontSize: 9, color: '#94a3b8' }
      },
      series: [{
        type: 'bar', barMaxWidth: 22,
        data: vals.map(function (v) {
          return {
            value: v,
            itemStyle: { color: v >= 0 ? upColor : downColor },
            label: {
              show: true, position: v >= 0 ? 'top' : 'bottom',
              fontSize: 8, color: '#6b7280',
              formatter: function (p) { return (p.value > 0 ? '+' : '') + p.value; }
            }
          };
        })
      }]
    };
  }

  function buildOption(chart, ds, yRange) {
    if (chart.type === 'bar') return buildBarOption(chart, ds);
    var axis = ds.axes[chart.axis];
    var series = chart.series.filter(function (s) { return !S.hidden[s.name]; });
    var maxYear = series.length ? series[series.length - 1].name : '';

    var opt = {
      animation: false,
      title: {
        text: chart.title, left: 'center', top: 6,
        textStyle: { fontSize: 13, fontWeight: 600, color: '#1f2937' }
      },
      tooltip: {
        trigger: 'axis', confine: true,
        textStyle: { fontSize: 11 },
        axisPointer: { type: 'line', lineStyle: { color: '#cbd5e1', width: 1 } }
      },
      legend: {
        top: 27, itemGap: 6, itemWidth: 16, itemHeight: 8,
        textStyle: { fontSize: 9, color: '#6b7280' },
        // legend 图标默认不继承 series 的线型，必须逐项显式带 lineStyle，
        // 否则 2022/2023 的虚线系列在图例里仍显示成实线。
        data: series.map(function (s) {
          return {
            name: s.name,
            itemStyle: { color: s.color },
            lineStyle: {
              color: s.color,
              type: dashOf(s.dash),
              width: s.width || 1.5
            }
          };
        })
      },
      grid: { left: 54, right: 20, top: 54, bottom: 32 },
      xAxis: {
        type: 'category', data: axis, boundaryGap: false,
        axisLine: { lineStyle: { color: '#d5dbe6' } },
        // 月度轴：在每月 1 号（日频数据则取每月首个交易日）位置显示刻度线
        axisTick: (MONTH_AXIS[S.dsId] || S.monthFirst)
          ? {
              show: true, length: 5, lineStyle: { color: '#cbd5e1' },
              interval: S.monthFirst
                ? function (i, v) { return S.monthFirst.indexOf(v) >= 0; }
                : function (i, v) { return /-01$/.test(v); }
            }
          : { show: false },
        axisLabel: {
          // 月度轴标签多，字号调小到 8
          fontSize: (MONTH_AXIS[S.dsId] || MONTH_NUM_AXIS[S.dsId]) ? 8 : 9, color: '#94a3b8',
          // 月度轴：横坐标显示「1月…12月」
          formatter: function (v) {
            if (S.dateAxis) {
              var dm = /^(\d{4})-(\d{2})-/.exec(v);
              if (dm) return dm[2] === '01' ? dm[1] : dm[2] + '月';
              return v;
            }
            if (MONTH_AXIS[S.dsId]) {
              var mm = /^(\d{2})-/.exec(v);
              if (mm) return parseInt(mm[1], 10) + '月';
              return v;
            }
            if (MONTH_NUM_AXIS[S.dsId]) {
              var mn = parseInt(v, 10);
              return isNaN(mn) ? v : mn + '月';
            }
            return v;
          },
          // 同时兼容五种轴：
          //  · 日历轴（MM-DD，周度数据）→ 每月 1 号（螺纹/热卷/带钢/出港/焊管）
          //  · 纯月份数字轴（'1'..'12'，月度数据）→ 12 个月全显示
          //  · 其他日历轴 → 季度首月
          //  · 连续日期轴（YYYY-MM-DD，日频序列）→ dateTicks 季度刻度
          //  · 日频季节性轴（MM-DD + monthFirst，如钢厂日接单）→ 每月首个交易日
          interval: function (i, v) {
            if (S.dateAxis) return S.dateTicks ? (S.dateTicks.indexOf(i) >= 0) : (i === 0);
            if (S.monthFirst) return S.monthFirst.indexOf(v) >= 0;
            if (MONTH_NUM_AXIS[S.dsId]) return true;
            if (MONTH_AXIS[S.dsId]) return /-01$/.test(v);
            var m = /^\s*(\d{1,2})\s*月?\s*$/.exec(v);
            if (m) { var n = parseInt(m[1], 10); return n === 1 || n === 4 || n === 7 || n === 10; }
            return i === 0 || /^(01|04|07|10)-01$/.test(v);
          }
        }
      },
      yAxis: {
        type: 'value',
        scale: !S.opt.zero,
        min: yRange ? yRange.min : null,
        max: yRange ? yRange.max : null,
        splitLine: { lineStyle: { color: '#eef1f6' } },
        axisLine: { show: false },
        axisTick: { show: false },
        axisLabel: { fontSize: 9, color: '#94a3b8' }
      },
      series: (function () {
        var ech = series.map(function (s) {
          return {
            name: s.name,
            type: 'line',
            data: s.data,
            connectNulls: S.opt.connect,
            smooth: !!s.smooth,
            showSymbol: !!s.marker && s.marker !== 'none',
            symbol: 'circle',
            symbolSize: 3.5,
            lineStyle: { width: s.width || 1.5, type: dashOf(s.dash), color: s.color },
            itemStyle: { color: s.color },
            emphasis: { focus: 'series' },
            z: s.name === maxYear ? 6 : 2
          };
        });
        // 参考线（如日产、盈亏平衡）：画一条水平 markLine
        if (chart.refLine !== undefined && chart.refLine !== null && ech.length) {
          ech[0].markLine = {
            silent: true, symbol: 'none',
            lineStyle: { color: '#9ca3af', type: 'dashed', width: 1 },
            label: {
              show: true, position: 'insideEndTop', fontSize: 8, color: '#9ca3af',
              formatter: chart.refLabel || ('参考 ' + chart.refLine)
            },
            data: [{ yAxis: chart.refLine }]
          };
        }
        return ech;
      })()
    };

    return opt;
  }

  /* ---------------- 汇总表 ---------------- */
  function fmtNum(v, suffix) {
    // 绝对值 >=100 整数，否则保留 1 位小数；自动千分位
    var abs = Math.abs(v);
    var fixed = abs >= 100 ? 0 : (abs >= 10 ? 1 : 2);
    var s = v.toFixed(fixed);
    var parts = s.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return parts.join('.') + (suffix || '');
  }

  /* ---------------- 图下指标条 + 下载按钮（表外 / 分品种表需，仿汾渭卡片） ---------------- */
  // 图下小表专用数值格式：固定 2 位小数 + 千分位（与 Excel 口径一致）
  function fmtCardNum(v) {
    var s = (Math.round(v * 100) / 100).toFixed(2);
    var p = s.split('.');
    p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return p.join('.');
  }

  // 图下指标条：本期/上期/去年同期/环比/同比（供给·需求类另有累计同比%）
  // kinds: val=绝对值 / delta=差值（正红负绿）/ pct=百分比（正红负绿）
  function buildChartData(ch) {
    var t = ch.table;
    var wrap = document.createElement('div');
    wrap.className = 'chart-data';
    var html = '<table><thead><tr>';
    t.labels.forEach(function (l) { html += '<th>' + l + '</th>'; });
    html += '</tr></thead><tbody><tr>';
    t.values.forEach(function (v, i) {
      var kind = (t.kinds && t.kinds[i]) || 'val';
      var txt = '—', cls = '';
      if (v !== null && v !== undefined && !isNaN(v)) {
        if (kind === 'pct') {
          txt = (v > 0 ? '+' : '') + fmtCardNum(v) + '%';
          cls = v > 0 ? 'up' : (v < 0 ? 'down' : '');
        } else if (kind === 'delta') {
          txt = (v > 0 ? '+' : '') + fmtCardNum(v);
          cls = v > 0 ? 'up' : (v < 0 ? 'down' : '');
        } else {
          txt = fmtCardNum(v);
        }
      }
      html += cls ? '<td class="' + cls + '">' + txt + '</td>' : '<td>' + txt + '</td>';
    });
    html += '</tr></tbody></table>';
    wrap.innerHTML = html;
    return wrap;
  }

  // 卡片脚：数据截至日期 + 下载 PNG 按钮（2× 清晰度；未初始化时先初始化该图）
  function buildCardFoot(ch, ds, box) {
    var foot = document.createElement('div');
    foot.className = 'card-foot';
    var left = document.createElement('span');
    left.textContent = '数据截至 ' + (ds.asOf || '-');
    foot.appendChild(left);

    var btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'dl';
    btn.title = '下载该图 PNG（2× 清晰度）';
    btn.textContent = '⤓ 下载图片';
    btn.onclick = function () {
      var it = findByEl(box);
      if (!it) return;
      if (!it.inst) initChart(it);
      if (!it.inst) return;
      var a = document.createElement('a');
      a.download = (ch.title || 'chart') + '.png';
      a.href = it.inst.getDataURL({ type: 'png', pixelRatio: 2, backgroundColor: '#fff' });
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    };
    foot.appendChild(btn);
    return foot;
  }

  function renderSummaryTable(ds) {
    var s = ds.summary;
    if (!s || !s.columns || !s.columns.length) return null;
    var asPct = !!s.asPercent;   // 盈利率：比例值→前端 ×100 显示百分比

    // 合并分组（连续相同 group 名算一段）
    var groups = [];
    var cur = null;
    s.columns.forEach(function (c) {
      if (!cur || cur.name !== c.group) {
        cur = { name: c.group, start: groups.length === 0 ? 0 : cur.end + 1, end: 0 };
        cur.start = cur.end = s.columns.indexOf(c);
        groups.push(cur);
      } else {
        cur.end = s.columns.indexOf(c);
      }
    });

    var wrap = document.createElement('div');
    wrap.className = 'summary';

    var head = document.createElement('div');
    head.className = 'summary-head';
    head.innerHTML =
      '<strong>' + ds.name + ' · 本期汇总</strong>' +
      '<span class="summary-dates">本期 ' + s.currentWeek +
      ' · 上期 ' + s.previousWeek + ' · 单位：' + (s.unit || '万吨') + '</span>';
    wrap.appendChild(head);

    var scroll = document.createElement('div');
    scroll.className = 'summary-scroll';

    var table = document.createElement('table');
    table.className = 'summary-table';

    // 列宽策略：
    //   全球发运 / 分省份盈利率 / 分省份铁水 → 固定像素宽 + 横向滚动（列多不可压扁）
    //     全球用长图截图量出的像素；省份用 标签列190 + 数据列58（够装数值与省名）
    //   其余（分区域盈利率/铁水、澳巴、到港、15港）→ 等百分比宽（#35 原始行为）
    var GLOBAL_FIT_PX = [120, 109, 84, 85, 90, 91, 91, 92, 106, 90, 90, 127, 90, 91, 109, 109, 109, 91, 84, 84];
    var PX_FIT_IDS = { global: true, profit_province: true, ironwater_province: true, juanluo_luowen: true, juanluo_rejuan: true, variety_inventory: true };
    var colgroup = document.createElement('colgroup');
    var totalCols = s.columns.length + 1;
    if (PX_FIT_IDS[ds.id]) {
      table.classList.add('global-fit');
      if (ds.id === 'global') {
        for (var gi = 0; gi < totalCols; gi++) {
          var gc = document.createElement('col');
          gc.style.width = (GLOBAL_FIT_PX[gi] || 90) + 'px';
          colgroup.appendChild(gc);
        }
      } else {
        var LABEL_W = 190, DATA_W = 58;
        var lc = document.createElement('col');
        lc.style.width = LABEL_W + 'px';
        colgroup.appendChild(lc);
        for (var di = 0; di < s.columns.length; di++) {
          var dc = document.createElement('col');
          dc.style.width = DATA_W + 'px';
          colgroup.appendChild(dc);
        }
      }
    } else {
      var each = (100 / totalCols).toFixed(3);
      for (var ci = 0; ci < totalCols; ci++) {
        var cc = document.createElement('col');
        cc.style.width = each + '%';
        colgroup.appendChild(cc);
      }
    }
    table.appendChild(colgroup);

    // thead: 分组行 + 列名行
    var thead = document.createElement('thead');
    var grpRow = document.createElement('tr');
    grpRow.className = 'grp-row';
    var th0 = document.createElement('th');
    th0.className = 'row-label';
    th0.rowSpan = 2;
    th0.textContent = '指标';
    grpRow.appendChild(th0);
    groups.forEach(function (g) {
      var th = document.createElement('th');
      th.colSpan = g.end - g.start + 1;
      th.className = 'grp grp-' + g.name;
      th.textContent = g.name;
      grpRow.appendChild(th);
    });
    thead.appendChild(grpRow);

    var labRow = document.createElement('tr');
    labRow.className = 'lab-row';
    s.columns.forEach(function (c) {
      var th = document.createElement('th');
      th.textContent = c.label;
      th.title = c.label;
      labRow.appendChild(th);
    });
    thead.appendChild(labRow);
    table.appendChild(thead);

    // tbody: 5 行（本期/上期/环比/累计同比/累计增幅%）
    var ROW_META = {
      '本期':      {},
      '上期':      {},
      '环比':      {signed: true},
      '累计同比':  {signed: true},
      '累计同比%': {signed: true, suffix: '%'},
      '累计增幅%': {signed: true, suffix: '%'},
      '较年初变化': {signed: true},
      '较年初':    {signed: true},
      '较年初%':   {signed: true, suffix: '%'},
      '同比':      {signed: true},
      '同比%':     {signed: true, suffix: '%'}
    };
    var rowOrder = s.rowOrder || ['本期', '上期', '环比', '累计同比', '累计增幅%'];
    var rowDefs = rowOrder.map(function (k) {
      // 复合键 '指标·视图'（profit/ironwater 用）：完整键作 label，让4 行字面就不同
      // （如「盈利率·本期（08-28）」「盈利率·较年初」），不再剥掉视图段。
      // 纯键（aubra/global/arrival/portinv 的 本期/上期/环比 等）label = k 自身。
      var seg = k.split('·');
      var view = seg[seg.length - 1];
      var label = k;
      var meta = ROW_META[view] || {};
      var d = { key: k, label: label };
      if (view === '本期') d.date = s.currentWeek;
      if (view === '上期') d.date = s.previousWeek;
      if (meta.signed) d.signed = true;
      if (meta.suffix) d.suffix = meta.suffix;
      if (asPct) d.suffix = (d.suffix || '') + '%';   // 盈利率：比例值→百分比显示
      // 差值行：环比 / 较年初 / 同比 / 累计同比 / 累计增幅% — 加 r-delta class，CSS 给浅灰背景
      d.isDelta = meta.signed === true;
      return d;
    });
    var tbody = document.createElement('tbody');
    rowDefs.forEach(function (rd) {
      var tr = document.createElement('tr');
      tr.className = 'r-' + rd.key + (rd.isDelta ? ' r-delta' : '');
      var lbl = document.createElement('td');
      lbl.className = 'row-label';
      lbl.textContent = rd.label + (rd.date ? ' (' + rd.date + ')' : '');
      tr.appendChild(lbl);

      var vals = s.rows[rd.key] || [];
      s.columns.forEach(function (c, i) {
        var td = document.createElement('td');
        var raw = vals[i];
        if (raw === null || raw === undefined) {
          td.textContent = '—';
          td.className = 'na';
        } else {
          var v = asPct ? raw * 100 : raw;   // 盈利率：比例值 ×100 显示百分比
          var sign = (rd.signed && v > 0) ? '+' : '';
          td.textContent = sign + fmtNum(v, rd.suffix || '');
          if (rd.signed) {
            if (v > 0) td.className = 'up';
            else if (v < 0) td.className = 'down';
          }
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    scroll.appendChild(table);
    wrap.appendChild(scroll);

    // 数据来源说明 / 口径备注（如钢厂日接单的「日产参考」）
    if (ds.note) {
      var note = document.createElement('div');
      note.className = 'summary-note';
      note.textContent = ds.note;
      wrap.appendChild(note);
    }
    return wrap;
  }

  /* ---------------- 渲染 ---------------- */
  function datasetById(id) {
    for (var i = 0; i < S.data.datasets.length; i++) {
      if (S.data.datasets[i].id === id) return S.data.datasets[i];
    }
    return S.data.datasets[0];
  }

  function initChart(item) {
    if (item.inst) return;
    var inst = echarts.init(item.el, null, { renderer: 'canvas' });
    item.inst = inst;
    inst.setOption(buildOption(item.chart, item.ds, null));
  }

  function refreshAll() {
    S.items.forEach(function (it) {
      if (!it.inst) return;
      it.inst.setOption(buildOption(it.chart, it.ds, null), true);
    });
    if (S.opt.sync) syncY();
  }

  /* 同行同 Y：按实际 offsetTop 分组，组内共享 Y 范围 */
  function syncY() {
    var rows = {};
    S.items.forEach(function (it) {
      if (!it.inst || !it.el || !it.chart.series) return;
      var top = it.card.offsetTop;
      (rows[top] = rows[top] || []).push(it);
    });
    Object.keys(rows).forEach(function (k) {
      var list = rows[k], mn = Infinity, mx = -Infinity;
      list.forEach(function (it) {
        it.chart.series.forEach(function (s) {
          if (S.hidden[s.name]) return;
          s.data.forEach(function (v) {
            if (v === null || v === undefined) return;
            if (v < mn) mn = v;
            if (v > mx) mx = v;
          });
        });
      });
      if (mn === Infinity) return;
      var pad = (mx - mn) * 0.06 || 1;
      var r = { min: S.opt.zero ? 0 : Math.max(0, mn - pad), max: mx + pad };
      list.forEach(function (it) {
        it.inst.setOption({ yAxis: { min: r.min, max: r.max } });
      });
    });
  }

  // 让 iframe 高度贴合其内部静态页面（同源可读 scrollHeight），避免出现双层滚动条
  function fitEmbedHeight(fr) {
    function apply() {
      try {
        var doc = fr.contentDocument;
        if (!doc) return;
        var h = Math.max(doc.documentElement.scrollHeight, doc.body ? doc.body.scrollHeight : 0);
        if (h > 200) fr.style.height = (h + 8) + 'px';
      } catch (e) { /* 跨域兜底：保持 min-height */ }
    }
    apply();
    // 图片解码完成后高度会变大，再补两次
    setTimeout(apply, 400);
    setTimeout(apply, 1500);
  }

  function renderDataset(dsId) {
    S.dsId = dsId;
    var ds = datasetById(dsId);

    // ---- 外链数据集：新窗口打开，不渲染本地内容 ----
    if (ds.link) {
      window.open(ds.link, '_blank');
      return;
    }

    S.dateAxis = (ds.axisType === 'date');
    S.dateTicks = (S.dateAxis && ds.dateTicks) ? ds.dateTicks : null;
    S.monthFirst = (ds.monthFirst && !S.dateAxis) ? ds.monthFirst : null;
    var main = $('main');

    if (S.io) { S.io.disconnect(); S.io = null; }
    S.items.forEach(function (it) { if (it.inst) it.inst.dispose(); });
    S.items = [];

    // ---- 嵌入式数据集：iframe 原样嵌入静态看板页，保持其原有格式 ----
    if (ds.embed) {
      main.innerHTML = '';
      main.setAttribute('data-ds', dsId);
      var fr = document.createElement('iframe');
      fr.className = 'embed-frame';
      fr.setAttribute('scrolling', 'no');
      fr.onload = function () { fitEmbedHeight(fr); };
      fr.src = ds.embed;
      main.appendChild(fr);
      S.years = [];
      renderYearToggles();
      window.scrollTo(0, 0);
      return;
    }

    // 汇总表（在图表网格之前）
    var sumEl = renderSummaryTable(ds);
    main.innerHTML = '';
    main.setAttribute('data-ds', dsId);   // 数据集级样式钩子（如带钢加宽）
    if (sumEl) main.appendChild(sumEl);

    // 按 ch.group 分区块渲染（钢材站：一个区块一组指标，区块内 5 图一行）
    var curGroup = null, grid = null;
    ds.charts.forEach(function (ch) {
      var g = ch.group || '';
      if (g !== curGroup) {
        curGroup = g;
        var host = main;
        if (g) {
          var sec = document.createElement('section');
          sec.className = 'rgroup';
          var h = document.createElement('h2');
          h.className = 'rgroup-title';
          h.textContent = g;
          sec.appendChild(h);
          main.appendChild(sec);
          host = sec;
        }
        grid = document.createElement('div');
        // 区块列数：默认 5 列；按数据集例外（PSI/带钢/出港/出口等一排 3 张）；
        // chart.full = true 的图（如排序柱）单独占满一整行
        var gcols = ch.cols ? (' cols' + ch.cols)
                            : (ch.full ? ' cols1'
                                       : (GRID_COLS[ds.id] ? (' cols' + GRID_COLS[ds.id]) : ''));
        grid.className = 'grid rgrid' + gcols;
        host.appendChild(grid);
      }
      var card = document.createElement('div');
      card.className = 'card' + (S.expanded ? ' wide' : '');

      var head = document.createElement('div');
      head.className = 'card-head';
      head.innerHTML = '<span>' + ch.title + '</span>';
      var tag = document.createElement('span');
      tag.className = 'tag';
      tag.textContent = ch.tag || (ch.series ? (ch.series.length > 1 ? (ch.series.length + ' 年对比') : '年内走势') : '排序');
      head.appendChild(tag);

      var box = document.createElement('div');
      box.className = 'chart';

      card.appendChild(head);
      card.appendChild(box);
      // 图下指标条 + 下载图片：仅当数据带 chart.table 时渲染（表外 / 分品种表需 两个页面）
      if (ch.table && ch.table.labels && ch.table.values) {
        card.appendChild(buildChartData(ch));
        card.appendChild(buildCardFoot(ch, ds, box));
      }
      grid.appendChild(card);

      S.items.push({ key: ch.key, chart: ch, ds: ds, el: box, card: card, inst: null });
    });

    // 年份开关（按当前数据集的系列名）
    var years = [];
    if (ds.charts[0] && ds.charts[0].series) {
      ds.charts[0].series.forEach(function (s) { years.push({ name: s.name, color: s.color }); });
    }
    S.years = years;
    renderYearToggles();

    if (S.shot) {
      document.body.classList.add('shot');
      S.items.forEach(function (it) { initChart(it); });
      S.items.forEach(function (it) { if (it.inst) it.inst.resize(); });
    } else {
      document.body.classList.remove('shot');
      if ('IntersectionObserver' in window) {
        S.io = new IntersectionObserver(function (entries) {
          entries.forEach(function (e) {
            if (e.isIntersecting) {
              var it = findByEl(e.target);
              if (it) initChart(it);
              S.io.unobserve(e.target);
            }
          });
        }, { rootMargin: '300px' });
        S.items.forEach(function (it) { S.io.observe(it.el); });
      } else {
        S.items.forEach(function (it) { initChart(it); });
      }
    }

    if (S.opt.sync) setTimeout(syncY, 60);
    window.scrollTo(0, 0);
  }

  function findByEl(el) {
    for (var i = 0; i < S.items.length; i++) if (S.items[i].el === el) return S.items[i];
    return null;
  }

  function renderYearToggles() {
    var box = $('yearToggles');
    box.innerHTML = '';
    if (S.dateAxis) return;   // 连续日期轴无「年份对比」开关
    S.years.forEach(function (y) {
      var on = !S.hidden[y.name];
      var b = document.createElement('button');
      b.className = 'ytoggle' + (on ? ' on' : '');
      b.innerHTML = '<span class="dot" style="background:' + y.color + '"></span>' + y.name;
      b.onclick = function () {
        S.hidden[y.name] = !S.hidden[y.name];
        b.className = 'ytoggle' + (!S.hidden[y.name] ? ' on' : '');
        refreshAll();
      };
      box.appendChild(b);
    });
  }

  function renderTabs() {
    var box = $('tabs');
    box.innerHTML = '';
    S.data.datasets.forEach(function (d) {
      var b = document.createElement('button');
      b.className = 'tab' + (d.id === S.dsId ? ' active' : '');
      b.innerHTML = d.name + '<span class="cnt">' +
        (d.link ? '外链' : (d.embed ? '看板' : d.charts.length)) + '</span>';
      if (d.link) {
        b.title = '点击在新窗口打开：' + d.link;
        b.onclick = function () { window.open(d.link, '_blank'); };
      } else {
        b.onclick = function () {
          if (d.id === S.dsId) return;
          renderDataset(d.id);
          renderTabs();
        };
      }
      box.appendChild(b);
    });
  }

  function bindToolbar() {
    $('optConnect').onchange = function () { S.opt.connect = this.checked; refreshAll(); };
    $('optZero').onchange = function () { S.opt.zero = this.checked; refreshAll(); };
    $('optSync').onchange = function () {
      S.opt.sync = this.checked;
      refreshAll();
      if (!this.checked) {
        S.items.forEach(function (it) {
          if (it.inst) it.inst.setOption({ yAxis: { min: null, max: null } });
        });
      }
    };
    $('btnExpand').onclick = function () {
      S.expanded = !S.expanded;
      this.textContent = S.expanded ? '收起全部' : '全部展开';
      S.items.forEach(function (it) {
        it.card.classList.toggle('wide', S.expanded);
      });
      setTimeout(function () {
        S.items.forEach(function (it) { if (it.inst) it.inst.resize(); });
        if (S.opt.sync) syncY();
      }, 120);
    };
    $('btnTop').onclick = function () { window.scrollTo({ top: 0, behavior: 'smooth' }); };
  }

  function boot() {
    var params = new URLSearchParams(location.search);
    S.shot = params.get('shot') === '1';
    var want = params.get('ds');

    loadData(function () {
      // 注入嵌入式数据集（追加在末尾，不影响原有默认首页）
      EMBED_DATASETS.forEach(function (e) {
        var exists = S.data.datasets.some(function (d) { return d.id === e.id; });
        if (!exists) {
          S.data.datasets.push({ id: e.id, name: e.name, charts: [], embed: e.src, asOf: e.asOf });
        }
      });
      // 注入外链数据集（只加 tab 入口，页面本身在外部网址）
      LINK_DATASETS.forEach(function (e) {
        var exists = S.data.datasets.some(function (d) { return d.id === e.id; });
        if (!exists) {
          S.data.datasets.push({ id: e.id, name: e.name, charts: [], link: e.url, asOf: e.asOf });
        }
      });
      var ds = (want && datasetById(want).id === want) ? datasetById(want) : S.data.datasets[0];
      if (ds.link) { window.open(ds.link, '_blank'); ds = S.data.datasets[0]; }
      S.dsId = ds.id;

      var d = new Date(S.data.updated.replace('T', ' ').replace(/-/g, '/'));
      var stamp = isNaN(d) ? S.data.updated : d.toLocaleString('zh-CN', { hour12: false });
      var parts = S.data.datasets.map(function (x) {
        if (x.link) return x.name + '（外链 · 至 ' + x.asOf + '）';
        return x.embed ? (x.name + '（至 ' + x.asOf + '）')
                       : (x.name + ' ' + x.charts.length + ' 图（至 ' + x.asOf + '）');
      });
      $('sub').textContent = '数据更新 ' + stamp + ' · ' + parts.join(' ｜ ');
      var nCharts = S.data.datasets.reduce(function (a, b) { return a + b.charts.length; }, 0);
      var nEmbed = S.data.datasets.filter(function (x) { return x.embed || x.link; }).length;
      $('footNote').textContent = '共 ' + nCharts + ' 张图' +
        (nEmbed ? ' + ' + nEmbed + ' 个直供看板/外链' : '') + ' · 数据提取自本地 Excel 数据库';

      renderTabs();
      renderDataset(S.dsId);
      bindToolbar();

      window.addEventListener('resize', function () {
        S.items.forEach(function (it) { if (it.inst) it.inst.resize(); });
        if (S.opt.sync) syncY();
      });
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
})();
