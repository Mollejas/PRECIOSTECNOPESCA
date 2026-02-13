<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="GridPrecios.aspx.vb" Inherits="TEST1.GridPrecios" %>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Grid Precios - fcuapr1</title>
    <link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600&family=DM+Sans:wght@400;500;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg: #0c0f1a;
            --surface: #131829;
            --surface2: #1a2035;
            --border: #252d45;
            --accent: #38bdf8;
            --accent2: #22d3ee;
            --green: #34d399;
            --red: #f87171;
            --orange: #fb923c;
            --text: #e2e8f0;
            --text2: #94a3b8;
            --text3: #64748b;
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'DM Sans', sans-serif;
            background: var(--bg);
            color: var(--text);
            min-height: 100vh;
            padding: 25px;
        }
        .header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 20px;
            flex-wrap: wrap;
            gap: 12px;
        }
        .header h1 {
            font-size: 1.4rem;
            font-weight: 700;
            color: var(--accent);
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .header h1 span {
            font-size: 0.7rem;
            background: var(--surface2);
            padding: 3px 10px;
            border-radius: 20px;
            color: var(--text3);
            font-weight: 400;
        }
        .status-bar {
            display: flex;
            align-items: center;
            gap: 6px;
            font-size: 0.8rem;
            color: var(--text3);
        }
        .status-dot {
            width: 8px; height: 8px;
            border-radius: 50%;
            background: var(--green);
        }
        .toolbar {
            display: flex;
            gap: 8px;
            margin-bottom: 15px;
            flex-wrap: wrap;
        }
        .btn {
            padding: 7px 16px;
            border: 1px solid var(--border);
            border-radius: 6px;
            cursor: pointer;
            font-size: 0.8rem;
            font-weight: 500;
            font-family: 'DM Sans', sans-serif;
            background: var(--surface2);
            color: var(--text2);
            transition: all 0.15s;
            display: flex;
            align-items: center;
            gap: 5px;
        }
        .btn:hover { border-color: var(--accent); color: var(--accent); }
        .btn-primary { background: var(--accent); color: #0c0f1a; border-color: var(--accent); font-weight: 700; }
        .btn-primary:hover { background: var(--accent2); border-color: var(--accent2); color: #0c0f1a; }
        .btn-success { background: var(--green); color: #0c0f1a; border-color: var(--green); font-weight: 700; }
        .btn-success:hover { background: #2cc88a; }
        .btn-danger { border-color: var(--red); color: var(--red); }
        .btn-danger:hover { background: rgba(248,113,113,0.1); }

        .grid-wrap {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 10px;
            overflow: hidden;
        }
        .grid-scroll {
            max-height: 62vh;
            overflow-y: auto;
        }
        .grid-scroll::-webkit-scrollbar { width: 6px; }
        .grid-scroll::-webkit-scrollbar-track { background: var(--surface); }
        .grid-scroll::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }
        table { width: 100%; border-collapse: collapse; }
        thead th {
            background: var(--surface2);
            padding: 10px 14px;
            text-align: left;
            font-size: 0.72rem;
            text-transform: uppercase;
            letter-spacing: 1.2px;
            color: var(--text3);
            position: sticky;
            top: 0;
            z-index: 10;
            border-bottom: 1px solid var(--border);
        }
        tbody tr {
            border-bottom: 1px solid rgba(37,45,69,0.5);
            transition: background 0.1s;
        }
        tbody tr:hover { background: rgba(56,189,248,0.03); }
        tbody tr.modified { background: rgba(251,146,60,0.06); border-left: 3px solid var(--orange); }
        tbody tr.not-found { background: rgba(248,113,113,0.06); border-left: 3px solid var(--red); }
        tbody tr.updated { background: rgba(52,211,153,0.06); border-left: 3px solid var(--green); }
        td { padding: 0; }
        td input {
            width: 100%;
            padding: 9px 14px;
            background: transparent;
            border: none;
            color: var(--text);
            font-size: 0.85rem;
            font-family: 'JetBrains Mono', monospace;
            outline: none;
        }
        td input:focus {
            background: rgba(56,189,248,0.05);
            box-shadow: inset 0 0 0 1px rgba(56,189,248,0.2);
        }
        td input.precio { text-align: right; }
        td .row-num {
            padding: 9px 10px;
            color: var(--text3);
            font-size: 0.72rem;
            font-family: 'JetBrains Mono', monospace;
            text-align: center;
        }
        td button.btn-x {
            background: none; border: none;
            color: var(--text3); cursor: pointer;
            padding: 6px 10px; font-size: 0.8rem;
        }
        td button.btn-x:hover { color: var(--red); }

        .footer-bar {
            display: flex;
            justify-content: space-between;
            padding: 10px 14px;
            background: var(--surface2);
            border-top: 1px solid var(--border);
            font-size: 0.78rem;
            color: var(--text3);
            flex-wrap: wrap;
            gap: 10px;
        }
        .footer-bar .stat { display: flex; align-items: center; gap: 5px; }
        .footer-bar .stat b { color: var(--text2); }

        .hint-box {
            margin-top: 18px;
            padding: 14px 20px;
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 8px;
            border-left: 3px solid var(--accent);
        }
        .hint-box h3 { font-size: 0.82rem; color: var(--accent); margin-bottom: 6px; }
        .hint-box p { font-size: 0.78rem; color: var(--text3); line-height: 1.6; }
        .hint-box kbd {
            background: var(--surface2);
            border: 1px solid var(--border);
            padding: 1px 6px;
            border-radius: 3px;
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.72rem;
            color: var(--accent);
        }

        .toast-container {
            position: fixed; bottom: 20px; right: 20px;
            display: flex; flex-direction: column; gap: 8px; z-index: 1000;
        }
        .toast {
            padding: 10px 18px; border-radius: 8px;
            font-size: 0.82rem; font-weight: 500; color: #fff;
            opacity: 0; transform: translateX(30px);
            transition: all 0.3s; max-width: 350px;
        }
        .toast.show { opacity: 1; transform: translateX(0); }
        .toast.success { background: #059669; }
        .toast.error { background: #dc2626; }
        .toast.info { background: #2563eb; }
        .toast.warning { background: #d97706; }

        .loading { pointer-events: none; opacity: 0.5; }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods="true" />

    <div class="header">
        <h1>⚡ Grid Precios <span>fcuapr1.dbf</span></h1>
        <div class="status-bar">
            <div class="status-dot"></div>
            <span>Conexión directa VB.NET → DBF</span>
        </div>
    </div>

    <div class="toolbar">
        <button type="button" class="btn" onclick="addRow()">＋ Fila</button>
        <button type="button" class="btn" onclick="addRows(10)">＋ 10 Filas</button>
        <button type="button" class="btn btn-primary" onclick="cargarDBF()">📂 Cargar de DBF</button>
        <button type="button" class="btn-success btn" onclick="guardarCambios()">💾 Guardar en DBF</button>
        <button type="button" class="btn btn-danger" onclick="limpiarGrid()">🗑 Limpiar</button>
    </div>

    <div class="grid-wrap" id="gridWrap">
        <div class="grid-scroll" id="gridScroll">
            <table id="grid">
                <thead>
                    <tr>
                        <th style="width:35px">#</th>
                        <th style="width:35%">APRCLAVE</th>
                        <th style="width:15%">APRLISTA</th>
                        <th style="width:25%">Precio</th>
                        <th style="width:35px"></th>
                    </tr>
                </thead>
                <tbody id="gridBody"></tbody>
            </table>
        </div>
        <div class="footer-bar">
            <div class="stat">Filas: <b id="sRows">0</b></div>
            <div class="stat">Con datos: <b id="sMod">0</b></div>
            <div class="stat">Errores: <b id="sErr">0</b></div>
        </div>
    </div>

    <div class="hint-box">
        <h3>💡 Instrucciones</h3>
        <p>
            1. En Excel, copia 3 columnas: <strong>Código, Lista, Precio</strong>.<br>
            2. Click en la primera celda de APRCLAVE y pega con <kbd>Ctrl+V</kbd>.<br>
            3. Click en <strong>💾 Guardar en DBF</strong> para actualizar los precios en la base de datos.<br>
            <br>
            O usa <strong>📂 Cargar de DBF</strong> para ver todos los registros y editar precios directo.
        </p>
    </div>

    </form>

    <div class="toast-container" id="toastBox"></div>

    <script>
        let rowN = 0;

        function esc(s) {
            return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
        }

        // ========================
        // FILAS
        // ========================
        function createRow(clave, lista, precio) {
            clave = clave || ''; lista = lista || ''; precio = precio || '';
            rowN++;
            const tr = document.createElement('tr');
            tr.innerHTML =
                '<td><div class="row-num">' + rowN + '</div></td>' +
                '<td><input type="text" class="c-clave" value="' + esc(clave) + '" placeholder="Código"></td>' +
                '<td><input type="text" class="c-lista" value="' + esc(lista) + '" placeholder="001"></td>' +
                '<td><input type="text" class="c-precio precio" value="' + esc(precio) + '" placeholder="Precio"></td>' +
                '<td><button type="button" class="btn-x" onclick="delRow(this)" title="Eliminar">✕</button></td>';

            tr.querySelectorAll('input').forEach(function (inp) {
                inp.addEventListener('paste', handlePaste);
            });

            document.getElementById('gridBody').appendChild(tr);
            updateStats();
            return tr;
        }

        function addRow() { var tr = createRow(); tr.querySelector('.c-clave').focus(); }
        function addRows(n) { for (var i = 0; i < n; i++) createRow(); toast(n + ' filas agregadas', 'info'); }

        function delRow(btn) {
            btn.closest('tr').remove();
            renumber();
            updateStats();
        }

        function limpiarGrid() {
            if (!confirm('¿Limpiar todas las filas?')) return;
            document.getElementById('gridBody').innerHTML = '';
            rowN = 0;
            updateStats();
            addRows(5);
        }

        function renumber() {
            rowN = 0;
            document.querySelectorAll('#gridBody tr').forEach(function (r) {
                rowN++;
                r.querySelector('.row-num').textContent = rowN;
            });
        }

        function updateStats() {
            var allRows = document.querySelectorAll('#gridBody tr');
            var conDatos = 0;
            allRows.forEach(function (tr) {
                var clave = tr.querySelector('.c-clave').value.trim();
                var lista = tr.querySelector('.c-lista').value.trim();
                var precio = tr.querySelector('.c-precio').value.trim();
                if (clave && lista && precio) conDatos++;
            });
            document.getElementById('sRows').textContent = allRows.length;
            document.getElementById('sMod').textContent = conDatos;
            document.getElementById('sErr').textContent = document.querySelectorAll('#gridBody tr.not-found').length;
        }

        function handlePaste(e) {
            e.preventDefault();
            var text = (e.clipboardData || window.clipboardData).getData('text');
            if (!text.trim()) return;

            var lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n').filter(function (l) { return l.trim(); });
            var isMulti = lines.some(function (l) { return l.indexOf('\t') >= 0; });

            var colMap = ['c-clave', 'c-lista', 'c-precio'];
            var inp = e.target;
            var curRow = inp.closest('tr');
            var colClass = colMap.filter(function (c) { return inp.classList.contains(c); })[0] || colMap[0];
            var startCol = colMap.indexOf(colClass);

            var allRows = Array.from(document.querySelectorAll('#gridBody tr'));
            var startIdx = allRows.indexOf(curRow);

            lines.forEach(function (line, i) {
                var cells = line.split('\t');
                var row;
                if (startIdx + i < allRows.length) {
                    row = allRows[startIdx + i];
                } else {
                    row = createRow();
                    allRows.push(row);
                }

                if (isMulti) {
                    cells.forEach(function (cell, j) {
                        var tIdx = startCol + j;
                        if (tIdx < colMap.length) {
                            var el = row.querySelector('.' + colMap[tIdx]);
                            if (el) el.value = cell.trim();
                        }
                    });
                } else {
                    var el = row.querySelector('.' + colClass);
                    if (el) el.value = line.trim();
                }
            });

            toast(lines.length + ' filas pegadas', 'info');
            updateStats();
        }

        function cargarDBF() {
            document.getElementById('gridWrap').classList.add('loading');
            toast('Cargando datos...', 'info');

            PageMethods.CargarTodos(function (raw) {
                var data = raw.d || raw || [];
                document.getElementById('gridBody').innerHTML = '';
                rowN = 0;
                data.forEach(function (item) {
                    createRow(item.Clave, item.Lista, item.Precio);
                });
                document.getElementById('gridWrap').classList.remove('loading');
                toast(data.length + ' registros cargados', 'success');
            }, function (err) {
                document.getElementById('gridWrap').classList.remove('loading');
                toast('Error: ' + err.get_message(), 'error');
            });
        }

        function guardarCambios() {
            var allRows = Array.from(document.querySelectorAll('#gridBody tr'));
            var items = [];
            var filasValidas = [];

            allRows.forEach(function (tr) {
                var clave = tr.querySelector('.c-clave').value.trim();
                var lista = tr.querySelector('.c-lista').value.trim();
                var precio = tr.querySelector('.c-precio').value.trim();
                if (clave && lista && precio) {
                    items.push({ Clave: clave, Lista: lista, Precio: precio });
                    filasValidas.push(tr);
                }
            });

            if (items.length === 0) {
                toast('No hay filas con datos completos para guardar', 'warning');
                return;
            }

            if (!confirm('¿Actualizar ' + items.length + ' registro(s) en la DBF?')) return;

            document.getElementById('gridWrap').classList.add('loading');

            PageMethods.GuardarPrecios(items, function (raw) {
                var result = raw.d || raw || {};
                document.getElementById('gridWrap').classList.remove('loading');

                toast((result.Actualizados || 0) + ' de ' + (result.Total || 0) + ' actualizados', 'success');

                filasValidas.forEach(function (tr) {
                    tr.className = 'updated';
                });

                var errores = result.Errores || [];
                if (errores.length > 0) {
                    errores.forEach(function (e) {
                        toast(e, 'error');
                    });
                }
                updateStats();
            }, function (err) {
                document.getElementById('gridWrap').classList.remove('loading');
                toast('Error: ' + err.get_message(), 'error');
            });
        }

        function toast(msg, type) {
            var el = document.createElement('div');
            el.className = 'toast ' + (type || 'info');
            el.textContent = msg;
            document.getElementById('toastBox').appendChild(el);
            requestAnimationFrame(function () { el.classList.add('show'); });
            setTimeout(function () {
                el.classList.remove('show');
                setTimeout(function () { el.remove(); }, 300);
            }, 3000);
        }

        addRows(5);
    </script>
</body>
</html>
