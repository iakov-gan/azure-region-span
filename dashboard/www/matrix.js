/**
 * Matrix View for Azure Region Latency Dashboard
 * 
 * Renders a heatmap matrix of source×destination latencies with:
 * - Color-coded cells (green→red by latency)
 * - Sticky row/column headers
 * - Sparkline history on cell hover
 * - Region info on header hover
 * - Geography filter, multi-select source/dest, sort, CSV export
 * 
 * Self-contained: all CSS is injected via JS.
 */

class MatrixView {
    constructor(containerEl) {
        this.container = containerEl;
        this.rawData = [];
        this.aggregatedData = null;
        this.sourceRegions = [];
        this.destRegions = [];
        this.filteredSources = [];
        this.filteredDests = [];
        this.sortMode = 'name'; // 'name' or 'latency'
        this.geoFilter = 'all';
        this.historyMap = new Map(); // "src|dst" → [{latency, timestamp}]
        this.latencyLookup = new Map(); // "src|dst" → latency
        this._styleInjected = false;
        this._tooltipEl = null;
        this._sparklineEl = null;
    }

    /** Called once after data loads */
    init(rawData, aggregatedData) {
        this.rawData = rawData || [];
        this.aggregatedData = aggregatedData;
        if (!aggregatedData) return;

        // Build history map from raw data
        this._buildHistoryMap();

        // Build latency lookup from aggregated data
        this.aggregatedData.connections.forEach(c => {
            this.latencyLookup.set(`${c.source}|${c.destination}`, c.latency);
        });

        // Get region lists
        this.sourceRegions = [...this.aggregatedData.regions].sort();
        this.destRegions = [...this.aggregatedData.regions].sort();
        this.filteredSources = [...this.sourceRegions];
        this.filteredDests = [...this.destRegions];

        // Read URL params for matrix filters
        this._readURLParams();

        this._injectStyles();
        this._render();
    }

    _readURLParams() {
        const params = new URLSearchParams(window.location.search);
        if (params.get('mgeo')) this.geoFilter = params.get('mgeo');
        if (params.get('msort')) this.sortMode = params.get('msort');
        if (params.get('msrc')) {
            this._sourceSelection = new Set(params.get('msrc').split(','));
        }
        if (params.get('mdst')) {
            this._destSelection = new Set(params.get('mdst').split(','));
        }
    }

    _syncURLParams() {
        const url = new URL(window.location);
        if (this.geoFilter !== 'all') url.searchParams.set('mgeo', this.geoFilter);
        else url.searchParams.delete('mgeo');
        if (this.sortMode !== 'name') url.searchParams.set('msort', this.sortMode);
        else url.searchParams.delete('msort');
        // Source selection
        if (this._sourceSelection && this._sourceSelection.size > 0 && this._sourceSelection.size < this.sourceRegions.length) {
            url.searchParams.set('msrc', [...this._sourceSelection].join(','));
        } else {
            url.searchParams.delete('msrc');
        }
        // Dest selection
        if (this._destSelection && this._destSelection.size > 0 && this._destSelection.size < this.destRegions.length) {
            url.searchParams.set('mdst', [...this._destSelection].join(','));
        } else {
            url.searchParams.delete('mdst');
        }
        url.searchParams.set('view', 'matrix');
        history.replaceState(null, '', url);
    }

    show() {
        this.container.style.display = '';
        // Re-render in case data changed
        if (this.aggregatedData) this._render();
    }

    hide() {
        this.container.style.display = 'none';
        this._hideTooltip();
        this._hideSparkline();
    }

    // ── Data helpers ──────────────────────────────────────────

    _buildHistoryMap() {
        this.historyMap.clear();
        this.rawData.forEach(entity => {
            const src = entity.source;
            const dst = entity.destination;
            const latencyRaw = entity.latency;
            const timestamp = entity.timestamp;
            if (!src || !dst || !latencyRaw) return;
            const latency = this._parseLatency(latencyRaw);
            if (latency === null) return;
            const key = `${src}|${dst}`;
            if (!this.historyMap.has(key)) this.historyMap.set(key, []);
            this.historyMap.get(key).push({ latency, timestamp });
        });
        // Sort each history by timestamp
        this.historyMap.forEach(arr => {
            arr.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
        });
    }

    _parseLatency(str) {
        if (typeof str === 'number') return str;
        if (!str) return null;
        const s = String(str).trim().toLowerCase();
        const num = parseFloat(s);
        if (isNaN(num)) return null;
        if (s.includes('us') || s.includes('µs')) return num / 1000;
        return num; // ms
    }

    _getLatency(src, dst) {
        return this.latencyLookup.get(`${src}|${dst}`) ?? null;
    }

    _getColor(latency) {
        if (latency === null) return '#2a2a3e';
        if (latency < 40) return '#1a5c2a';
        if (latency < 100) return '#5c5c1a';
        if (latency < 200) return '#6c3a0a';
        if (latency < 300) return '#6c1a1a';
        return '#4c0a0a';
    }

    _getTextColor(latency) {
        if (latency === null) return '#555';
        if (latency < 40) return '#a6e3a1';
        if (latency < 100) return '#f9e2af';
        if (latency < 200) return '#fab387';
        if (latency < 300) return '#f38ba8';
        return '#e06080';
    }

    _regionName(id) {
        const r = typeof AZURE_REGIONS !== 'undefined' && AZURE_REGIONS[id];
        return r ? r.displayName : id;
    }

    _regionInfo(id) {
        const r = typeof AZURE_REGIONS !== 'undefined' && AZURE_REGIONS[id];
        if (!r) return null;
        return r;
    }

    _geoGroupName(key) {
        if (typeof GEO_GROUPS !== 'undefined' && GEO_GROUPS[key]) return GEO_GROUPS[key].displayName;
        return key;
    }

    _avgLatencyForRegion(regionId, peerList) {
        let sum = 0, count = 0;
        peerList.forEach(peerId => {
            if (regionId === peerId) return;
            const lat = this.latencyLookup.get(`${regionId}|${peerId}`);
            if (lat != null) { sum += lat; count++; }
        });
        return count > 0 ? sum / count : Infinity;
    }

    // ── Filtering & sorting ──────────────────────────────────

    _applyFilters() {
        let sources = [...this.sourceRegions];
        let dests = [...this.destRegions];

        // Geo filter
        if (this.geoFilter !== 'all') {
            const filterFn = id => {
                const r = this._regionInfo(id);
                return r && r.geoGroup === this.geoFilter;
            };
            sources = sources.filter(filterFn);
            dests = dests.filter(filterFn);
        }

        // Source multi-select
        if (this._sourceSelection && this._sourceSelection.size > 0 && this._sourceSelection.size < this.sourceRegions.length) {
            sources = sources.filter(s => this._sourceSelection.has(s));
        }
        // Dest multi-select
        if (this._destSelection && this._destSelection.size > 0 && this._destSelection.size < this.destRegions.length) {
            dests = dests.filter(d => this._destSelection.has(d));
        }

        // Sort — use filtered peers for avg calculation
        if (this.sortMode === 'latency') {
            sources.sort((a, b) => this._avgLatencyForRegion(a, dests) - this._avgLatencyForRegion(b, dests));
            dests.sort((a, b) => this._avgLatencyForRegion(a, sources) - this._avgLatencyForRegion(b, sources));
        } else {
            const nameCmp = (a, b) => this._regionName(a).localeCompare(this._regionName(b));
            sources.sort(nameCmp);
            dests.sort(nameCmp);
        }

        this.filteredSources = sources;
        this.filteredDests = dests;
    }

    // ── Rendering ────────────────────────────────────────────

    _render() {
        this._applyFilters();
        this.container.innerHTML = '';
        this.container.appendChild(this._buildToolbar());
        this.container.appendChild(this._buildMatrix());
    }

    _buildToolbar() {
        const bar = document.createElement('div');
        bar.className = 'matrix-toolbar';

        // Geography filter
        const geoLabel = this._el('span', 'Geography:');
        geoLabel.style.color = '#94a3b8';
        geoLabel.style.fontSize = '0.85rem';
        const geoSelect = document.createElement('select');
        geoSelect.className = 'matrix-select';
        geoSelect.innerHTML = '<option value="all">All Regions</option>';
        const groups = typeof GEO_GROUPS !== 'undefined' ? GEO_GROUPS : {};
        Object.keys(groups).sort().forEach(k => {
            geoSelect.innerHTML += `<option value="${k}">${groups[k].displayName}</option>`;
        });
        geoSelect.value = this.geoFilter;
        geoSelect.addEventListener('change', () => {
            this.geoFilter = geoSelect.value;
            this._syncURLParams();
            this._render();
        });

        // Source multi-select button
        const srcBtn = document.createElement('button');
        srcBtn.className = 'matrix-btn';
        this._updateMultiSelectBtn('source', srcBtn);
        srcBtn.addEventListener('click', () => this._showMultiSelect('source', srcBtn));

        // Dest multi-select button
        const dstBtn = document.createElement('button');
        dstBtn.className = 'matrix-btn';
        this._updateMultiSelectBtn('dest', dstBtn);
        dstBtn.addEventListener('click', () => this._showMultiSelect('dest', dstBtn));

        // Sort
        const sortLabel = this._el('span', 'Sort:');
        sortLabel.style.color = '#94a3b8';
        sortLabel.style.fontSize = '0.85rem';
        const sortSelect = document.createElement('select');
        sortSelect.className = 'matrix-select';
        sortSelect.innerHTML = '<option value="name">By Name</option><option value="latency">By Avg Latency</option>';
        sortSelect.value = this.sortMode;
        sortSelect.addEventListener('change', () => {
            this.sortMode = sortSelect.value;
            this._syncURLParams();
            this._render();
        });

        // Spacer
        const spacer = document.createElement('div');
        spacer.style.marginLeft = 'auto';

        // Region count
        const countLabel = this._el('span', `${this.filteredSources.length}×${this.filteredDests.length}`);
        countLabel.style.color = '#00d9ff';
        countLabel.style.fontSize = '0.8rem';

        // Legend
        const legend = document.createElement('div');
        legend.style.cssText = 'display:flex;gap:10px;align-items:center;font-size:0.75rem;';
        const bands = [
            { color: '#1a5c2a', text: '#a6e3a1', label: '< 40' },
            { color: '#5c5c1a', text: '#f9e2af', label: '40–99' },
            { color: '#6c3a0a', text: '#fab387', label: '100–199' },
            { color: '#6c1a1a', text: '#f38ba8', label: '200–299' },
            { color: '#4c0a0a', text: '#e06080', label: '≥ 300' },
        ];
        bands.forEach(b => {
            const item = document.createElement('span');
            item.style.cssText = `display:inline-flex;align-items:center;gap:3px;`;
            const swatch = document.createElement('span');
            swatch.style.cssText = `width:12px;height:12px;border-radius:2px;background:${b.color};display:inline-block;`;
            const lbl = document.createElement('span');
            lbl.style.color = '#94a3b8';
            lbl.textContent = b.label;
            item.appendChild(swatch);
            item.appendChild(lbl);
            legend.appendChild(item);
        });

        // CSV export — subtle style
        const csvBtn = document.createElement('button');
        csvBtn.className = 'matrix-btn';
        csvBtn.style.cssText = 'font-size:0.75rem;opacity:0.6;padding:4px 10px;';
        csvBtn.title = 'Download filtered view as CSV';
        csvBtn.innerHTML = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" style="vertical-align:-1px;margin-right:3px"><path d="M5.625 15C5.625 14.586 5.289 14.25 4.875 14.25C4.461 14.25 4.125 14.586 4.125 15H5.625ZM4.875 16H4.125H4.875ZM19.275 15C19.275 14.586 18.939 14.25 18.525 14.25C18.111 14.25 17.775 14.586 17.775 15H19.275ZM11.109 15.539C10.854 15.865 10.912 16.337 11.239 16.591C11.565 16.846 12.037 16.788 12.291 16.461L11.109 15.539ZM16.191 11.461C16.446 11.135 16.388 10.663 16.061 10.409C15.735 10.154 15.263 10.212 15.009 10.539L16.191 11.461ZM11.109 16.461C11.363 16.788 11.835 16.846 12.161 16.591C12.488 16.337 12.546 15.865 12.291 15.539L11.109 16.461ZM8.391 10.539C8.137 10.212 7.665 10.154 7.339 10.409C7.012 10.663 6.954 11.135 7.209 11.461L8.391 10.539ZM10.95 16C10.95 16.414 11.286 16.75 11.7 16.75C12.114 16.75 12.45 16.414 12.45 16H10.95ZM12.45 5C12.45 4.586 12.114 4.25 11.7 4.25C11.286 4.25 10.95 4.586 10.95 5H12.45ZM4.125 15V16H5.625V15H4.125ZM4.125 16C4.125 18.053 5.753 19.75 7.8 19.75V18.25C6.617 18.25 5.625 17.261 5.625 16H4.125ZM7.8 19.75H15.6V18.25H7.8V19.75ZM15.6 19.75C17.647 19.75 19.275 18.053 19.275 16H17.775C17.775 17.261 16.783 18.25 15.6 18.25V19.75ZM19.275 16V15H17.775V16H19.275ZM12.291 16.461L16.191 11.461L15.009 10.539L11.109 15.539L12.291 16.461ZM12.291 15.539L8.391 10.539L7.209 11.461L11.109 16.461L12.291 15.539ZM12.45 16V5H10.95V16H12.45Z" fill="#94a3b8"/></svg>CSV';
        csvBtn.addEventListener('click', () => this._exportCSV());

        [geoLabel, geoSelect, srcBtn, dstBtn, sortLabel, sortSelect, spacer, countLabel, legend, csvBtn].forEach(el => bar.appendChild(el));
        return bar;
    }

    _buildMatrix() {
        const wrapper = document.createElement('div');
        wrapper.className = 'matrix-scroll';

        if (this.filteredSources.length === 0 || this.filteredDests.length === 0) {
            wrapper.innerHTML = '<div style="padding:40px;text-align:center;color:#94a3b8;">No regions match the current filters.</div>';
            return wrapper;
        }

        const table = document.createElement('table');
        table.className = 'matrix-table';

        // Header row
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        const cornerCell = document.createElement('th');
        cornerCell.className = 'matrix-corner';
        cornerCell.textContent = 'Source \\ Dest';
        headerRow.appendChild(cornerCell);

        this.filteredDests.forEach(dst => {
            const th = document.createElement('th');
            th.className = 'matrix-col-header';
            const innerDiv = document.createElement('div');
            innerDiv.className = 'matrix-col-header-text';
            innerDiv.textContent = this._regionName(dst);
            th.appendChild(innerDiv);
            th.addEventListener('mouseenter', e => this._showRegionTooltip(dst, e));
            th.addEventListener('mouseleave', () => this._hideTooltip());
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        table.appendChild(thead);

        // Body rows
        const tbody = document.createElement('tbody');
        this.filteredSources.forEach(src => {
            const row = document.createElement('tr');
            const rowHeader = document.createElement('td');
            rowHeader.className = 'matrix-row-header';
            rowHeader.textContent = this._regionName(src);
            rowHeader.addEventListener('mouseenter', e => this._showRegionTooltip(src, e));
            rowHeader.addEventListener('mouseleave', () => this._hideTooltip());
            row.appendChild(rowHeader);

            this.filteredDests.forEach(dst => {
                const td = document.createElement('td');
                td.className = 'matrix-cell';
                const lat = this._getLatency(src, dst);
                if (src === dst) {
                    td.style.background = '#1a1a2e';
                    td.textContent = '—';
                    td.style.color = '#555';
                } else if (lat !== null) {
                    td.style.background = this._getColor(lat);
                    td.style.color = this._getTextColor(lat);
                    td.textContent = lat.toFixed(1);
                    td.addEventListener('mouseenter', e => this._showSparkline(src, dst, lat, e));
                    td.addEventListener('mouseleave', () => this._hideSparkline());
                } else {
                    td.style.background = '#2a2a3e';
                    td.textContent = '–';
                    td.style.color = '#555';
                }
                row.appendChild(td);
            });
            tbody.appendChild(row);
        });
        table.appendChild(tbody);
        wrapper.appendChild(table);
        return wrapper;
    }

    // ── Tooltips & sparklines ────────────────────────────────

    _showRegionTooltip(regionId, event) {
        const info = this._regionInfo(regionId);
        if (!info) return;
        if (!this._tooltipEl) {
            this._tooltipEl = document.createElement('div');
            this._tooltipEl.className = 'matrix-tooltip';
            document.body.appendChild(this._tooltipEl);
        }
        const az = info.hasAvailabilityZones ? '✅ Yes' : '❌ No';
        this._tooltipEl.innerHTML = `
            <strong>${info.displayName}</strong><br>
            📍 ${info.country}<br>
            🌍 ${this._geoGroupName(info.geoGroup)}<br>
            🏢 AZ Support: ${az}
        `;
        this._tooltipEl.style.display = 'block';
        this._positionFloating(this._tooltipEl, event);
    }

    _hideTooltip() {
        if (this._tooltipEl) this._tooltipEl.style.display = 'none';
    }

    _showSparkline(src, dst, latency, event) {
        if (!this._sparklineEl) {
            this._sparklineEl = document.createElement('div');
            this._sparklineEl.className = 'matrix-sparkline-popup';
            document.body.appendChild(this._sparklineEl);
        }

        const history = this.historyMap.get(`${src}|${dst}`) || [];
        const srcName = this._regionName(src);
        const dstName = this._regionName(dst);

        let html = `<div style="margin-bottom:6px;font-weight:600;font-size:0.85rem;">${srcName} → ${dstName}</div>`;
        html += `<div style="font-size:0.8rem;color:#00d9ff;margin-bottom:6px;">Current: ${latency.toFixed(2)} ms</div>`;

        if (history.length > 1) {
            html += `<canvas id="matrix-sparkline-canvas" width="200" height="60"></canvas>`;
            html += `<div style="font-size:0.7rem;color:#94a3b8;margin-top:4px;">${history.length} measurements</div>`;
        } else {
            html += `<div style="font-size:0.75rem;color:#94a3b8;">No history data</div>`;
        }

        this._sparklineEl.innerHTML = html;
        this._sparklineEl.style.display = 'block';
        this._positionFloating(this._sparklineEl, event);

        // Draw sparkline canvas
        if (history.length > 1) {
            requestAnimationFrame(() => {
                const canvas = document.getElementById('matrix-sparkline-canvas');
                if (canvas) this._drawSparkline(canvas, history);
            });
        }
    }

    _hideSparkline() {
        if (this._sparklineEl) this._sparklineEl.style.display = 'none';
    }

    _positionFloating(el, event) {
        const x = event.clientX + 12;
        const y = event.clientY + 12;
        el.style.left = `${Math.min(x, window.innerWidth - 260)}px`;
        el.style.top = `${Math.min(y, window.innerHeight - 150)}px`;
    }

    _drawSparkline(canvas, history) {
        const ctx = canvas.getContext('2d');
        const w = canvas.width;
        const h = canvas.height;
        const values = history.map(h => h.latency);
        const min = 0;
        const max = Math.max(...values);
        const range = max || 1;
        const pad = 4;

        ctx.clearRect(0, 0, w, h);

        // Line
        ctx.beginPath();
        ctx.strokeStyle = '#00d9ff';
        ctx.lineWidth = 1.5;
        values.forEach((v, i) => {
            const x = pad + (i / (values.length - 1)) * (w - 2 * pad);
            const y = h - pad - ((v - min) / range) * (h - 2 * pad);
            if (i === 0) ctx.moveTo(x, y);
            else ctx.lineTo(x, y);
        });
        ctx.stroke();

        // Fill under
        const lastX = pad + ((values.length - 1) / (values.length - 1)) * (w - 2 * pad);
        ctx.lineTo(lastX, h - pad);
        ctx.lineTo(pad, h - pad);
        ctx.closePath();
        ctx.fillStyle = 'rgba(0,217,255,0.1)';
        ctx.fill();

        // Min/max labels
        ctx.fillStyle = '#94a3b8';
        ctx.font = '9px sans-serif';
        ctx.fillText(`${min.toFixed(1)}`, 2, h - 2);
        ctx.fillText(`${max.toFixed(1)}`, 2, 10);
    }

    // ── Multi-select popover ─────────────────────────────────

    _showMultiSelect(type, anchorBtn) {
        // Remove existing popover
        const existing = document.querySelector('.matrix-multiselect-popover');
        if (existing) { existing.remove(); return; }

        const allRegions = type === 'source' ? this.sourceRegions : this.destRegions;
        // Initialize selection set if null (means "all")
        if (type === 'source' && !this._sourceSelection) {
            this._sourceSelection = new Set(allRegions);
        }
        if (type === 'dest' && !this._destSelection) {
            this._destSelection = new Set(allRegions);
        }
        const sel = type === 'source' ? this._sourceSelection : this._destSelection;

        const popover = document.createElement('div');
        popover.className = 'matrix-multiselect-popover';

        // Search
        const search = document.createElement('input');
        search.className = 'matrix-search';
        search.placeholder = 'Search regions...';
        popover.appendChild(search);

        // Select All / Clear
        const controls = document.createElement('div');
        controls.style.cssText = 'display:flex;gap:6px;margin-bottom:8px;';
        const selAllBtn = document.createElement('button');
        selAllBtn.className = 'matrix-btn';
        selAllBtn.textContent = 'All';
        selAllBtn.style.fontSize = '0.75rem';
        const clrBtn = document.createElement('button');
        clrBtn.className = 'matrix-btn';
        clrBtn.textContent = 'None';
        clrBtn.style.fontSize = '0.75rem';
        controls.appendChild(selAllBtn);
        controls.appendChild(clrBtn);
        popover.appendChild(controls);

        // Checkbox list
        const list = document.createElement('div');
        list.className = 'matrix-multiselect-list';

        const renderList = (filter) => {
            list.innerHTML = '';
            const filtered = filter
                ? allRegions.filter(r => this._regionName(r).toLowerCase().includes(filter.toLowerCase()))
                : allRegions;
            filtered.forEach(rId => {
                const item = document.createElement('label');
                item.className = 'matrix-ms-item';
                const cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.checked = sel.has(rId);
                cb.addEventListener('change', () => {
                    if (cb.checked) sel.add(rId);
                    else sel.delete(rId);
                    this._updateMultiSelectBtn(type, anchorBtn);
                });
                item.appendChild(cb);
                item.appendChild(document.createTextNode(' ' + this._regionName(rId)));
                list.appendChild(item);
            });
        };
        renderList('');

        search.addEventListener('input', () => renderList(search.value));
        selAllBtn.addEventListener('click', () => {
            allRegions.forEach(r => sel.add(r));
            this._updateMultiSelectBtn(type, anchorBtn);
            renderList(search.value);
        });
        clrBtn.addEventListener('click', () => {
            sel.clear();
            this._updateMultiSelectBtn(type, anchorBtn);
            renderList(search.value);
        });

        popover.appendChild(list);

        // Apply button
        const applyBtn = document.createElement('button');
        applyBtn.className = 'matrix-btn matrix-btn-accent';
        applyBtn.textContent = 'Apply';
        applyBtn.style.marginTop = '8px';
        applyBtn.style.width = '100%';
        applyBtn.addEventListener('click', () => {
            popover.remove();
            this._syncURLParams();
            this._render();
        });
        popover.appendChild(applyBtn);

        // Position near the button
        const rect = anchorBtn.getBoundingClientRect();
        popover.style.position = 'fixed';
        popover.style.left = `${rect.left}px`;
        popover.style.top = `${rect.bottom + 4}px`;
        popover.style.zIndex = '10000';
        document.body.appendChild(popover);

        // Close on outside click (apply + close)
        setTimeout(() => {
            const closeHandler = (e) => {
                if (!popover.contains(e.target) && e.target !== anchorBtn) {
                    popover.remove();
                    document.removeEventListener('mousedown', closeHandler);
                    this._syncURLParams();
                    this._render();
                }
            };
            document.addEventListener('mousedown', closeHandler);
        }, 0);
    }

    _updateMultiSelectBtn(type, btn) {
        const sel = type === 'source' ? this._sourceSelection : this._destSelection;
        const all = type === 'source' ? this.sourceRegions : this.destRegions;
        if (!sel || sel.size === all.length) {
            btn.textContent = `${type === 'source' ? 'Sources' : 'Destinations'} (all)`;
        } else {
            btn.textContent = `${type === 'source' ? 'Sources' : 'Destinations'} (${sel.size}/${all.length})`;
        }
    }

    // ── CSV export ───────────────────────────────────────────

    _exportCSV() {
        const rows = [];
        const header = ['Source \\ Destination', ...this.filteredDests.map(d => this._regionName(d))];
        rows.push(header.join(','));

        this.filteredSources.forEach(src => {
            const row = [this._regionName(src)];
            this.filteredDests.forEach(dst => {
                if (src === dst) row.push('');
                else {
                    const lat = this._getLatency(src, dst);
                    row.push(lat !== null ? lat.toFixed(2) : '');
                }
            });
            rows.push(row.join(','));
        });

        const blob = new Blob([rows.join('\n')], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'azure-latency-matrix.csv';
        a.click();
        URL.revokeObjectURL(url);
    }

    // ── Helpers ──────────────────────────────────────────────

    _el(tag, text) {
        const el = document.createElement(tag);
        if (text) el.textContent = text;
        return el;
    }

    // ── Styles (injected once) ───────────────────────────────

    _injectStyles() {
        if (this._styleInjected) return;
        this._styleInjected = true;

        const style = document.createElement('style');
        style.textContent = `
            #matrix {
                flex: 1;
                display: flex;
                flex-direction: column;
                background: #1a1a2e;
                overflow: hidden;
            }

            .matrix-toolbar {
                display: flex;
                align-items: center;
                gap: 10px;
                padding: 10px 16px;
                background: #16213e;
                border-bottom: 1px solid #0f3460;
                flex-wrap: wrap;
                flex-shrink: 0;
            }

            .matrix-select {
                padding: 6px 10px;
                border: 1px solid #0f3460;
                border-radius: 6px;
                background: #1a1a2e;
                color: #fff;
                font-size: 0.85rem;
                cursor: pointer;
            }
            .matrix-select:focus { outline: none; border-color: #0078d4; }

            .matrix-btn {
                padding: 6px 12px;
                border: 1px solid #0f3460;
                border-radius: 6px;
                background: #1a1a2e;
                color: #e2e8f0;
                font-size: 0.8rem;
                cursor: pointer;
                transition: all 0.2s;
            }
            .matrix-btn:hover {
                background: #0f3460;
                border-color: #0078d4;
            }
            .matrix-btn-accent {
                background: linear-gradient(135deg, #0078d4 0%, #106ebe 100%);
                border-color: #0078d4;
                color: #fff;
                font-weight: 600;
            }
            .matrix-btn-accent:hover {
                background: linear-gradient(135deg, #106ebe 0%, #0078d4 100%);
            }

            .matrix-scroll {
                flex: 1;
                overflow: auto;
                position: relative;
            }

            .matrix-table {
                border-collapse: separate;
                border-spacing: 0;
                font-size: 0.75rem;
                min-width: max-content;
            }

            .matrix-corner {
                position: sticky;
                top: 0;
                left: 0;
                z-index: 3;
                background: #16213e;
                padding: 8px 12px;
                font-size: 0.75rem;
                color: #94a3b8;
                border-bottom: 2px solid #0f3460;
                border-right: 2px solid #0f3460;
                min-width: 130px;
                text-align: left;
            }

            .matrix-col-header {
                position: sticky;
                top: 0;
                z-index: 2;
                background: #16213e;
                border-bottom: 2px solid #0f3460;
                padding: 0;
                height: 120px;
                min-width: 44px;
                max-width: 44px;
                vertical-align: bottom;
                cursor: default;
            }

            .matrix-col-header-text {
                writing-mode: vertical-rl;
                transform: rotate(180deg);
                padding: 8px 4px;
                white-space: nowrap;
                color: #e2e8f0;
                font-weight: 600;
                font-size: 0.72rem;
            }

            .matrix-row-header {
                position: sticky;
                left: 0;
                z-index: 1;
                background: #16213e;
                padding: 6px 12px;
                font-weight: 600;
                color: #e2e8f0;
                border-right: 2px solid #0f3460;
                white-space: nowrap;
                min-width: 130px;
                font-size: 0.78rem;
                cursor: default;
            }

            .matrix-cell {
                padding: 4px 2px;
                text-align: center;
                font-size: 0.7rem;
                font-weight: 600;
                cursor: default;
                min-width: 44px;
                max-width: 44px;
                border: 1px solid rgba(0,0,0,0.2);
                transition: opacity 0.15s;
            }
            .matrix-cell:hover {
                opacity: 0.8;
                outline: 2px solid #00d9ff;
                outline-offset: -2px;
            }

            .matrix-tooltip {
                position: fixed;
                background: #16213e;
                border: 1px solid #0f3460;
                border-radius: 8px;
                padding: 10px 14px;
                color: #e2e8f0;
                font-size: 0.82rem;
                line-height: 1.6;
                z-index: 10001;
                pointer-events: none;
                box-shadow: 0 4px 20px rgba(0,0,0,0.5);
                display: none;
                max-width: 240px;
            }

            .matrix-sparkline-popup {
                position: fixed;
                background: #16213e;
                border: 1px solid #0f3460;
                border-radius: 8px;
                padding: 10px 14px;
                color: #e2e8f0;
                z-index: 10001;
                pointer-events: none;
                box-shadow: 0 4px 20px rgba(0,0,0,0.5);
                display: none;
                min-width: 220px;
            }

            .matrix-multiselect-popover {
                background: #16213e;
                border: 1px solid #0f3460;
                border-radius: 8px;
                padding: 12px;
                box-shadow: 0 8px 30px rgba(0,0,0,0.6);
                width: 260px;
                max-height: 400px;
                display: flex;
                flex-direction: column;
            }

            .matrix-search {
                padding: 6px 10px;
                border: 1px solid #0f3460;
                border-radius: 6px;
                background: #1a1a2e;
                color: #fff;
                font-size: 0.85rem;
                margin-bottom: 8px;
                width: 100%;
            }
            .matrix-search:focus { outline: none; border-color: #0078d4; }
            .matrix-search::placeholder { color: #64748b; }

            .matrix-multiselect-list {
                overflow-y: auto;
                max-height: 260px;
                display: flex;
                flex-direction: column;
                gap: 2px;
            }

            .matrix-ms-item {
                display: flex;
                align-items: center;
                gap: 8px;
                padding: 5px 8px;
                border-radius: 4px;
                font-size: 0.82rem;
                color: #e2e8f0;
                cursor: pointer;
            }
            .matrix-ms-item:hover { background: rgba(0,120,212,0.15); }
            .matrix-ms-item input { accent-color: #0078d4; cursor: pointer; }

            @media (max-width: 768px) {
                .matrix-toolbar {
                    padding: 8px 10px;
                    gap: 6px;
                }
                .matrix-col-header { height: 80px; min-width: 36px; max-width: 36px; }
                .matrix-cell { min-width: 36px; max-width: 36px; font-size: 0.65rem; }
                .matrix-row-header { min-width: 100px; font-size: 0.72rem; padding: 4px 8px; }
                .matrix-corner { min-width: 100px; }
            }
        `;
        document.head.appendChild(style);
    }
}
