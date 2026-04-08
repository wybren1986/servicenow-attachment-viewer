import { createCustomElement, actionTypes } from '@servicenow/ui-core';
const { COMPONENT_BOOTSTRAPPED, COMPONENT_PROPERTY_CHANGED } = actionTypes;
import snabbdom from '@servicenow/ui-renderer-snabbdom';
import styles from './styles.scss';
import '@servicenow/now-button';
import '@servicenow/now-icon';
import '@servicenow/now-loader';
import '@servicenow/now-message';
import '@servicenow/now-illustration';
import { renderAsync } from 'docx-preview';
import * as XLSX from 'xlsx';
import MsgReaderDefault from 'msgreader/lib/MsgReader';
const MsgReader = MsgReaderDefault.default || MsgReaderDefault;

const AV_VERSION = '2.0.0';
const AV_BUILD = '2026-04-08';

// ─── Helpers ────────────────────────────────────────────────────────────────

const formatDate = (dateStr) => {
	if (!dateStr) return '';
	const d = new Date(dateStr.replace(' ', 'T'));
	if (isNaN(d)) return '';
	return d.toLocaleDateString('nl-NL', { day: '2-digit', month: '2-digit', year: 'numeric' });
};

const formatSize = (bytes) => {
	if (!bytes) return '';
	if (bytes < 1024) return bytes + ' B';
	if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
	return (bytes / 1048576).toFixed(1) + ' MB';
};

const IMAGE_TYPES = ['jpg', 'jpeg', 'png', 'gif', 'svg', 'webp', 'bmp'];
const TEXT_TYPES  = ['txt', 'log', 'json', 'xml', 'csv', 'js', 'ts', 'html', 'css'];
const MAX_FILE_SIZE = 24 * 1024 * 1024; // 24MB ServiceNow default limit

const getCategory = (ext) => {
	if (ext === 'pdf')  return 'pdf';
	if (ext === 'msg')  return 'msg';
	if (['xls', 'xlsx'].includes(ext)) return 'excel';
	if (['ppt', 'pptx'].includes(ext)) return 'ppt';
	if (['doc', 'docx'].includes(ext)) return 'word';
	if (IMAGE_TYPES.includes(ext))     return 'image';
	if (TEXT_TYPES.includes(ext))      return 'text';
	return 'unsupported';
};

const fileIcon = (ext) => {
	const cat = getCategory(ext);
	const icons = {
		pdf: 'document-pdf-outline',
		excel: 'document-excel-outline',
		word: 'document-outline',
		msg: 'envelope-outline',
		image: 'document-image-outline',
		text: 'document-code-outline',
		ppt: 'document-powerpoint-outline',
		unsupported: 'document-blank-outline'
	};
	const colors = {
		pdf: '#e53e3e',
		excel: '#38a169',
		word: '#3182ce',
		msg: '#d69e2e',
		image: '#805ad5',
		text: '#718096',
		ppt: '#e05d2c',
		unsupported: '#a0aec0'
	};
	const icon = icons[cat] || 'document-blank-outline';
	const color = colors[cat] || '#a0aec0';
	return <now-icon icon={icon} size="lg" style={{ color }} />;
};

const downloadUrl = (sys_id) => `/api/now/attachment/${sys_id}/file`;

// ─── MSG Preview ────────────────────────────────────────────────────────────

const buildMsgHtml = (MsgReader, buf, rawBuf) => {
	const reader = new MsgReader(buf);
	const d = reader.getFileData();
	const from = d.senderName ? `${d.senderName} &lt;${d.senderEmail || ''}&gt;` : (d.senderEmail || '—');
	const to = (d.recipients || []).map(r => r.name || r.email || '').filter(Boolean).join('; ') || '—';

	const msgAttachments = (d.attachments || []).map(att => {
		try {
			const attData = reader.getAttachment(att);
			const ext = (att.extension || (att.fileName || '').split('.').pop() || '').toLowerCase().replace('.', '');
			const isImage = IMAGE_TYPES.includes(ext);
			if (attData && attData.content && attData.content.length) {
				const mime = isImage ? `image/${ext === 'jpg' ? 'jpeg' : ext}` : 'application/octet-stream';
				const blob = new Blob([attData.content], { type: mime });
				return { name: att.fileName || att.fileNameShort || 'bijlage', isImage, url: URL.createObjectURL(blob) };
			}
		} catch (_) {}
		return { name: att.fileName || att.fileNameShort || 'bijlage', isImage: false, url: null };
	});

	const imageHtml = msgAttachments
		.filter(a => a.isImage && a.url)
		.map(a => `<div style="margin-bottom:16px">
			<img src="${a.url}" alt="${a.name}" style="max-width:100%;border-radius:4px;border:1px solid #e2e8f0;display:block" />
			<div style="font-size:11px;color:#718096;margin-top:4px">${a.name}</div>
		</div>`).join('');

	const fileListHtml = msgAttachments
		.filter(a => !a.isImage)
		.map(a => `<li style="padding:6px 10px;background:#f7fafc;border-radius:4px;margin-bottom:4px;font-size:13px">${a.name}</li>`)
		.join('');

	const attSection = (imageHtml || fileListHtml) ? `
		<div style="margin-top:24px;padding-top:16px;border-top:1px solid #e2e8f0">
			<strong style="font-size:12px;color:#718096;text-transform:uppercase;letter-spacing:.5px">Bijlagen (${msgAttachments.length})</strong>
			${imageHtml ? `<div style="margin-top:12px">${imageHtml}</div>` : ''}
			${fileListHtml ? `<ul style="list-style:none;padding:0;margin:8px 0 0">${fileListHtml}</ul>` : ''}
		</div>` : '';

	let bodyHtml;
	if (d.body) {
		bodyHtml = `<pre style="white-space:pre-wrap;font-family:inherit;font-size:14px;margin:0;line-height:1.6">${d.body}</pre>`;
	} else {
		let rawStr;
		try {
			rawStr = new TextDecoder('windows-1252').decode(new Uint8Array(rawBuf));
		} catch (_) {
			rawStr = new TextDecoder('utf-8', { fatal: false }).decode(new Uint8Array(rawBuf));
		}
		const htmlStart = rawStr.indexOf('<html');
		const htmlEnd = rawStr.indexOf('</html>');
		if (htmlStart > -1 && htmlEnd > -1) {
			bodyHtml = rawStr.substring(rawStr.indexOf('<body', htmlStart), htmlEnd + 7)
				.replace(/<\/?body[^>]*>/gi, '');
		} else {
			bodyHtml = '<em style="color:#718096">Geen inhoud</em>';
		}
	}

	return `
		<div class="av-msg-header">
			<div><b>Van:</b> ${from}</div>
			<div><b>Aan:</b> ${to}</div>
			<div><b>Onderwerp:</b> ${d.subject || '—'}</div>
		</div>
		<div class="av-msg-body">${bodyHtml}</div>
		${attSection}
	`;
};

// ─── Excel Preview ──────────────────────────────────────────────────────────

const buildExcelSheets = (buf) => {
	const wb = XLSX.read(buf, { type: 'array', cellStyles: true });
	const fonts = (wb.Styles && wb.Styles.Fonts) || [];
	const fills = (wb.Styles && wb.Styles.Fills) || [];
	const cellXf = (wb.Styles && wb.Styles.CellXf) || [];

	// Build fill-to-font lookup
	const fillFontMap = {};
	cellXf.forEach(xf => {
		const fillId = parseInt(xf.fillId);
		const fontId = parseInt(xf.fontId);
		if (!isNaN(fillId) && !isNaN(fontId) && fonts[fontId]) {
			if (!fillFontMap[fillId]) fillFontMap[fillId] = {};
			const fk = JSON.stringify(fonts[fontId]);
			fillFontMap[fillId][fk] = (fillFontMap[fillId][fk] || 0) + 1;
		}
	});
	const fillFont = {};
	Object.keys(fillFontMap).forEach(fid => {
		let best = null, bestCount = 0;
		Object.entries(fillFontMap[fid]).forEach(([fk, count]) => {
			if (count > bestCount) { best = JSON.parse(fk); bestCount = count; }
		});
		fillFont[fid] = best;
	});

	return wb.SheetNames.map(name => {
		const ws = wb.Sheets[name];
		const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
		let lastRow = 0, lastCol = 0;
		data.forEach((row, r) => {
			if (Array.isArray(row)) {
				row.forEach((cell, c) => {
					if (cell !== null && cell !== undefined && String(cell).trim() !== '') {
						if (r + 1 > lastRow) lastRow = r + 1;
						if (c + 1 > lastCol) lastCol = c + 1;
					}
				});
			}
		});
		if (!lastRow || !lastCol) return { name, html: '<table></table>' };

		const range = { s: { r: 0, c: 0 }, e: { r: lastRow - 1, c: lastCol - 1 } };
		let html = '<table>';
		for (let r = range.s.r; r <= range.e.r; r++) {
			html += '<tr>';
			for (let c = range.s.c; c <= range.e.c; c++) {
				const cell = ws[XLSX.utils.encode_cell({ r, c })];
				const val = cell ? (cell.w || String(cell.v || '')) : '';
				let style = '';
				if (cell && cell.s) {
					const s = cell.s;
					const fg = (s.fill && s.fill.fgColor) || s.fgColor;
					if (fg && fg.rgb && fg.theme !== 0) {
						style += 'background-color:#' + fg.rgb.slice(-6) + ';';
					}
					const matchFillIdx = fills.findIndex(f => {
						const ffg = f.fgColor;
						if (!ffg || !fg) return false;
						if (ffg.rgb && fg.rgb) return ffg.rgb === fg.rgb;
						return ffg.theme === fg.theme && ffg.tint === fg.tint;
					});
					const font = s.font || (matchFillIdx >= 0 && fillFont[matchFillIdx]) || null;
					if (font) {
						if (font.bold) style += 'font-weight:700;';
						if (font.italic) style += 'font-style:italic;';
						if (font.color && font.color.rgb) {
							const rgb = font.color.rgb.slice(-6);
							if (rgb !== '000000') style += 'color:#' + rgb + ';';
						}
						if (font.sz && font.sz !== 11) style += 'font-size:' + font.sz + 'pt;';
					}
					if (s.alignment) {
						if (s.alignment.horizontal) style += 'text-align:' + s.alignment.horizontal + ';';
						if (s.alignment.vertical) style += 'vertical-align:' + s.alignment.vertical + ';';
					}
				}
				html += '<td' + (style ? ' style="' + style + '"' : '') + '>' + val + '</td>';
			}
			html += '</tr>';
		}
		html += '</table>';
		return { name, html };
	});
};

// ─── Fetch Preview ──────────────────────────────────────────────────────────

const PREVIEW_CATS = ['pdf', 'word', 'excel', 'msg', 'text'];

const fetchPreview = (attachment, updateState) => {
	const { sys_id, file_type } = attachment;
	const cat = getCategory(file_type);
	if (!PREVIEW_CATS.includes(cat)) return;

	updateState({ blobUrl: null, previewText: null, previewHtml: null, previewSheets: null, docxData: null, activeSheet: 0, previewLoading: true });

	if (cat === 'text') {
		fetch(downloadUrl(sys_id), { credentials: 'same-origin' })
			.then(res => res.text())
			.then(text => updateState({ previewText: text, previewLoading: false }))
			.catch(() => updateState({ previewLoading: false }));
		return;
	}

	if (cat === 'pdf') {
		fetch(downloadUrl(sys_id), { credentials: 'same-origin' })
			.then(res => res.blob())
			.then(blob => updateState({ blobUrl: URL.createObjectURL(blob), previewLoading: false }))
			.catch(() => updateState({ previewLoading: false }));
		return;
	}

	fetch(downloadUrl(sys_id), { credentials: 'same-origin' })
		.then(res => res.arrayBuffer())
		.then(buf => {
			if (cat === 'word') {
				updateState({ docxData: buf, previewLoading: false });
				return;
			}
			if (cat === 'excel') {
				const sheets = buildExcelSheets(buf);
				updateState({ previewSheets: sheets, activeSheet: 0, previewLoading: false });
			}
			if (cat === 'msg') {
				const html = buildMsgHtml(MsgReader, buf, buf);
				updateState({ previewHtml: html, previewLoading: false });
			}
		})
		.catch(() => updateState({ previewLoading: false }));
};

// ─── Upload ─────────────────────────────────────────────────────────────────

const uploadFiles = (files, table, sysid, updateState) => {
	if (!files || !files.length || !sysid) return;

	const fileList = Array.from(files);
	const tooLarge = fileList.filter(f => f.size > MAX_FILE_SIZE);
	if (tooLarge.length) {
		const names = tooLarge.map(f => `${f.name} (${formatSize(f.size)})`).join(', ');
		updateState({ uploadError: `Bestand(en) te groot (max ${formatSize(MAX_FILE_SIZE)}): ${names}` });
		return;
	}

	updateState({ uploading: true, uploadError: null });
	const uploads = fileList.map(file => {
		const url = `/api/now/attachment/file?table_name=${table}&table_sys_id=${sysid}&file_name=${encodeURIComponent(file.name)}`;
		return fetch(url, {
			method: 'POST',
			credentials: 'same-origin',
			headers: { 'Content-Type': file.type || 'application/octet-stream' },
			body: file
		}).then(res => {
			if (!res.ok) throw new Error(`Upload mislukt voor ${file.name}: ${res.status}`);
			return res;
		});
	});
	Promise.all(uploads)
		.then(() => {
			updateState({ uploading: false });
			fetchAttachments(table, sysid, updateState);
		})
		.catch(err => {
			updateState({ uploading: false, uploadError: err.message });
			fetchAttachments(table, sysid, updateState);
		});
};

// ─── Fetch Attachments ──────────────────────────────────────────────────────

const fetchAttachments = (table, sysid, updateState) => {
	const query = `table_name=${table}^table_sys_id=${sysid}`;
	const url = `/api/now/attachment?sysparm_query=${encodeURIComponent(query)}&sysparm_limit=200`;
	updateState({ loading: true });

	fetch(url, { credentials: 'same-origin', headers: { Accept: 'application/json' } })
		.then(res => res.json())
		.then(data => {
			const result = Array.isArray(data?.result) ? data.result : [];
			const attachments = result.map(item => ({
				sys_id: item.sys_id,
				file_name: item.file_name,
				size_bytes: item.size_bytes,
				content_type: item.content_type,
				file_type: (item.file_name || '').split('.').pop().toLowerCase(),
				sys_created_on: item.sys_created_on
			}));
			const first = attachments.length ? attachments[0] : null;
			updateState({ attachments, selectedId: first ? first.sys_id : null, loading: false, blobUrl: null, previewText: null, previewHtml: null, previewSheets: null, docxData: null, activeSheet: 0 });
			if (first) fetchPreview(first, updateState);
		})
		.catch(() => updateState({ loading: false }));
};

// ─── Render Preview ─────────────────────────────────────────────────────────

const renderPreview = (attachment, state, dispatch) => {
	if (!attachment) {
		return (
			<div className="av-preview-state">
				<p>Selecteer een bijlage om te bekijken</p>
			</div>
		);
	}

	const { sys_id, file_name, file_type } = attachment;
	const cat = getCategory(file_type);
	const url = downloadUrl(sys_id);
	const { blobUrl, previewText, previewHtml, previewSheets, previewLoading } = state;
	const activeSheet = state.activeSheet || 0;

	if (previewLoading && PREVIEW_CATS.includes(cat)) {
		return (
			<div className="av-empty">
				<now-loader label="Loading" />
			</div>
		);
	}

	if (cat === 'image') {
		return (
			<div className="av-preview -image">
				<img className="av-preview-img" src={url} alt={file_name} />
			</div>
		);
	}

	if (cat === 'pdf' && blobUrl) {
		return (
			<div className="av-preview -pdf">
				<object data={blobUrl} type="application/pdf" aria-label={file_name} />
			</div>
		);
	}

	if (cat === 'word' && state.docxData) {
		return (
			<div key={'docx-' + sys_id} className="av-preview -docx"
				hook-insert={(vnode) => {
					renderAsync(state.docxData, vnode.elm, null, {
						className: 'av-docx-container',
						inWrapper: true,
						ignoreWidth: false,
						ignoreHeight: false,
						renderHeaders: true,
						renderFooters: true,
						renderFootnotes: true
					}).catch(() => {});
				}}
			/>
		);
	}

	if (cat === 'excel' && previewSheets && previewSheets.length) {
		const sheet = previewSheets[activeSheet];
		return (
			<div key={'excel-' + sys_id} className="av-preview -excel">
				{previewSheets.length > 1 && (
					<div className="av-excel-tabs">
						{previewSheets.map((s, i) => (
							<button
								key={i}
								className={'av-excel-tab-btn' + (i === activeSheet ? ' -active' : '')}
								onclick={() => dispatch(() => ({ type: 'SET_ACTIVE_SHEET', payload: { index: i } }))}
							>
								{s.name}
							</button>
						))}
					</div>
				)}
				<div className="av-excel-sheet" hook-insert={(vnode) => { vnode.elm.innerHTML = sheet.html; }} hook-update={(_, vnode) => { vnode.elm.innerHTML = sheet.html; }} />
			</div>
		);
	}

	if (cat === 'msg' && previewHtml) {
		return (
			<div key={'msg-' + sys_id} className="av-preview -msg" hook-insert={(vnode) => { vnode.elm.innerHTML = previewHtml; }} />
		);
	}

	if (cat === 'text' && previewText !== null) {
		return (
			<div className="av-preview -text">
				<pre className="av-preview-text">{previewText}</pre>
			</div>
		);
	}

	return (
		<div className="av-preview -unsupported">
			<div className="av-unsupported">
				{fileIcon(file_type)}
				<h3>{file_name}</h3>
				<p>Preview niet beschikbaar voor dit bestandstype</p>
				<a href={url} download={file_name} className="av-download-link">
					<now-button label="Downloaden" variant="primary" size="md" icon="download-outline" />
				</a>
			</div>
		</div>
	);
};

// ─── View ───────────────────────────────────────────────────────────────────

const view = (state, { dispatch }) => {
	const attachments = state.attachments || [];
	const selectedId = state.selectedId || null;
	const loading = state.loading || false;
	const selected = attachments.find(a => a.sys_id === selectedId) || null;

	if (loading) {
		return (
			<div className="av-empty">
				<now-loader label="Loading" />
			</div>
		);
	}

	if (!attachments.length) {
		return (
			<div className="av-empty">
				<now-message alignment="vertical-centered">
					<now-illustration slot="media" illustration="add-attachment" size="auto" />
					<div slot="message">
						<h3 className="now-heading--md now-m-block--0">No attachments available</h3>
						<p className="now-m-block-start--sm now-m-block-end--0 now-color_text--tertiary">Drag or select files to upload</p>
					</div>
					<div slot="actions">
						<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'OPEN_FILE_PICKER' }))}>
							<now-button label="Select file" variant="primary" size="md" />
						</span>
					</div>
				</now-message>
			</div>
		);
	}

	const dragging = state.dragging || false;
	const uploading = state.uploading || false;
	const uploadError = state.uploadError || null;
	const confirmDelete = state.confirmDelete || null;

	return (
		<div className="av-root">
			{confirmDelete && (
				<div className="av-drop-overlay" style={{ pointerEvents: 'auto' }}>
					<div className="av-confirm-card">
						<now-message alignment="vertical-centered">
							<now-icon slot="media" icon="trash-outline" size="lg" />
							<div slot="message">
								<h3 className="now-heading--md now-m-block--0">Bestand verwijderen</h3>
								<p className="now-m-block-start--sm now-m-block-end--0">Weet je zeker dat je <strong>{confirmDelete.file_name}</strong> wilt verwijderen?</p>
							</div>
							<div slot="actions" className="av-confirm-actions">
								<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'DELETE_ATTACHMENT', payload: { sys_id: confirmDelete.sys_id } }))}>
									<now-button label="Verwijderen" variant="primary-negative" size="md" />
								</span>
								<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'DISMISS_DELETE' }))}>
									<now-button label="Annuleren" variant="secondary" size="md" />
								</span>
							</div>
						</now-message>
					</div>
				</div>
			)}

			{uploadError && (
				<div className="av-drop-overlay" style={{ pointerEvents: 'auto' }}>
					<div className="av-confirm-card">
						<now-message alignment="vertical-centered">
							<now-icon slot="media" icon="circle-exclamation-outline" size="lg" style={{ color: '#e53e3e' }} />
							<div slot="message">
								<h3 className="now-heading--md now-m-block--0">Upload mislukt</h3>
								<p className="now-m-block-start--sm now-m-block-end--0">{uploadError}</p>
							</div>
							<div slot="actions">
								<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'DISMISS_ERROR' }))}>
									<now-button label="Sluiten" variant="secondary" size="md" />
								</span>
							</div>
						</now-message>
					</div>
				</div>
			)}

			{dragging && (
				<div className="av-drop-overlay">
					<div className="av-drop-message">
						<now-icon icon="upload-outline" size="lg" />
						<p>Drop bestanden om te uploaden</p>
					</div>
				</div>
			)}

			{uploading && (
				<div className="av-drop-overlay">
					<div className="av-upload-card">
						<now-loader label="Uploading..." action="Cancel" />
					</div>
				</div>
			)}

			{state.deleting && (
				<div className="av-drop-overlay">
					<div className="av-upload-card">
						<now-loader label="Verwijderen..." />
					</div>
				</div>
			)}

			<div className="av-sidebar">
				<ul className="av-list">
					{attachments.map(a => (
						<li
							key={a.sys_id}
							className={'av-list-item' + (a.sys_id === selectedId ? ' -active' : '')}
							onclick={() => dispatch(() => ({ type: 'SELECT_ATTACHMENT', payload: { sys_id: a.sys_id } }))}
						>
							<div className="av-file-icon">{fileIcon(a.file_type)}</div>
							<div className="av-list-item-info">
								<div className="av-list-item-name">{a.file_name}</div>
								<div className="av-list-item-meta">{formatSize(a.size_bytes)}{a.sys_created_on ? ' · ' + formatDate(a.sys_created_on) : ''}</div>
							</div>
						</li>
					))}
				</ul>
				<div className="av-sidebar-actions">
					<div className="av-upload-wrap" onclick={() => dispatch(() => ({ type: 'OPEN_FILE_PICKER' }))}>
						<now-button label="Upload" variant="primary" size="md" icon="upload-outline" />
					</div>
					<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'FETCH_ATTACHMENTS' }))}>
						<now-button-iconic icon="sync-outline" variant="tertiary" size="md" tooltipContent="Vernieuwen" />
					</span>
				</div>
			</div>

			<div className="av-main">
				{selected && (
					<div className="av-header">
						<div className="av-header-title">
							<div className="av-file-icon">{fileIcon(selected.file_type)}</div>
							<div>
								<div className="av-header-filename">{selected.file_name}</div>
								<div className="av-header-meta">{formatSize(selected.size_bytes)}</div>
							</div>
						</div>
						<div className="av-header-actions">
							<a href={downloadUrl(selected.sys_id)} download={selected.file_name} className="av-download-link">
								<now-button-iconic icon="download-outline" variant="tertiary" size="md" tooltipContent="Download" />
							</a>
							<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'CONFIRM_DELETE', payload: { sys_id: selected.sys_id, file_name: selected.file_name } }))}>
								<now-button-iconic icon="trash-outline" variant="tertiary" size="md" tooltipContent="Verwijderen" />
							</span>
						</div>
					</div>
				)}
				<div className="av-content">
					{renderPreview(selected, state, dispatch)}
				</div>
			</div>
		</div>
	);
};

// ─── Action Handlers ────────────────────────────────────────────────────────

const actionHandlers = {
	[COMPONENT_BOOTSTRAPPED]: ({ host, dispatch, updateState, properties }) => {
		if (properties.sysid) {
			fetchAttachments(properties.table, properties.sysid, updateState);
		}
		// Auto-size height to available space
		const setHeight = () => {
			const rect = host.getBoundingClientRect();
			const available = window.innerHeight - rect.top;
			host.style.height = available + 'px';
			host.style.maxHeight = available + 'px';
			host.style.overflow = 'hidden';
		};
		setTimeout(setHeight, 100);
		window.addEventListener('resize', setHeight);
		// Drag & drop
		host.addEventListener('dragenter', (e) => {
			e.preventDefault();
			dispatch(() => ({ type: 'SET_DRAGGING', payload: true }));
		});
		host.addEventListener('dragover', (e) => { e.preventDefault(); });
		host.addEventListener('dragleave', (e) => {
			if (!host.contains(e.relatedTarget)) {
				dispatch(() => ({ type: 'SET_DRAGGING', payload: false }));
			}
		});
		host.addEventListener('drop', (e) => {
			e.preventDefault();
			dispatch(() => ({ type: 'SET_DRAGGING', payload: false }));
			dispatch(() => ({ type: 'UPLOAD_FILES', payload: { files: e.dataTransfer.files } }));
		});
	},

	[COMPONENT_PROPERTY_CHANGED]: ({ action, updateState, properties }) => {
		const { name } = action.payload || {};
		if ((name === 'sysid' || name === 'table') && properties.sysid) {
			fetchAttachments(properties.table, properties.sysid, updateState);
		}
	},

	'FETCH_ATTACHMENTS': ({ updateState, properties }) => {
		fetchAttachments(properties.table, properties.sysid, updateState);
	},

	'SELECT_ATTACHMENT': ({ action, updateState, state }) => {
		const sys_id = action.payload.sys_id;
		updateState({ selectedId: sys_id, blobUrl: null, previewText: null, previewHtml: null, previewSheets: null, docxData: null, activeSheet: 0 });
		const attachment = (state.attachments || []).find(a => a.sys_id === sys_id);
		if (attachment) fetchPreview(attachment, updateState);
	},

	'OPEN_FILE_PICKER': ({ updateState, properties }) => {
		const input = document.createElement('input');
		input.type = 'file';
		input.multiple = true;
		input.onchange = () => {
			if (input.files.length) {
				uploadFiles(input.files, properties.table, properties.sysid, updateState);
			}
		};
		input.click();
	},

	'UPLOAD_FILES': ({ action, updateState, properties }) => {
		uploadFiles(action.payload.files, properties.table, properties.sysid, updateState);
	},

	'SET_DRAGGING': ({ action, updateState }) => {
		updateState({ dragging: action.payload });
	},

	'SET_ACTIVE_SHEET': ({ action, updateState }) => {
		updateState({ activeSheet: action.payload.index });
	},

	'NOW_LOADER#ACTION_CLICKED': ({ updateState }) => {
		updateState({ uploading: false });
	},

	'DISMISS_ERROR': ({ updateState }) => {
		updateState({ uploadError: null });
	},

	'CONFIRM_DELETE': ({ action, updateState }) => {
		updateState({ confirmDelete: action.payload });
	},

	'DISMISS_DELETE': ({ updateState }) => {
		updateState({ confirmDelete: null });
	},

	'DELETE_ATTACHMENT': ({ action, updateState, properties }) => {
		const { sys_id } = action.payload;
		updateState({ confirmDelete: null, deleting: true });
		fetch(`/api/now/attachment/${sys_id}`, {
			method: 'DELETE',
			credentials: 'same-origin',
			headers: { Accept: 'application/json' }
		})
			.then(res => {
				if (!res.ok) throw new Error('Verwijderen mislukt: ' + res.status);
				updateState({ deleting: false });
				fetchAttachments(properties.table, properties.sysid, updateState);
			})
			.catch(err => {
				updateState({ deleting: false, uploadError: err.message });
			});
	},
};

// ─── Register ───────────────────────────────────────────────────────────────

createCustomElement('sofbv-attachment-viewer', {
	renderer: { type: snabbdom },
	view,
	styles,
	actionHandlers,
	properties: {
		table: { default: 'incident', schema: { type: 'string' } },
		sysid: { default: '', schema: { type: 'string' } }
	}
});
