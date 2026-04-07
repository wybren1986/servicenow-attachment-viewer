const AV_VERSION = '1.0.0';
const AV_BUILD = '2026-04-07';

import { createCustomElement, actionTypes } from '@servicenow/ui-core';
const { COMPONENT_BOOTSTRAPPED, COMPONENT_PROPERTY_CHANGED } = actionTypes;
import snabbdom from '@servicenow/ui-renderer-snabbdom';
import styles from './styles.scss';
import '@servicenow/now-button';
import '@servicenow/now-icon';
import '@servicenow/now-loader';
import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import MsgReaderDefault from 'msgreader/lib/MsgReader';
const MsgReader = MsgReaderDefault.default || MsgReaderDefault;

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

// ─── Preview fetching ────────────────────────────────────────────────────────

const PREVIEW_CATS = ['pdf', 'word', 'excel', 'msg', 'text'];

const buildMsgHtml = (MsgReader, buf, rawBuf) => {
	const reader = new MsgReader(buf);
	const d = reader.getFileData();
	const from = d.senderName ? `${d.senderName} &lt;${d.senderEmail || ''}&gt;` : (d.senderEmail || '—');
	const to   = (d.recipients || []).map(r => r.name || r.email || '').filter(Boolean).join('; ') || '—';

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

	// Try to extract HTML body from raw buffer if plaintext body is missing
	let bodyHtml;
	if (d.body) {
		bodyHtml = `<pre style="white-space:pre-wrap;font-family:inherit;font-size:14px;margin:0;line-height:1.6">${d.body}</pre>`;
	} else {
		// MsgReader doesn't parse HTML body (MAPI 0x1013), extract from raw buffer
		// Try Windows-1252 first (common for Outlook), fallback to UTF-8
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

const fetchPreview = (attachment, updateState) => {
	const { sys_id, file_type } = attachment;
	const cat = getCategory(file_type);
	if (!PREVIEW_CATS.includes(cat)) return;

	updateState({ blobUrl: null, previewText: null, previewHtml: null, previewPages: null, previewSheets: null, activeSheet: 0, previewLoading: true });

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

	// DOCX / XLSX / MSG — fetch as ArrayBuffer then process
	fetch(downloadUrl(sys_id), { credentials: 'same-origin' })
		.then(res => res.arrayBuffer())
		.then(buf => {
			if (cat === 'word') {
				return mammoth.convertToHtml({ arrayBuffer: buf }, {
					styleMap: [
						"br[type='page'] => hr",
						"p[style-name='Title'] => h1.doc-title:fresh",
						"p[style-name='Titel'] => h1.doc-title:fresh",
						"p[style-name='Koptekst 1'] => h1:fresh",
						"p[style-name='Kop 1'] => h1:fresh",
						"p[style-name='Heading 1'] => h1:fresh",
						"p[style-name='Ondertitel'] => h2.doc-subtitle:fresh",
						"p[style-name='Subtitle'] => h2.doc-subtitle:fresh",
						"r[style-name='Kop 1 Char'] => span.av-run-h1",
						"r[style-name='Heading 1 Char'] => span.av-run-h1"
					]
				}).then(r => {
					let html = r.value;
					// Strip av-run-h1 markers inside real headings
					html = html.replace(/<(h[1-6][^>]*)><span class="av-run-h1">([\s\S]*?)<\/span>(\s*)<\/(h[1-6])>/gi, '<$1>$2$3</$4>');
					// Promote av-run-h1 inside <p> to <h1>
					html = html.replace(/<p>([\s\S]*?)<span class="av-run-h1">([\s\S]*?)<\/span>([\s\S]*?)<\/p>/gi, (match, before, heading, after) => {
						const parts = [];
						const trimBefore = before.replace(/<br\s*\/?>/g, '').trim();
						if (trimBefore) parts.push('<p>' + before + '</p>');
						parts.push('<h1>' + heading + '</h1>');
						const trimAfter = after.replace(/<br\s*\/?>/g, '').trim();
						if (trimAfter) parts.push('<p>' + after + '</p>');
						return parts.join('');
					});
					const pages = html.split(/<hr\s*\/?>/i).filter(p => p.trim());
					updateState({ previewPages: pages, previewLoading: false });
				});
			}
			if (cat === 'excel') {
				const wb = XLSX.read(buf, { type: 'array' });
				const sheets = wb.SheetNames.map(name => {
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
					if (lastRow && lastCol) {
						ws['!ref'] = 'A1:' + XLSX.utils.encode_cell({ r: lastRow - 1, c: lastCol - 1 });
					}
					return { name, html: XLSX.utils.sheet_to_html(ws) };
				});
				updateState({ previewSheets: sheets, activeSheet: 0, previewLoading: false });
			}
			if (cat === 'msg') {
				const html = buildMsgHtml(MsgReader, buf, buf);
				updateState({ previewHtml: html, previewLoading: false });
			}
		})
		.catch(err => {
			console.error('[AV] Preview error:', err);
			updateState({ previewLoading: false });
		});
};

// ─── Upload attachments ──────────────────────────────────────────────────────

const uploadFile = (file, table, sysid, updateState) => {
	const url = `/api/now/attachment/file?table_name=${table}&table_sys_id=${sysid}&file_name=${encodeURIComponent(file.name)}`;
	updateState({ uploading: true });

	fetch(url, {
		method: 'POST',
		credentials: 'same-origin',
		headers: { 'Content-Type': file.type || 'application/octet-stream' },
		body: file
	})
		.then(res => {
			if (!res.ok) throw new Error('Upload failed: ' + res.status);
			return res.json();
		})
		.then(() => {
			updateState({ uploading: false });
			fetchAttachments(table, sysid, updateState);
		})
		.catch(err => {
			console.error('[AV] Upload error:', err);
			updateState({ uploading: false });
		});
};

const MAX_FILE_SIZE = 24 * 1024 * 1024; // 24MB ServiceNow default limit

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
			console.error('[AV] Upload error:', err);
			updateState({ uploading: false, uploadError: err.message });
			fetchAttachments(table, sysid, updateState);
		});
};

// ─── Fetch attachments ────────────────────────────────────────────────────────

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
			updateState({ attachments, selectedId: first ? first.sys_id : null, loading: false, blobUrl: null, previewText: null, previewHtml: null, previewPages: null, previewSheets: null, activeSheet: 0 });
			if (first) fetchPreview(first, updateState);
		})
		.catch(() => updateState({ loading: false }));
};

// ─── Preview ─────────────────────────────────────────────────────────────────

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
	const { blobUrl, previewText, previewHtml, previewPages, previewSheets, previewLoading } = state;
	const activeSheet = state.activeSheet || 0;

	// Loading spinner (PDF / DOCX / XLSX / MSG)
	if (previewLoading && PREVIEW_CATS.includes(cat)) {
		return (
			<div className="av-empty">
				<div className="av-spinner" />
				<p>Laden...</p>
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

	// PDF — use <object> to avoid CSP frame-src restrictions on blob: URLs
	if (cat === 'pdf' && blobUrl) {
		return (
			<div className="av-preview -pdf">
				<object data={blobUrl} type="application/pdf" aria-label={file_name} />
			</div>
		);
	}

	// DOCX — pages
	if (cat === 'word' && previewPages) {
		return (
			<div key={'docx-' + sys_id} className="av-preview -docx">
				{previewPages.map((page, i) => (
					<div key={i} className="av-docx-page" hook-insert={(vnode) => { vnode.elm.innerHTML = page; }} />
				))}
			</div>
		);
	}

	// XLSX — tabs + table
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
				<div className="av-excel-sheet" hook-insert={(vnode) => { vnode.elm.innerHTML = sheet.html; }} hook-update={(o, vnode) => { vnode.elm.innerHTML = sheet.html; }} />
			</div>
		);
	}

	// MSG — email
	if (cat === 'msg' && previewHtml) {
		return (
			<div key={'msg-' + sys_id} className="av-preview -msg" hook-insert={(vnode) => { vnode.elm.innerHTML = previewHtml; }} />
		);
	}

	// Plain text
	if (cat === 'text' && previewText !== null) {
		return (
			<div className="av-preview -text">
				<pre className="av-preview-text">{previewText}</pre>
			</div>
		);
	}

	// Unsupported / failed to load
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

// ─── View ─────────────────────────────────────────────────────────────────────

const view = (state, { dispatch }) => {
	const attachments = state.attachments || [];
	const selectedId  = state.selectedId || null;
	const loading     = state.loading || false;
	const selected    = attachments.find(a => a.sys_id === selectedId) || null;

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
				<div className="av-empty-card">
					<h2 className="av-empty-title">No attachments</h2>
					<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'OPEN_FILE_PICKER' }))}>
						<now-button
							label="Upload"
							variant="primary"
							size="md"
							icon="upload-outline"
						/>
					</span>
				</div>
			</div>
		);
	}

	const dragging = state.dragging || false;
	const uploading = state.uploading || false;
	const uploadError = state.uploadError || null;
	const confirmDelete = state.confirmDelete || null;

	return (
		<div className="av-root">
			{/* Delete confirmation overlay */}
			{confirmDelete && (
				<div className="av-drop-overlay" style={{ pointerEvents: 'auto' }}>
					<div className="av-confirm-card">
						<now-icon icon="trash-outline" size="lg" />
						<p>Weet je zeker dat je <strong>{confirmDelete.file_name}</strong> wilt verwijderen?</p>
						<div className="av-confirm-actions">
							<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'DELETE_ATTACHMENT', payload: { sys_id: confirmDelete.sys_id } }))}>
								<now-button label="Verwijderen" variant="primary-negative" size="md" />
							</span>
							<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'DISMISS_DELETE' }))}>
								<now-button label="Annuleren" variant="secondary" size="md" />
							</span>
						</div>
					</div>
				</div>
			)}

			{/* Upload error overlay */}
			{uploadError && (
				<div className="av-drop-overlay" style={{ pointerEvents: 'auto' }}
					onclick={() => dispatch(() => ({ type: 'DISMISS_ERROR' }))}
				>
					<div className="av-upload-card -error">
						<now-icon icon="circle-exclamation-outline" size="lg" />
						<p>{uploadError}</p>
						<now-button label="Sluiten" variant="secondary" size="sm" />
					</div>
				</div>
			)}

			{/* Drag overlay */}
			{dragging && (
				<div className="av-drop-overlay">
					<div className="av-drop-message">
						<now-icon icon="upload-outline" size="lg" />
						<p>Drop bestanden om te uploaden</p>
					</div>
				</div>
			)}

			{/* Upload progress overlay */}
			{uploading && (
				<div className="av-drop-overlay">
					<div className="av-upload-card">
						<now-loader label="Uploading..." action="Cancel" />
					</div>
				</div>
			)}

			{/* Delete progress overlay */}
			{state.deleting && (
				<div className="av-drop-overlay">
					<div className="av-upload-card">
						<now-loader label="Verwijderen..." />
					</div>
				</div>
			)}

			{/* Sidebar */}
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
						<now-button
							label="Upload"
							variant="secondary"
							size="md"
							icon="upload-outline"
						/>
					</div>
					<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'FETCH_ATTACHMENTS' }))}>
						<now-button-iconic icon="sync-outline" variant="tertiary" size="md" tooltipContent="Vernieuwen" />
					</span>
				</div>
			</div>

			{/* Main */}
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

// ─── Action Handlers ──────────────────────────────────────────────────────────

const actionHandlers = {
	[COMPONENT_BOOTSTRAPPED]: ({ host, dispatch, updateState, properties }) => {
		console.log(`[AV] Attachment Viewer v${AV_VERSION} (${AV_BUILD})`);
		if (properties.sysid) {
			fetchAttachments(properties.table, properties.sysid, updateState);
		}
		// Bind drag & drop on host element
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
		updateState({ selectedId: sys_id, blobUrl: null, previewText: null, previewHtml: null, previewPages: null, previewSheets: null, activeSheet: 0 });
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
				console.error('[AV] Delete error:', err);
				updateState({ deleting: false, uploadError: err.message });
			});
	},

};

// ─── Register ─────────────────────────────────────────────────────────────────

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
