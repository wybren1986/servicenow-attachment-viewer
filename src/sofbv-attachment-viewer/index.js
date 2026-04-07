import { createCustomElement, COMPONENT_BOOTSTRAPPED, COMPONENT_PROPERTY_CHANGED } from '@servicenow/ui-core';
import snabbdom from '@servicenow/ui-renderer-snabbdom';
import styles from './styles.scss';

// ─── Helpers ────────────────────────────────────────────────────────────────

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
	if (['doc', 'docx'].includes(ext)) return 'word';
	if (IMAGE_TYPES.includes(ext))     return 'image';
	if (TEXT_TYPES.includes(ext))      return 'text';
	return 'unsupported';
};

const fileIcon = (ext) => {
	const cat = getCategory(ext);
	const colors = { pdf: '#e53e3e', excel: '#38a169', word: '#3182ce', msg: '#d69e2e', image: '#805ad5', text: '#718096', unsupported: '#a0aec0' };
	const labels = { pdf: 'PDF', excel: 'XLS', word: 'DOC', msg: 'MSG', image: 'IMG', text: 'TXT', unsupported: (ext || '???').toUpperCase().slice(0, 3) };
	const color = colors[cat] || '#a0aec0';
	const label = labels[cat] || (ext || '').toUpperCase().slice(0, 3);
	return (
		<div style={{
			width: '36px', height: '42px', borderRadius: '4px',
			border: `1.5px solid ${color}`, background: color + '1a',
			display: 'flex', alignItems: 'center', justifyContent: 'center',
			flexShrink: '0', fontSize: '9px', fontWeight: '700',
			color: color, fontFamily: 'sans-serif', letterSpacing: '0.5px'
		}}>
			{label}
		</div>
	);
};

const downloadUrl = (sys_id) => `/api/now/attachment/${sys_id}/file`;

const wrapHtml = (body) => `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; font-size: 14px; line-height: 1.6; color: #2d3748; padding: 24px; margin: 0; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 16px; }
  td, th { border: 1px solid #e2e8f0; padding: 6px 10px; }
  th { background: #f7fafc; font-weight: 600; }
  img { max-width: 100%; }
</style></head><body>${body}</body></html>`;

// ─── Preview fetching ────────────────────────────────────────────────────────

const PREVIEW_CATS = ['pdf', 'word', 'excel', 'msg', 'text'];

const buildMsgHtml = (MsgReader, buf) => {
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

	const bodyHtml = d.body
		? `<pre style="white-space:pre-wrap;font-family:inherit;font-size:14px;margin:0;line-height:1.6">${d.body}</pre>`
		: '<em style="color:#718096">Geen inhoud</em>';

	return wrapHtml(`
		<div style="background:#f7fafc;border:1px solid #e2e8f0;border-radius:8px;padding:16px;margin-bottom:20px;font-size:13px;line-height:1.8">
			<div><b>Van:</b> ${from}</div>
			<div><b>Aan:</b> ${to}</div>
			<div><b>Onderwerp:</b> ${d.subject || '—'}</div>
		</div>
		<div>${bodyHtml}</div>
		${attSection}
	`);
};

const fetchPreview = (attachment, updateState) => {
	const { sys_id, file_type } = attachment;
	const cat = getCategory(file_type);
	if (!PREVIEW_CATS.includes(cat)) return;

	updateState({ blobUrl: null, previewText: null, previewLoading: true });

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
				const mammoth = require('mammoth');
				return mammoth.convertToHtml({ arrayBuffer: buf }).then(r => wrapHtml(r.value));
			}
			if (cat === 'excel') {
				const XLSX = require('xlsx');
				const wb = XLSX.read(buf, { type: 'array' });
				const tabs = wb.SheetNames.map(name =>
					`<h3 style="margin:16px 0 8px;font-size:13px;color:#718096">${name}</h3>${XLSX.utils.sheet_to_html(wb.Sheets[name])}`
				).join('');
				return Promise.resolve(wrapHtml(tabs));
			}
			if (cat === 'msg') {
				const MsgReader = require('msgreader').default || require('msgreader');
				return Promise.resolve(buildMsgHtml(MsgReader, buf));
			}
		})
		.then(html => {
			const blob = new Blob([html], { type: 'text/html' });
			updateState({ blobUrl: URL.createObjectURL(blob), previewLoading: false });
		})
		.catch(() => updateState({ previewLoading: false }));
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
				file_type: (item.file_name || '').split('.').pop().toLowerCase()
			}));
			const first = attachments.length ? attachments[0] : null;
			updateState({ attachments, selectedId: first ? first.sys_id : null, loading: false, blobUrl: null, previewText: null });
			if (first) fetchPreview(first, updateState);
		})
		.catch(() => updateState({ loading: false }));
};

// ─── Preview ─────────────────────────────────────────────────────────────────

const renderPreview = (attachment, state) => {
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
	const { blobUrl, previewText, previewLoading } = state;

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
			<div className="av-preview -iframe">
				<object data={blobUrl} type="application/pdf" aria-label={file_name} />
			</div>
		);
	}

	// DOCX, XLSX, MSG — rendered in sandboxed iframe
	if (['word', 'excel', 'msg'].includes(cat) && blobUrl) {
		return (
			<div className="av-preview -iframe">
				<iframe src={blobUrl} title={file_name} sandbox="allow-same-origin" />
			</div>
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
				<a href={url} download={file_name} style={{ color: '#0073e6', textDecoration: 'none', fontWeight: '600' }}>
					Downloaden
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
	const sysid       = state.properties.sysid;
	const table       = state.properties.table;

	if (loading) {
		return (
			<div className="av-empty">
				<div className="av-spinner" />
				<p>Bijlagen laden...</p>
			</div>
		);
	}

	if (!attachments.length) {
		return (
			<div className="av-empty">
				<p style={{ color: '#718096' }}>
					{sysid ? 'Geen bijlagen gevonden' : 'Geen record geselecteerd'}
				</p>
				<p style={{ fontSize: '11px', color: '#a0aec0', fontFamily: 'monospace' }}>
					table: {table || '(leeg)'} | sysid: {sysid || '(leeg)'}
				</p>
				<button className="av-fetch-btn" onclick={() => dispatch(() => ({ type: 'FETCH_ATTACHMENTS' }))}>
					Haal attachments op
				</button>
			</div>
		);
	}

	return (
		<div className="av-root">
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
								<div className="av-list-item-meta">{formatSize(a.size_bytes)}</div>
							</div>
						</li>
					))}
				</ul>
				<button className="av-fetch-btn" onclick={() => dispatch(() => ({ type: 'FETCH_ATTACHMENTS' }))}>
					Vernieuwen
				</button>
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
							<a
								href={downloadUrl(selected.sys_id)}
								download={selected.file_name}
								style={{ display: 'inline-flex', alignItems: 'center', padding: '6px 12px', borderRadius: '6px', background: '#0073e6', color: '#fff', textDecoration: 'none', fontSize: '0.85rem', fontWeight: '600' }}
							>
								Download
							</a>
						</div>
					</div>
				)}
				<div className="av-content">
					{renderPreview(selected, state)}
				</div>
			</div>
		</div>
	);
};

// ─── Action Handlers ──────────────────────────────────────────────────────────

const actionHandlers = {
	[COMPONENT_BOOTSTRAPPED]: ({ updateState, properties }) => {
		if (properties.sysid) {
			fetchAttachments(properties.table, properties.sysid, updateState);
		}
	},

	[COMPONENT_PROPERTY_CHANGED]: ({ action, updateState, properties }) => {
		if (action.payload?.name === 'sysid' && action.payload?.value) {
			fetchAttachments(properties.table, action.payload.value, updateState);
		}
	},

	'FETCH_ATTACHMENTS': ({ updateState, properties }) => {
		fetchAttachments(properties.table, properties.sysid, updateState);
	},

	'SELECT_ATTACHMENT': ({ action, updateState, state }) => {
		const sys_id = action.payload.sys_id;
		updateState({ selectedId: sys_id, blobUrl: null, previewText: null });
		const attachment = (state.attachments || []).find(a => a.sys_id === sys_id);
		if (attachment) fetchPreview(attachment, updateState);
	}
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
