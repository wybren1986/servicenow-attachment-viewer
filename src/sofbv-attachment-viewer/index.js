import { createCustomElement, actionTypes } from '@servicenow/ui-core';
const { COMPONENT_BOOTSTRAPPED, COMPONENT_PROPERTY_CHANGED } = actionTypes;
import { createHttpEffect } from '@servicenow/ui-effect-http';
import snabbdom from '@servicenow/ui-renderer-snabbdom';
import styles from './styles.scss';
import '@servicenow/now-button';
import '@servicenow/now-icon';
import '@servicenow/now-loader';
import '@servicenow/now-message';
import '@servicenow/now-illustration';
import '@servicenow/now-toggle';
import * as XLSX from 'xlsx';
import MsgReaderDefault from 'msgreader/lib/MsgReader';
const MsgReader = MsgReaderDefault.default || MsgReaderDefault;

const AV_VERSION = '4.0.0';
const AV_BUILD = '2026-04-14';

// ─── setImmediate polyfill ────────────────────────────────────────────────────
// docx-preview uses a setImmediate polyfill that falls back to postMessage.
// ServiceNow's post-message.js intercepts all postMessage events and tries to
// JSON.parse them, which breaks on the "setImmediate" string payload.
if (typeof window !== 'undefined' && !window.setImmediate) {
	window.setImmediate = (fn, ...args) => setTimeout(fn, 0, ...args);
	window.clearImmediate = (id) => clearTimeout(id);
}

// Lazy-load docx-preview so the polyfill above is registered first
let _renderAsync = null;
const getRenderAsync = () => {
	if (!_renderAsync) {
		_renderAsync = import('docx-preview').then(m => m.renderAsync);
	}
	return _renderAsync;
};

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

// sys_attachment.do is the UI-servlet path: on session expiry it redirects
// to /login.do instead of returning 401 WWW-Authenticate: Basic, so it won't
// trigger the browser's native auth popup for <a href> / <img src> usage.
const downloadUrl = (sys_id) => `/sys_attachment.do?sys_id=${encodeURIComponent(sys_id)}`;

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
			<strong style="font-size:12px;color:#718096;text-transform:uppercase;letter-spacing:.5px">Attachments (${msgAttachments.length})</strong>
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
			bodyHtml = '<em style="color:#718096">No content</em>';
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

// ─── Helpers ────────────────────────────────────────────────────────────────

const PREVIEW_CATS = ['pdf', 'word', 'excel', 'msg', 'text'];

const mapAttachment = (item) => ({
	sys_id: item.sys_id,
	file_name: item.file_name,
	size_bytes: item.size_bytes,
	content_type: item.content_type,
	file_type: (item.file_name || '').split('.').pop().toLowerCase(),
	sys_created_on: item.sys_created_on
});

const findAttachment = (state, sys_id) => {
	const direct = (state.attachments || []).find(a => a.sys_id === sys_id);
	if (direct) return direct;
	const parents = (state.parentChain || []).flatMap(p => p.attachments);
	return parents.find(a => a.sys_id === sys_id) || null;
};

// Process a downloaded blob into preview state based on category.
const processPreviewBlob = (blob, cat, updateState) => {
	if (cat === 'pdf') {
		updateState({ blobUrl: URL.createObjectURL(blob), previewLoading: false, pendingPreviewId: null });
		return;
	}
	if (cat === 'text') {
		blob.text().then(text => updateState({ previewText: text, previewLoading: false, pendingPreviewId: null }));
		return;
	}
	blob.arrayBuffer().then(buf => {
		if (cat === 'word') {
			updateState({ docxData: buf, previewLoading: false, pendingPreviewId: null });
		} else if (cat === 'excel') {
			updateState({ previewSheets: buildExcelSheets(buf), activeSheet: 0, previewLoading: false, pendingPreviewId: null });
		} else if (cat === 'msg') {
			updateState({ previewHtml: buildMsgHtml(MsgReader, buf, buf), previewLoading: false, pendingPreviewId: null });
		}
	});
};

// ─── Render Preview ─────────────────────────────────────────────────────────

const renderPreview = (attachment, state, dispatch) => {
	if (!attachment) {
		return (
			<div className="av-preview-state">
				<p>Select an attachment to preview</p>
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
					getRenderAsync().then(renderAsync => {
						renderAsync(state.docxData, vnode.elm, null, {
							className: 'av-docx-container',
							inWrapper: true,
							ignoreWidth: false,
							ignoreHeight: false,
							ignoreFonts: false,
							renderFontTable: true,
							useBase64URL: true,
							renderHeaders: true,
							renderFooters: true,
							renderFootnotes: true
						});
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
	const allParentAtts = (state.parentChain || []).flatMap(p => p.attachments);
	const isParentAtt = allParentAtts.some(a => a.sys_id === selectedId);
	const selected = attachments.find(a => a.sys_id === selectedId)
		|| allParentAtts.find(a => a.sys_id === selectedId)
		|| null;

	if (loading) {
		return (
			<div className="av-empty">
				<now-loader label="Loading" />
			</div>
		);
	}

	if (!attachments.length) {
		const emptyDragging = state.dragging || false;
		const emptyUploading = state.uploading || false;
		const parentHasAtts = (state.parentChain || []).some(p => p.attachments.length > 0);
		const showParentList = state.properties.enableParent && state.showParent && parentHasAtts;

		if (showParentList) {
			return (
				<div className="av-root">
					{emptyDragging && (
						<div className="av-drop-overlay">
							<div className="av-drop-message">
								<now-icon icon="upload-outline" size="lg" />
								<p>Drop files to upload</p>
							</div>
						</div>
					)}
					{emptyUploading && (
						<div className="av-drop-overlay">
							<div className="av-upload-card">
								<now-loader label="Uploading..." action="Cancel" />
							</div>
						</div>
					)}
					<div className="av-sidebar">
						<div className="av-list">
							{state.parentChain.map(parent => (
								parent.attachments.length > 0 && (
									<div className="av-parent-section" key={parent.sysid}>
										<div className="av-parent-header">
											<span className="av-parent-label">{parent.display}</span>
										</div>
										<ul className="av-list-items">
											{parent.attachments.map(a => (
												<li
													key={'p-' + a.sys_id}
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
									</div>
								)
							))}
						</div>
						<div className="av-sidebar-footer">
							<div className="av-sidebar-actions">
								<div className="av-toggle-inline">
									<span className="av-toggle-label">Show parent</span>
									<now-toggle checked={true} size="sm" />
								</div>
								<div className="av-actions-right">
									<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'OPEN_FILE_PICKER' }))}>
										<now-button-iconic icon="plus-outline" variant="primary" size="md" tooltipContent="Upload" />
									</span>
									<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'FETCH_ATTACHMENTS' }))}>
										<now-button-iconic icon="sync-outline" variant="tertiary" size="md" tooltipContent="Refresh" />
									</span>
								</div>
							</div>
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
								</div>
							</div>
						)}
						<div className="av-content">
							{renderPreview(selected, state, dispatch)}
						</div>
					</div>
				</div>
			);
		}

		return (
			<div className="av-empty" style={{ position: 'relative' }}>
				{emptyDragging && (
					<div className="av-drop-overlay">
						<div className="av-drop-message">
							<now-icon icon="upload-outline" size="lg" />
							<p>Drop files to upload</p>
						</div>
					</div>
				)}
				{emptyUploading && (
					<div className="av-drop-overlay">
						<div className="av-upload-card">
							<now-loader label="Uploading..." action="Cancel" />
						</div>
					</div>
				)}
				<now-message alignment="vertical-centered">
					<now-illustration slot="media" illustration="add-attachment" size="auto" />
					<div slot="message">
						<h3 className="now-heading--md now-m-block--0">No attachments available</h3>
						<p className="now-m-block-start--sm now-m-block-end--0 now-color_text--tertiary">Drag or select files to upload</p>
					</div>
					<div slot="actions" className="av-empty-actions -row">
						<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'OPEN_FILE_PICKER' }))}>
							<now-button label="Select file" variant="primary" size="md" />
						</span>
						{state.properties.enableParent && parentHasAtts && (
							<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'NOW_TOGGLE#CHECKED_SET', payload: { value: true } }))}>
								<now-button label="Show parent attachments" variant="secondary" size="md" />
							</span>
						)}
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
								<h3 className="now-heading--md now-m-block--0">Delete file</h3>
								<p className="now-m-block-start--sm now-m-block-end--0">Are you sure you want to delete <strong>{confirmDelete.file_name}</strong>?</p>
							</div>
							<div slot="actions" className="av-confirm-actions">
								<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'DELETE_ATTACHMENT', payload: { sys_id: confirmDelete.sys_id } }))}>
									<now-button label="Delete" variant="primary-negative" size="md" />
								</span>
								<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'DISMISS_DELETE' }))}>
									<now-button label="Cancel" variant="secondary" size="md" />
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
								<h3 className="now-heading--md now-m-block--0">Upload failed</h3>
								<p className="now-m-block-start--sm now-m-block-end--0">{uploadError}</p>
							</div>
							<div slot="actions">
								<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'DISMISS_ERROR' }))}>
									<now-button label="Close" variant="secondary" size="md" />
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
						<p>Drop files to upload</p>
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
						<now-loader label="Deleting..." />
					</div>
				</div>
			)}

			<div className="av-sidebar">
				<div className="av-list">
					<ul className="av-list-items">
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

					{state.showParent && state.parentChain && state.parentChain.map(parent => (
						parent.attachments.length > 0 && (
							<div className="av-parent-section" key={parent.sysid}>
								<div className="av-parent-header">
									<span className="av-parent-label">{parent.display}</span>
								</div>
								<ul className="av-list-items">
									{parent.attachments.map(a => (
										<li
											key={'p-' + a.sys_id}
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
							</div>
						)
					))}
				</div>

				<div className="av-sidebar-footer">
					<div className="av-sidebar-actions">
						{state.properties.enableParent && (
							<div className="av-toggle-inline">
								<span className="av-toggle-label">Show parent</span>
								<now-toggle checked={state.showParent || false} disabled={!state.parentChain || !state.parentChain.length} size="sm" />
							</div>
						)}
						<div className="av-actions-right">
							<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'OPEN_FILE_PICKER' }))}>
								<now-button-iconic icon="plus-outline" variant="primary" size="md" tooltipContent="Upload" />
							</span>
							<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'FETCH_ATTACHMENTS' }))}>
								<now-button-iconic icon="sync-outline" variant="tertiary" size="md" tooltipContent="Refresh" />
							</span>
						</div>
					</div>
				</div>
			</div>

			<div className="av-main">
				{selected && (
					<div className="av-header">
						{state.editingName && !isParentAtt ? (
							<div className="av-header-edit">
								<div className="av-file-icon">{fileIcon(selected.file_type)}</div>
								<input
									className="av-rename-input"
									value={state.editingName}
									hook-insert={(vnode) => {
										const el = vnode.elm;
										el.focus();
										el.select();
										el.addEventListener('input', () => {
											dispatch(() => ({ type: 'SET_EDITING_NAME', payload: el.value }));
										});
										el.addEventListener('keydown', (e) => {
											if (e.key === 'Enter') dispatch(() => ({ type: 'SAVE_RENAME', payload: { value: el.value } }));
											if (e.key === 'Escape') dispatch(() => ({ type: 'CANCEL_RENAME' }));
										});
									}}
								/>
								<span className="av-rename-ext">{state.editingExt}</span>
								<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'SAVE_RENAME' }))}>
									<now-button-iconic icon="check-outline" variant="primary" size="md" tooltipContent="Save" />
								</span>
								<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'CANCEL_RENAME' }))}>
									<now-button-iconic icon="ban-outline" variant="tertiary" size="md" tooltipContent="Cancel" />
								</span>
							</div>
						) : (
							<div className="av-header-title">
								<div className="av-file-icon">{fileIcon(selected.file_type)}</div>
								<div>
									<div className="av-header-filename">{selected.file_name}</div>
									<div className="av-header-meta">{formatSize(selected.size_bytes)}</div>
								</div>
							</div>
						)}
						{!state.editingName && (
							<div className="av-header-actions">
								{!isParentAtt && (
									<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'START_RENAME', payload: selected.file_name }))}>
										<now-button-iconic icon="pencil-page-outline" variant="tertiary" size="md" tooltipContent="Rename" />
									</span>
								)}
								<a href={downloadUrl(selected.sys_id)} download={selected.file_name} className="av-download-link">
									<now-button-iconic icon="download-outline" variant="tertiary" size="md" tooltipContent="Download" />
								</a>
								{!isParentAtt && (
									<span className="av-icon-wrap" onclick={() => dispatch(() => ({ type: 'CONFIRM_DELETE', payload: { sys_id: selected.sys_id, file_name: selected.file_name } }))}>
										<now-button-iconic icon="trash-outline" variant="tertiary" size="md" tooltipContent="Delete" />
									</span>
								)}
							</div>
						)}
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
	// ─── Lifecycle ──────────────────────────────────────────────────────────
	[COMPONENT_BOOTSTRAPPED]: ({ host, dispatch, properties }) => {
		if (properties.sysid) {
			dispatch('LIST_ATTACHMENTS', { table: properties.table, sysid: properties.sysid });
			if (properties.enableParent) {
				dispatch('PARENT_FETCH_RECORD', { table: properties.table, sysid: properties.sysid });
			}
		}
		// Auto-size height
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
			const active = host.shadowRoot && host.shadowRoot.activeElement;
			if (active && active.tagName === 'INPUT') return;
			dispatch('SET_DRAGGING', true);
		});
		host.addEventListener('dragover', (e) => { e.preventDefault(); });
		host.addEventListener('dragleave', (e) => {
			if (!host.contains(e.relatedTarget)) dispatch('SET_DRAGGING', false);
		});
		host.addEventListener('drop', (e) => {
			e.preventDefault();
			dispatch('SET_DRAGGING', false);
			if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
				dispatch('UPLOAD_FILES', { files: e.dataTransfer.files });
			}
		});
	},

	[COMPONENT_PROPERTY_CHANGED]: ({ action, dispatch, properties }) => {
		const { name } = action.payload || {};
		if ((name === 'sysid' || name === 'table') && properties.sysid) {
			dispatch('LIST_ATTACHMENTS', { table: properties.table, sysid: properties.sysid });
			if (properties.enableParent) {
				dispatch('PARENT_FETCH_RECORD', { table: properties.table, sysid: properties.sysid });
			}
		}
	},

	// ─── List attachments ───────────────────────────────────────────────────
	'FETCH_ATTACHMENTS': ({ dispatch, properties }) => {
		dispatch('LIST_ATTACHMENTS', { table: properties.table, sysid: properties.sysid });
	},

	'LIST_ATTACHMENTS': ({ dispatch, updateState, action }) => {
		const { table, sysid, keepSelectedId } = action.payload;
		updateState({ loading: true, _keepSelectedId: keepSelectedId || null });
		dispatch('LIST_HTTP', {
			sysparm_query: `table_name=${table}^table_sys_id=${sysid}`,
			sysparm_limit: '200'
		});
	},

	'LIST_HTTP': createHttpEffect('/api/now/attachment', {
		method: 'GET',
		queryParams: ['sysparm_query', 'sysparm_limit'],
		successActionType: 'LIST_HTTP_SUCCESS',
		errorActionType: 'LIST_HTTP_ERROR'
	}),

	'LIST_HTTP_SUCCESS': ({ action, dispatch, updateState, state }) => {
		const result = Array.isArray(action.payload?.result) ? action.payload.result : [];
		const attachments = result.map(mapAttachment);
		const keepId = state._keepSelectedId;
		const kept = keepId && attachments.find(a => a.sys_id === keepId);
		const selected = kept || (attachments.length ? attachments[0] : null);
		updateState({
			attachments,
			selectedId: selected ? selected.sys_id : null,
			loading: false,
			blobUrl: null, previewText: null, previewHtml: null,
			previewSheets: null, docxData: null, activeSheet: 0,
			_keepSelectedId: null
		});
		if (selected && PREVIEW_CATS.includes(getCategory(selected.file_type))) {
			dispatch('PREVIEW_DOWNLOAD', { sys_id: selected.sys_id, cat: getCategory(selected.file_type) });
		}
	},

	'LIST_HTTP_ERROR': ({ action, updateState }) => {
		console.warn('[AV] list error', action.payload);
		updateState({ loading: false, uploadError: 'Failed to load attachments' });
	},

	// ─── Preview download ───────────────────────────────────────────────────
	'SELECT_ATTACHMENT': ({ action, dispatch, updateState, state }) => {
		const sys_id = action.payload.sys_id;
		const att = findAttachment(state, sys_id);
		if (!att) return;
		const cat = getCategory(att.file_type);
		updateState({
			selectedId: sys_id,
			blobUrl: null, previewText: null, previewHtml: null,
			previewSheets: null, docxData: null, activeSheet: 0
		});
		if (PREVIEW_CATS.includes(cat)) {
			dispatch('PREVIEW_DOWNLOAD', { sys_id, cat });
		}
	},

	// Binary preview download via the platform's own HTTP transport.
	//
	// `createHttpEffect` does not forward `responseType` to sn-http-request,
	// but the underlying transport (exposed as `window.nowUiFramework.snHttp`)
	// DOES support `responseType: 'blob'`. Calling it directly gives us:
	// - Same authenticated cookie + X-UserToken handling as createHttpEffect
	//   → no browser Basic-auth popup on session expiry.
	// - A real Blob back, usable by docx-preview / xlsx / msgreader.
	'PREVIEW_DOWNLOAD': ({ action, updateState }) => {
		const { sys_id, cat } = action.payload;
		updateState({ previewLoading: true, pendingPreviewId: sys_id, pendingPreviewCat: cat });
		const snHttp = window.nowUiFramework && window.nowUiFramework.snHttp;
		if (!snHttp) {
			console.warn('[AV] preview: nowUiFramework.snHttp unavailable');
			updateState({ previewLoading: false, pendingPreviewId: null, pendingPreviewCat: null });
			return;
		}
		// batch: false — otherwise snHttp wraps this in /api/now/batch which
		// multiplexes responses and can't pass a Blob back cleanly.
		snHttp.request(`/api/now/attachment/${sys_id}/file`, 'GET', { responseType: 'blob', batch: false })
			.then(res => {
				const blob = res && (res.data instanceof Blob ? res.data : (res instanceof Blob ? res : null));
				if (!blob) throw new Error('preview: unexpected response');
				processPreviewBlob(blob, cat, updateState);
			})
			.catch(err => {
				console.warn('[AV] preview error', err);
				updateState({ previewLoading: false, pendingPreviewId: null, pendingPreviewCat: null });
			});
	},

	// ─── Upload ─────────────────────────────────────────────────────────────
	'OPEN_FILE_PICKER': ({ dispatch }) => {
		const input = document.createElement('input');
		input.type = 'file';
		input.multiple = true;
		input.onchange = () => {
			if (input.files.length) dispatch('UPLOAD_FILES', { files: input.files });
		};
		input.click();
	},

	'UPLOAD_FILES': ({ action, dispatch, updateState, properties }) => {
		const files = Array.from(action.payload.files || []);
		if (!files.length || !properties.sysid) return;
		const tooLarge = files.filter(f => f.size > MAX_FILE_SIZE);
		if (tooLarge.length) {
			const names = tooLarge.map(f => `${f.name} (${formatSize(f.size)})`).join(', ');
			updateState({ uploadError: `File(s) too large (max ${formatSize(MAX_FILE_SIZE)}): ${names}` });
			return;
		}
		updateState({ uploading: true, uploadError: null, _uploadPending: files.length });
		files.forEach(file => {
			const fd = new FormData();
			fd.append('table_name', properties.table);
			fd.append('table_sys_id', properties.sysid);
			fd.append('uploadFile', file, file.name);
			dispatch('UPLOAD_HTTP', { formData: fd, _fileName: file.name });
		});
	},

	'UPLOAD_HTTP': createHttpEffect('/api/now/attachment/upload', {
		method: 'POST',
		dataParam: 'formData',
		successActionType: 'UPLOAD_HTTP_SUCCESS',
		errorActionType: 'UPLOAD_HTTP_ERROR'
	}),

	'UPLOAD_HTTP_SUCCESS': ({ dispatch, updateState, state, properties }) => {
		const remaining = (state._uploadPending || 1) - 1;
		updateState({ _uploadPending: remaining });
		if (remaining <= 0) {
			updateState({ uploading: false, _uploadPending: 0 });
			dispatch('LIST_ATTACHMENTS', { table: properties.table, sysid: properties.sysid });
		}
	},

	'UPLOAD_HTTP_ERROR': ({ action, dispatch, updateState, state, properties }) => {
		console.warn('[AV] upload error', action.payload);
		const remaining = (state._uploadPending || 1) - 1;
		const msg = action.payload?.statusText || action.payload?.message || 'Upload failed';
		updateState({ _uploadPending: remaining, uploadError: msg });
		if (remaining <= 0) {
			updateState({ uploading: false, _uploadPending: 0 });
			dispatch('LIST_ATTACHMENTS', { table: properties.table, sysid: properties.sysid });
		}
	},

	// ─── Rename ─────────────────────────────────────────────────────────────
	'START_RENAME': ({ action, updateState }) => {
		const fileName = action.payload;
		const lastDot = fileName.lastIndexOf('.');
		const nameOnly = lastDot > 0 ? fileName.substring(0, lastDot) : fileName;
		updateState({ editingName: nameOnly, editingExt: lastDot > 0 ? fileName.substring(lastDot) : '' });
	},

	'SET_EDITING_NAME': ({ action, updateState }) => {
		updateState({ editingName: action.payload });
	},

	'CANCEL_RENAME': ({ updateState }) => {
		updateState({ editingName: null, editingExt: null });
	},

	'SAVE_RENAME': ({ action, state, dispatch, updateState }) => {
		const nameOnly = (action.payload?.value || state.editingName || '').trim();
		const ext = state.editingExt || '';
		const selected = (state.attachments || []).find(a => a.sys_id === state.selectedId);
		const fullName = nameOnly + ext;
		if (!nameOnly || !selected || fullName === selected.file_name) {
			updateState({ editingName: null, editingExt: null });
			return;
		}
		updateState({ editingName: null, editingExt: null, _renameKeepId: selected.sys_id });
		dispatch('RENAME_HTTP', { sys_id: selected.sys_id, data: { file_name: fullName } });
	},

	'RENAME_HTTP': createHttpEffect('/api/now/table/sys_attachment/:sys_id', {
		method: 'PATCH',
		pathParams: ['sys_id'],
		dataParam: 'data',
		headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
		successActionType: 'RENAME_HTTP_SUCCESS',
		errorActionType: 'RENAME_HTTP_ERROR'
	}),

	'RENAME_HTTP_SUCCESS': ({ dispatch, state, properties }) => {
		dispatch('LIST_ATTACHMENTS', { table: properties.table, sysid: properties.sysid, keepSelectedId: state._renameKeepId });
	},

	'RENAME_HTTP_ERROR': ({ action, updateState }) => {
		updateState({ uploadError: action.payload?.statusText || 'Rename failed' });
	},

	// ─── Delete ─────────────────────────────────────────────────────────────
	'CONFIRM_DELETE': ({ action, updateState }) => {
		updateState({ confirmDelete: action.payload });
	},

	'DISMISS_DELETE': ({ updateState }) => {
		updateState({ confirmDelete: null });
	},

	'DELETE_ATTACHMENT': ({ action, dispatch, updateState }) => {
		updateState({ confirmDelete: null, deleting: true });
		dispatch('DELETE_HTTP', { sys_id: action.payload.sys_id });
	},

	'DELETE_HTTP': createHttpEffect('/api/now/attachment/:sys_id', {
		method: 'DELETE',
		pathParams: ['sys_id'],
		successActionType: 'DELETE_HTTP_SUCCESS',
		errorActionType: 'DELETE_HTTP_ERROR'
	}),

	'DELETE_HTTP_SUCCESS': ({ dispatch, updateState, properties }) => {
		updateState({ deleting: false });
		dispatch('LIST_ATTACHMENTS', { table: properties.table, sysid: properties.sysid });
	},

	'DELETE_HTTP_ERROR': ({ action, updateState }) => {
		updateState({ deleting: false, uploadError: action.payload?.statusText || 'Delete failed' });
	},

	// ─── Parent chain (recursive 3-step traversal) ──────────────────────────
	// Step 1: get .parent of current record
	'PARENT_FETCH_RECORD': ({ action, dispatch, state, updateState }) => {
		const { table, sysid } = action.payload;
		const seen = state._parentSeen || new Set();
		if (seen.has(sysid)) return;
		seen.add(sysid);
		updateState({ _parentSeen: seen, _parentChainBuild: state._parentChainBuild || [] });
		dispatch('PARENT_RECORD_HTTP', { table, sys_id: sysid });
	},

	'PARENT_RECORD_HTTP': createHttpEffect('/api/now/table/:table/:sys_id', {
		method: 'GET',
		pathParams: ['table', 'sys_id'],
		queryParams: ['sysparm_fields', 'sysparm_display_value'],
		successActionType: 'PARENT_RECORD_HTTP_SUCCESS',
		errorActionType: 'PARENT_CHAIN_DONE'
	}),

	'PARENT_RECORD_HTTP_SUCCESS': ({ action, dispatch, state, updateState }) => {
		const parent = action.payload?.result?.parent;
		if (!parent || !parent.value || (state._parentSeen && state._parentSeen.has(parent.value))) {
			updateState({ parentChain: state._parentChainBuild || [], _parentSeen: null, _parentChainBuild: null });
			return;
		}
		// Step 2: resolve actual table via sys_class_name
		dispatch('PARENT_TASK_HTTP', { sys_id: parent.value, _parentValue: parent.value, _parentDisplay: parent.display_value });
		updateState({ _pendingParentValue: parent.value, _pendingParentDisplay: parent.display_value });
	},

	'PARENT_TASK_HTTP': createHttpEffect('/api/now/table/task/:sys_id', {
		method: 'GET',
		pathParams: ['sys_id'],
		queryParams: ['sysparm_fields', 'sysparm_display_value', 'sysparm_limit'],
		successActionType: 'PARENT_TASK_HTTP_SUCCESS',
		errorActionType: 'PARENT_CHAIN_DONE'
	}),

	'PARENT_TASK_HTTP_SUCCESS': ({ action, dispatch, state, updateState }) => {
		const r = action.payload?.result;
		const parentTable = r?.sys_class_name?.value || r?.sys_class_name || 'task';
		const number = r?.number?.display_value || r?.number || state._pendingParentDisplay || state._pendingParentValue;
		const sysid = state._pendingParentValue;
		const entry = { table: parentTable, sysid, display: number, attachments: [] };
		const chain = (state._parentChainBuild || []).concat(entry);
		updateState({ _parentChainBuild: chain, _pendingParentTable: parentTable });
		// Step 3: fetch attachments for this parent
		dispatch('PARENT_ATT_HTTP', {
			sysparm_query: `table_name=${parentTable}^table_sys_id=${sysid}`,
			sysparm_limit: '200'
		});
	},

	'PARENT_ATT_HTTP': createHttpEffect('/api/now/attachment', {
		method: 'GET',
		queryParams: ['sysparm_query', 'sysparm_limit'],
		successActionType: 'PARENT_ATT_HTTP_SUCCESS',
		errorActionType: 'PARENT_CHAIN_DONE'
	}),

	'PARENT_ATT_HTTP_SUCCESS': ({ action, dispatch, state, updateState }) => {
		const result = Array.isArray(action.payload?.result) ? action.payload.result : [];
		const chain = (state._parentChainBuild || []).slice();
		if (chain.length) chain[chain.length - 1].attachments = result.map(mapAttachment);
		updateState({ _parentChainBuild: chain, parentChain: chain });
		// Recurse: get this parent's parent
		dispatch('PARENT_FETCH_RECORD', { table: state._pendingParentTable, sysid: state._pendingParentValue });
	},

	'PARENT_CHAIN_DONE': ({ state, updateState }) => {
		updateState({ parentChain: state._parentChainBuild || [], _parentSeen: null, _parentChainBuild: null });
	},

	// ─── UI state ───────────────────────────────────────────────────────────
	'SET_DRAGGING': ({ action, updateState }) => {
		updateState({ dragging: action.payload });
	},

	'SET_ACTIVE_SHEET': ({ action, updateState }) => {
		updateState({ activeSheet: action.payload.index });
	},

	'NOW_TOGGLE#CHECKED_SET': ({ action, updateState }) => {
		updateState({ showParent: action.payload.value });
	},

	'NOW_LOADER#ACTION_CLICKED': ({ updateState }) => {
		updateState({ uploading: false });
	},

	'DISMISS_ERROR': ({ updateState }) => {
		updateState({ uploadError: null });
	}
};

// ─── Register ───────────────────────────────────────────────────────────────

createCustomElement('sofbv-attachment-viewer', {
	renderer: { type: snabbdom },
	view,
	styles,
	actionHandlers,
	properties: {
		table: { default: 'incident', schema: { type: 'string' } },
		sysid: { default: '', schema: { type: 'string' } },
		enableParent: { default: false, schema: { type: 'boolean' } }
	}
});
