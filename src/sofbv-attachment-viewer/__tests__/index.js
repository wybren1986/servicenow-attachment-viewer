// Tests for sofbv-attachment-viewer
import { createCustomElement } from '@servicenow/ui-core';
import { createHttpEffect } from '@servicenow/ui-effect-http';

describe('sofbv-attachment-viewer', () => {
	it('should render loading state', () => {
		const element = createCustomElement('sofbv-attachment-viewer', {
			initialState: { loading: true }
		});
		expect(element).toBeDefined();
	});

	it('should render attachments list', () => {
		const attachments = [{ sys_id: '1', file_name: 'test.pdf' }];
		const element = createCustomElement('sofbv-attachment-viewer', {
			initialState: { attachments, loading: false }
		});
		expect(element).toBeDefined();
	});

	it('should render view modal', () => {
		const element = createCustomElement('sofbv-attachment-viewer', {
			initialState: { isViewing: true, viewContent: '<p>Test</p>', viewTitle: 'Test File' }
		});
		expect(element).toBeDefined();
	});
});
