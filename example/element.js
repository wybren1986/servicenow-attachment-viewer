import '../src/sofbv-attachment-viewer';

const el = document.createElement('DIV');
document.body.appendChild(el);

el.innerHTML = `<sofbv-attachment-viewer table="incident"></sofbv-attachment-viewer>`;

// Stel sysid iets later in zodat COMPONENT_PROPERTY_CHANGED correct vuurt
setTimeout(() => {
	const el = document.querySelector('sofbv-attachment-viewer');
	el.sysid = '2edffdbe97448b105e06b586f053af5d';
	el.enableParent = true;
}, 100);
