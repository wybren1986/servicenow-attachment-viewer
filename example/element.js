import '../src/sofbv-attachment-viewer';

const el = document.createElement('DIV');
document.body.appendChild(el);

el.innerHTML = `<sofbv-attachment-viewer table="incident"></sofbv-attachment-viewer>`;

// Stel sysid iets later in zodat COMPONENT_PROPERTY_CHANGED correct vuurt
setTimeout(() => {
	document.querySelector('sofbv-attachment-viewer').sysid = 'd24d517d9737fe505e06b586f053af11';
}, 100);
