const { ipcRenderer } = require('electron');

document.getElementById('runButton').addEventListener('click', () => {
  const sharedMailbox = document.getElementById('sharedMailbox').value;
  const email = document.getElementById('email').value;
  if (!sharedMailbox || !email) {
    document.getElementById('status').innerText = 'Please fill in both fields.';
    return;
  }
  document.getElementById('status').innerText = 'Running automation...';
  ipcRenderer.send('run-automation', { sharedMailbox, email });
});

ipcRenderer.on('automation-complete', (event, message) => {
  document.getElementById('status').innerText = message;
});

ipcRenderer.on('automation-error', (event, message) => {
  document.getElementById('status').innerText = message;
});