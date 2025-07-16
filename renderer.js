const { ipcRenderer } = require('electron');

document.getElementById('runButton').addEventListener('click', () => {
  const sharedMailbox = document.getElementById('sharedMailbox').value.trim();

  // collect all email inputs (email, email2, email3, …)
  const emails = Array.from(
    document.querySelectorAll('input[name^="email"]')
  )
  .map(input => input.value.trim())
  .filter(val => val.length > 0);

  // basic validation
  if (!sharedMailbox || emails.length === 0) {
    document.getElementById('status').innerText = 'Please fill in shared mailbox and at least one user email.';
    return;
  }

  document.getElementById('status').innerText = 'Running automation…';
  ipcRenderer.send('run-automation', { sharedMailbox, emails });
});

ipcRenderer.on('automation-complete', (event, message) => {
  document.getElementById('status').innerText = message;
});

ipcRenderer.on('automation-error', (event, message) => {
  document.getElementById('status').innerText = message;
});
