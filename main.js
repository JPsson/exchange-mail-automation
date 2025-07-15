const { app, BrowserWindow, ipcMain } = require('electron');
const { chromium } = require('playwright');
const path = require('path');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });
  mainWindow.loadFile('index.html');
}

app.on('ready', createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

ipcMain.on('run-automation', async (event, data) => {
  const { sharedMailbox, email } = data;
  const userName = extractUserName(email);
  const fullName = toFullName(userName);

  try {
    const context = await chromium.launchPersistentContext(
      'C:/Users/JakobPettersson/PlaywrightChromiumProfile',
      { headless: false }
    );
    const page = await context.newPage();
    await page.goto('https://admin.exchange.microsoft.com');

    await page.getByRole('menuitem', { name: 'mailboxes' }).click();
    await page.getByRole('searchbox', { name: 'Search Mailboxes' }).click();
    await page.getByRole('searchbox', { name: 'Search Mailboxes' }).fill(sharedMailbox); 
    await page.getByRole('searchbox', { name: 'Search Mailboxes' }).press('Enter');

    //await page.getByLabel(`Display name ${sharedMailbox}`).click(); t
    await page.getByLabel(`Email address ${sharedMailbox}`).click(); 
    await page.getByRole('tab', { name: 'Delegation' }).click();
    await page.getByRole('button', { name: 'modifying send as member edit' }).click();
    await searchUserInDelegationPanels(page, userName, fullName);
    await page.getByRole('button', { name: 'modifying full access members' }).click();
    await searchUserInDelegationPanels(page, userName, fullName);


     //done
    await context.close();
    mainWindow.webContents.send('automation-complete', 'Automation completed successfully!');
  } catch (error) {
    mainWindow.webContents.send('automation-error', `Error: ${error.message}`);
  }
});

function extractUserName(email) {
  const localPart = email.split('@')[0];
  return localPart.replace(/\d+$/, '');
}

function toFullName(userName) {
  const [first, last] = userName.split('.');
  return `${first.charAt(0).toUpperCase() + first.slice(1)} ${last.charAt(0).toUpperCase() + last.slice(1)}`;
}

//Search user in Send as and Read and Manage (Full access) sections under delegation panel.
async function searchUserInDelegationPanels(page, userName, fullName) {
    await page.getByRole('menuitem', { name: 'Add members' }).click();
    const panelBody = page.locator('#AddPermissions_fp_body');
    const searchBox = panelBody.locator('.ms-SearchBox-field');
    await searchBox.click();
    await searchBox.fill(userName);
    await page.getByRole('row').filter({ hasText: fullName }).getByLabel('').click();
    await page.getByRole('button', { name: 'Save' }).click();
    await page.getByRole('button', { name: 'Confirm' }).click();
    await page.getByRole('dialog').filter({ hasText: 'BackCloseMailbox' }).getByLabel('Back').click();
}
