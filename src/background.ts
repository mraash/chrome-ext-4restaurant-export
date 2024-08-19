const CONTEXT_MENU_ID = 'export_to_complex_excel'

chrome.contextMenus.create({
    id: CONTEXT_MENU_ID,
    title: 'Export to complex Excel',
    contexts: ['page'],
});

chrome.contextMenus.onClicked.addListener((info, tab) => {
    if (info.menuItemId === CONTEXT_MENU_ID) {
        console.log(info);
        chrome.tabs.sendMessage(tab!.id!, { action: 'ACTION_EXPORT' });
    }
});
