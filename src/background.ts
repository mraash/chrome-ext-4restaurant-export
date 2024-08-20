chrome.contextMenus.create({
    id: 'export_to_complex_excel',
    title: 'Export to complex Excel',
    contexts: ['page'],
});

chrome.contextMenus.onClicked.addListener((info, tab) => {
    if (info.menuItemId === 'export_to_complex_excel') {
        console.log(info);
        chrome.tabs.sendMessage(tab!.id!, { action: 'ACTION_EXPORT' });
    }
});
