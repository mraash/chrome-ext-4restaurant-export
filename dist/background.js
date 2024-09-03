chrome.contextMenus.create({id:"export_to_complex_excel",title:"Export to complex Excel",contexts:["page"]});chrome.contextMenus.onClicked.addListener((e,t)=>{e.menuItemId==="export_to_complex_excel"&&chrome.tabs.sendMessage(t.id,{action:"ACTION_EXPORT"})});
//# sourceMappingURL=background.js.map
