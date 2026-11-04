// Open the side panel when the toolbar icon is clicked. Nothing else runs here:
// all SharePoint requests are made from inside the SharePoint tab.
chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch(() => {});
