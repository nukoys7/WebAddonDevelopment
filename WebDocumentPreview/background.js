chrome.runtime.onMessage.addListener((message) => {
  if (message.action === "openViewer") {
    const viewerUrl = new URL(chrome.runtime.getURL("viewer.html"));

    viewerUrl.searchParams.set("url", message.fileUrl);

    chrome.tabs.create({
        url:viewerUrl.href
    });
  }
});