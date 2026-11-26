chrome.runtime.onMessage.addListener((message) => {
  if (message.action === "openViewer") {
    checkFileAndOpenViewer(message.fileUrl);
  }
});

async function checkFileAndOpenViewer(fileUrl){
  try{
    const response = await fetch(fileUrl, {
      credentials: "include"
    });

    if(!response.ok){
      throw new Error("ファイルの取得に失敗しました");
    }

    const contentType = (
      response.headers.get("content-Type") || ""
    ).split(";")[0].trim().toLowerCase();

    const finalUrl = new URL(response.url);
    const filename = decodeURIComponent(
      finalUrl.pathname.split("/").pop()
    ).toLowerCase();

    if(
      contentType === "application/pdf" ||
      filename.endsWith(".pdf")
    ){
      console.log("PDFのため通常のブラウザで開きます");

      chrome.tabs.create({
        url: fileUrl
      });
      return;
    }

    const viewerUrl = new URL(
      chrome.runtime.getURL("viewer.html")
    );

    viewerUrl.searchParams.set("url", fileUrl);

    chrome.tabs.create({
      url: viewerUrl.href
    });
  } catch(error){
    console.error("ファイル形式の確認に失敗しました:", error);
  }
}