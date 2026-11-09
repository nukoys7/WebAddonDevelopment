//alert("Web Document Preview is working!");
console.log("WebDocumentPreview content.js loaded");

const supportedExtensions =[
    ".docx",
    ".pptx",
    ".xlsx",
    ".doc"
];

document.addEventListener("click",(event)=> {
    const link = event.target.closest("a[href]");

    if(!link) return;

    let fileUrl;

    try{
        fileUrl = new URL(link.href);
    }catch(error){
        console.error("URLの解析に失敗", error);
        return;
    }

    const pathname = fileUrl.pathname.toLowerCase();

    const isDirectFile = supportedExtensions.some(
        extension => pathname.endsWith(extension)
    );

    const isMoodleResource =
        pathname === "/mod/resource/view.php" &&
        fileUrl.searchParams.has("id");

    if(!isDirectFile && !isMoodleResource) return;

    console.log("リンクURL:", fileUrl.href);
    console.log("通常ファイル:", isDirectFile);
    console.log("Moodleリソース:", isMoodleResource);
    // console.log("リンク全体:", fileUrl.href);
    // console.log("パス部分:", fileUrl.pathname);
    // console.log("クエリ部分:", fileUrl.search);
    // console.log("取得したファイル名:", pathname.split("/").pop());

    // const isSupported = supportedExtensions.some(
    //     (extension) => pathname.endsWith(extension)
    // );

    console.log("対象ファイル", isDirectFile || isMoodleResource);

    //if(!isSupported) return;

    event.preventDefault();

    console.log("プレビュー対象を検出しました");

    //const viewerUrl = chrome.runtime.getURL("viewer.html");

    chrome.runtime.sendMessage({
        action: "openViewer",
        //url: viewerUrl,
        fileUrl: fileUrl.href
    });

    //console.log("プレビュー画面", viewerUrl.href);

    // window.open(viewerUrl.href, "_blank");
},true);