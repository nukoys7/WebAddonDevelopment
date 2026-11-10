//const statusElement = document.getElementById('status');
//const previewElement = document.getElementById('preview');
console.log("viewer started");
console.log("Viewer started");

console.log(
    "Mammoth.js読み込み確認:",
    typeof mammoth
);

const params = new URLSearchParams(window.location.search);
const fileUrl = params.get('url');

const status = document.getElementById('status');
const documentFrame = document.getElementById('documentFrame');
// 
// document.body.appendChild(status);

async function loadFile(){
    try{
        if(!fileUrl){
            throw new Error('URLが指定されていません');
        }

        status.textContent = 'ファイルを取得しています...';

        const response = await fetch(fileUrl, {
            credentials: "include"
        });

        if(!response.ok){
            throw new Error('ファイルの取得に失敗しました: ' + response.status);
        }

        const contentType = (
            response.headers.get("content-type") || ""
        ).split(",")[0].trim().toLowerCase();
            

        const finalUrl = new URL(response.url);

        const filename = decodeURIComponent(finalUrl.pathname.split("/").pop());

        const isDocx =
            filename.toLowerCase().endsWith(".docx") ||
            contentType === 
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

        const isDoc = 
            filename.toLowerCase().endsWith(".doc") || 
            contentType === "application/msword";

        console.log("取得ファイル名:", filename);
        console.log("取得MIMEタイプ:", contentType);
        console.log("DOCX判定:", isDocx);
        console.log("DOC判定:", isDoc);

        const arrayBuffer = await response.arrayBuffer();

        let html;

        if(isDocx){

            if(typeof mammoth === "undifined"){
                throw new Error("Mammoth.jsが読み込まれていません");
            }

            status.textContent = "DOCXを変換しています...";

            const result = await mammoth.convertToHtml({
                arrayBuffer: arrayBuffer
            });

            console.log("変換時の警告:", result.messages);

            html = result.value;

        } else if(isDoc){

            if(typeof docToText === "undifined"){
                throw new Error("JSDocが読み込まれていません");
            }

            status.textContent = "DOCを変換しています...";

            const text = docToText(arrayBuffer);

            if(text === null){
                throw new Error( "このDOCファイルは解析できませんでした。");
            }

            const escapedText = text
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;");

            html = `
                <div style="
                    white-space: pre-wrap;
                    font-family: sans-serif;
                    line-height: 1.8;
                    overflow-wrap: anywhere;
                ">${escapedText}</div>
            `;

        } else {

            throw new Error(
                 `未対応のファイル形式です。\nファイル名: ${filename}\n形式: ${contentType}`
            );
        }

        

        //console.log("変換時の警告:", result.messages);

        const previewHtml = `
            <!DOCTYPE html>
            <html lang="ja">
            <head>
                <meta charset="UTF-8">
                <style>
                    body {
                        max-width: 900px;
                        margin: 40px auto;
                        padding: 40px;
                        font-family: sans-serif;
                        line-height: 1.8;
                        color: #222;
                        background: white;
                    }

                    img {
                        max-width: 100%;
                        height: auto;
                    }

                    table {
                        border-collapse: collapse;
                        max-width: 100%;
                    }

                    td, th {
                        border: 1px solid #aaa;
                        padding: 8px;
                    }
                </style>
            </head>
            <body>
                ${html}
            </body>
            </html>
        `;

        documentFrame.srcdoc = previewHtml;

        status.textContent =
            `表示完了: ${filename}`;

        // info.textContent = 
        //     `ファイルURL: ${response.url}\n` +
        //     `ファイル形式: ${blob.type || '不明'}\n` +
        //     `ファイルサイズ: ${blob.size} バイト\n`;

        // document.body.appendChild(info);

        // console.log("ファイル取得成功:", blob);
    }catch(error){

        console.error(error);
        status.textContent = 
            'エラー: ' + error.message;

        
    }
}

loadFile();