//const statusElement = document.getElementById('status');
//const previewElement = document.getElementById('preview');
console.log("viewer started");

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

        const contentType =
            response.headers.get("content-type") || "";

        const finalUrl = new URL(response.url);

        const filename = decodeURIComponent(finalUrl.pathname.split("/").pop());

        const isDocx =
            filename.toLowerCase().endsWith(".docx") ||
            contentType.includes("application/vnd.openxmlformats-officedocument.wordprocessingml.document");

        if(isDocx){
            throw new Error(
                `このファイル形式にはまだ対応していません。\nファイル名: ${filename}\n形式: ${contentType}`
            )
        }

        status.textContent = "Word文書を変換しています...";

        const arrayBuffer = await response.arrayBuffer();

        const result = await mammoth.convertToHtml({
            arrayBuffer: arrayBuffer
        });

        console.log("変換時の警告:", result.messages);

        const html = `
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
                ${result.value}
            </body>
            </html>
        `;

        documentFrame.srcdoc = html;

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