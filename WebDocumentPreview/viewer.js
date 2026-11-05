const statusElement = document.getElementById('status');
const previewElement = document.getElementById('preview');

const params = new URLSearchParams(window.location.search);
const fileUrl = params.get('url');

const status = document.createElement('p');
status.textContent = 'ファイルを取得しています...';
document.body.appendChild(status);

async function loadFile(){
    try{
        if(!fileUrl){
            throw new Error('URLが指定されていません');
        }

        const response = await fetch(fileUrl, {
            credentials: "include"
        });

        if(!response.ok){
            throw new Error('ファイルの取得に失敗しました: ' + response.status);
        }

        const blob = await response.blob();

        status.textContent = 'ファイルを取得しました。プレビューを表示します...';

        const info = document.createElement('pre');

        info.textContent = 
            `ファイルURL: ${response.url}\n` +
            `ファイル形式: ${blob.type || '不明'}\n` +
            `ファイルサイズ: ${blob.size} バイト\n`;

        document.body.appendChild(info);

        console.log("ファイル取得成功:", blob);
    }catch(error){
        status.textContent = 
            'エラー: ' + error.message;

        console.error(error);
    }
}

loadFile();