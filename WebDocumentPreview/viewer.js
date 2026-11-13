//const statusElement = document.getElementById('status');
//const previewElement = document.getElementById('preview');
console.log("viewer started");
console.log("Viewer started");

console.log(
    "Mammoth.js読み込み確認:",
    typeof mammoth
);
console.log("JSZip", typeof JSZip);

const params = new URLSearchParams(window.location.search);
const fileUrl = params.get('url');

const status = document.getElementById('status');
const documentFrame = document.getElementById('documentFrame');
const pptxPreview = document.getElementById("pptxPreview");
// 
// document.body.appendChild(status);

const DRAWING_NS = 
    "http://schemas.openxmlformats.org/drawingml/2006/main";

const PRESENTATION_NS =
    "http://schemas.openxmlformats.org/presentationml/2006/main";

const REL_NS =
    "http://schemas.openxmlformats.org/package/2006/relationships";

const OFFICE_REL_NS = 
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function parseXml(xml){
    const parser = new DOMParser();
    const result = parser.parseFromString(xml, "application/xml");

    if(result.querySelector("parsererror")){
        throw new Error("XML解析に失敗しました");
    }

    return result;
}

function resolvePath(baseDir, target){
    const url = new URL(
        target,
        "https://ppax.local/" + baseDir
    );

    return url.pathname.replace(/^\/+/, "");
}

async function getRelationships(zip, path, baseDir) {
    const file = zip.file(path);//現在ここ

    if(!file){
        return {};
    }

    const xml = await file.async("string");
    const doc = parseXml(xml);

    const relationships = {};

    const elements = doc.getElementsByTagNameNS(
        REL_NS,
        "Relationship"
    );

    for(const element of elements){
        const id = element.getAttribute("Id");
        const target = element.getAttribute("Target");
        const mode = element.getAttribute("TargetMode");

        if(mode === "External"){
            continue;
        }

        relationships[id] = resolvePath(baseDir, target);
    }

    return relationships;
}

function firstDescendendantNS(element, namespace, localName){
    return element.getElementsByTagNameNS(
        namespace,
        localName
    )[0] || null;
}

function getPlaceholderKey(shape){
    const placeholder = firstDescendendantNS(
        shape,
        PRESENTATION_NS,
        "ph"
    );

    if(!placeholder) return null;

    const type = placeholder.getAttribute("type") || "obj";
    const index = placeholder.getAttribute("idx") || "0";

    return `${type}:$index`;
}

function getShapeTransform(shape){
    const shapeProperties = firstDescendendantNS(
        shpe,
        PRESENTATION_NS,
        "spPr"
    ) || firstDescendendantNS(
        shape,
        PRESENTATION_NS,
        "grpSpPr"
    );

    if(!shapeProperties) return null;

    const transform = firstDescendendantNS(
        shapeProperties,
        DRAWING_NS,
        "xfrm"
    );

    if(!transform) return null;

    const offset = firstDescendendantNS(
        transform,
        DRAWING_NS,
        "off"
    );

    const extent = firstDescendendantNS(
        transform,
        DRAWING_NS,
        "ext"
    );

    if(!offset || !extent) return null;

    return {
        x: Number(offset.getAttribute("x") || 0),
        y: Number(offset.getAttribute("y") || 0),
        cx: Number(offset.getAttribute("cx") || 0),
        cy: Number(offset.getAttribute("cy") || 0)
    };
}

async function getLayoutPlaceholderTransforms(zip, slidePath){
    const slideName = slidePath.split("/").pop();

    const relPath =
    "ppt/slides/_rels/" + slideName + ".rels";

    const relationships = await getRelationships(
        zip,
        relPath,
        "ppt/slides/"
    );

    const layoutPath = Object.values(relationships).find(path =>
        path.includes("slideLayouts/")
    );

    if(!layoutPath || !zip.file(layoutPath)){
        return {};
    }

    const layoutXml = await zip.file(layoutPath).async("string");
    const layoutDoc = parseXml(layoutXml);

    const shapes = layoutDoc.getElementsByTagNameNS(
        PRESENTATION_NS,
        "sp"
    );

    const transform = {};

    for(const shape of shapes){
        const key = getPlaceholderKey(shape);
        const transform = getShapeTransform(shape);

        if(key && transform){
            Transforms[key] = transform;
        }
    }

    return transforms;
}

async function getSlideImages(zip, slidePath) {
    const slideName = slidePath.split("/").pop();

    const relPath =
        "ppt/slides/_rels/" + slideName + ".rels";

    const relationships = await getRelationships(
        zip,
        relPath,
        "ppt/slides/"
    );

    const slideFile = await zip.file(slidePath).async("string");
    const slidexml = parseXml(slideFile);

    const pictures = slidexml.getElementsByTagNameNS(
        PRESENTATION_NS,
        "pic"
    );

    const images = [];

    for(const picture of pictures){
        const blips = picture.getElementsByTagNameNS(
            DRAWING_NS,
            "blip"
        );

        if(blips.length === 0){
            continue;
        }

        const relationshipId = blips[0].getAttributeNS(
            OFFICE_REL_NS,
            "embed"
        );

        const imagePath = relationships[relationshipId];

        if(!imagePath){
            continue;
        }

        const imageFile = zip.file(imagePath);

        if(!imageFile){
            continue;
        }

        const extension = imagePath.split(".").pop().toLowerCase();

        const mimeTypes = {
            png: "image/png",
            jpg: "image/jpeg",
            jpeg: "image/jpeg",
            gif: "image/gif",
            bmp: "image/bmp",
            webp: "image/webp"
        };

        const mime = mimeTypes[extension];

        if(!mime){
            continue;
        }

        const base64 = await imageFile.async("base64");

        images.push(
            `data:${mime};base64,${base64}`
        );
    }

    return images;
}

async function getSlideText(zip, slidePath){
    const xml = await zip.file(slidePath).async("string");
    const doc = parseXml(xml);

    const paragraphs = doc.getElementsByTagNameNS(
        DRAWING_NS,
        "p"
    );

    const texts = [];

    for(const paragraph of paragraphs){
        const textElements = paragraph.getElementsByTagNameNS(
            DRAWING_NS,
            "t"
        );

        let text = "";

        for(const element of textElements){
            text += element.textContent;
        }

        if(text.trim()){
            texts.push(text);
        }
    }

    return texts;
}

async function convertPptx(arrayBuffer) {
    const zip = await JSZip.loadAsync(arrayBuffer);

    const presentationFile = await zip
        .file("ppt/presentation.xml")
        .async("string");
    
    const presentationXml = parseXml(presentationFile);

    const presentationRels = await getRelationships(
        zip,
        "ppt/_rels/presentation.xml.rels",
        "ppt/"
    );

    const slideIds = presentationXml.getElementsByTagNameNS(
        PRESENTATION_NS,
        "sldId"
    );

    const slides = [];

    for(const slideId of slideIds){
        const relationshipId = slideId.getAttributeNS(
            OFFICE_REL_NS,
            "id"
        );

        const slidePath = presentationRels[relationshipId];

        if(!slidePath || !zip.file(slidePath)){
            continue;
        }

        const texts = await getSlideText(zip, slidePath);
        const images = await getSlideImages(zip,slidePath);

        slides.push({
            texts,
            images
        });
    }

    return slides;
}

function displaySlides(slides){
    documentFrame.hidden = true;
    pptxPreview.hidden = false;

    pptxPreview.replaceChildren();

    slides.forEach((slide, index) => {
        const slideElement = document.createElement("section");
        slideElement.className = "slide";

        const title = document.createElement("h2");
        title.textContent = `スライド ${index + 1}`;

        slideElement.appendChild(title);

        const content = document.createElement("div");
        content.className = "slide-content";
        content.textContent = slide.texts.join("\n");

        slideElement.appendChild(content);

        const imageContainer = document.createElement("div");
        imageContainer.className = "slide-images";

        for(const imageUrl of slide.images){
            const image = document.createElement("img");
            image.src = imageUrl;
            image.alt = "スライド内の画像";

            imageContainer.appendChild(image);
        }

        slideElement.appendChild(imageContainer);
        pptxPreview.appendChild(slideElement);
    });
}

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

        const lowerFilename = filename.toLowerCase();

        const isDocx =
            filename.toLowerCase().endsWith(".docx") ||
            contentType === 
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

        const isDoc = 
            filename.toLowerCase().endsWith(".doc") || 
            contentType === "application/msword";

        const isPptx = 
            lowerFilename.endsWith(".pptx") ||
            contentType ===
            "application/vnd.openxmlformats-officedocument.presentationml.presentation";

        console.log("取得ファイル名:", filename);
        console.log("取得MIMEタイプ:", contentType);
        console.log("DOCX判定:", isDocx);
        console.log("DOC判定:", isDoc);
        console.log("PPTX:", isPptx);

        const arrayBuffer = await response.arrayBuffer();

        let html;

        if(isPptx){
            if(typeof JSZip === "undefined"){
                throw new Error("JSZipが読まれていません");
            }

            status.textContent = "PowerPointを解析しています...";

            const slides = await convertPptx(arrayBuffer);

            if(slides.length === 0){
                throw new Error("slide is missing...");
            }

            displaySlides(slides);

            status.textContent = 
                `表示完了: ${filename} (${slides.length}枚)`;

            return;
        } else if(isDocx){

            if(typeof mammoth === "undefined"){
                throw new Error("Mammoth.jsが読み込まれていません");
            }

            status.textContent = "DOCXを変換しています...";

            const result = await mammoth.convertToHtml({
                arrayBuffer: arrayBuffer
            });

            console.log("変換時の警告:", result.messages);

            html = result.value;

        } else if(isDoc){

            if(typeof docToText === "undefined"){
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

        pptxPreview.hidden = true;
        documentFrame.hidden = false;
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