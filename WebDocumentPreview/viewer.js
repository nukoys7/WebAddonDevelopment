//const statusElement = document.getElementById('status');
//const previewElement = document.getElementById('preview');
console.log("viewer started");
//console.log("Viewer started");

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
const downloadButton = document.getElementById("downloadButton");

let originalFile = null;
let originalFilename = "document";
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

    return `${type}:${index}`;
}

function getShapeTransform(shape){
    const shapeProperties = firstDescendendantNS(
        shape,
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
        cx: Number(extent.getAttribute("cx") || 0),
        cy: Number(extent.getAttribute("cy") || 0)
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

    const transforms = {};

    for(const shape of shapes){
        const key = getPlaceholderKey(shape);
        const shapeTransform = getShapeTransform(shape);

        if(key && shapeTransform){
            transforms[key] = shapeTransform;
        }
    }

    return transforms;
}

function getShapeText(shape){
    const textBody = firstDescendendantNS(
        shape,
        PRESENTATION_NS,
        "txBody"
    );

    if (!textBody) return "";

    const paragraphs = textBody.getElementsByTagNameNS(
        DRAWING_NS,
        "p"
    );

    const lines = [];

    for (const paragraph of paragraphs){
        const runs = paragraph.getElementsByTagNameNS(
            DRAWING_NS,
            "t"
        );

        let line = "";

        for (const run of runs){
            line += run.textContent || "";
        }

        lines.push(line);
    }

    return lines.join("\n").trim();
}

function getShapeFontSize(shape){
    const runProperties = shape.getElementsByTagNameNS(
        DRAWING_NS,
        "rPr"
    );

    if(runProperties.length > 0){
        const size = Number(
            runProperties[0].getAttribute("sz")
        );

        if(size > 0) return size / 100;
    }

    const endProperties = shape.getElementsByTagNameNS(
        DRAWING_NS,
        "endParaRPr"
    );

    if(endProperties.length > 0){
        const size = Number(
            endProperties[0].getAttribute("sz")
        );

        if(size > 0) return size / 100;
    }

    return 18;
}

function getShapeColor(shape){
    const solidFill = firstDescendendantNS(
        shape,
        DRAWING_NS,
        "srgbClr"
    );

    if(solidFill){
        const value = solidFill.getAttribute("val");

        if(/^[0-9a-f]{6}$/i.test(value)){
            return `#${value}`;
        }
    }

    return "#222";
}

async function getSlideObjects(zip, slidePath){
    const slideName = slidePath.split("/").pop();

    const relationships = await getRelationships(
        zip,
        `ppt/slides/_rels/${slideName}.rels`,
        "ppt/slides/"
    );

    const layoutTransforms = 
        await getLayoutPlaceholderTransforms(
            zip,
            slidePath
        );
    
    const slideXml = await zip.file(slidePath).async("string");
    const slideDoc = parseXml(slideXml);

    const shapeTree = firstDescendendantNS(
        slideDoc,
        PRESENTATION_NS,
        "spTree"
    );

    if(!shapeTree) return [];

    const objects = [];

    for(const node of Array.from(shapeTree.children)){
        const localName = node.localName;

        if(localName !== "sp"  && localName !== "pic"){
            continue;
        }

        let transform = getShapeTransform(node);

        if(!transform){
            const key = getPlaceholderKey(node);

            if(key){
                transform = layoutTransforms[key] || null;
            }
        }

        if(!transform || !transform.cx || !transform.cy){
            continue;
        }

        if(localName === "sp"){
            const text = getShapeText(node);

            if(!text) continue;

            objects.push({
                type: "text",
                ...transform,
                text,
                fontSize: getShapeFontSize(node),
                color: getShapeColor(node)
            });
        }

        if(localName === "pic"){
            const blip = firstDescendendantNS(
                node,
                DRAWING_NS,
                "blip"
            );

            if(!blip) continue;

            const relationshipId = blip.getAttributeNS(
                OFFICE_REL_NS,
                "embed"
            );

            const imagePath = relationships[relationshipId];

            if(!imagePath) continue;

            const imageFile = zip.file(imagePath);

            if(!imageFile) continue;

            const extension = imagePath
                .split(".")
                .pop()
                .toLowerCase();

            const mimeTypes = {
                png: "image/png",
                jpg: "image/jpeg",
                jpeg: "image/jpeg",
                gif: "image/gif",
                bmp: "image/bmp",
                webp: "image/webp",
                svg: "image/svg+xml"
            };

            const mime = mimeTypes[extension];

            if(!mime){
                console.warn(
                   "未対応の画像形式:",
                    extension,
                    imagePath  
                );

                continue;
            }

            const base64 = await imageFile.async("base64");

            objects.push({
                type: "image",
                ...transform,
                src: `data:${mime};base64,${base64}`
            });
        }
    }

    return objects;
}

// async function getSlideImages(zip, slidePath) {
//     const slideName = slidePath.split("/").pop();

//     const relPath =
//         "ppt/slides/_rels/" + slideName + ".rels";

//     const relationships = await getRelationships(
//         zip,
//         relPath,
//         "ppt/slides/"
//     );

//     const slideFile = await zip.file(slidePath).async("string");
//     const slidexml = parseXml(slideFile);

//     const pictures = slidexml.getElementsByTagNameNS(
//         PRESENTATION_NS,
//         "pic"
//     );

//     const images = [];

//     for(const picture of pictures){
//         const blips = picture.getElementsByTagNameNS(
//             DRAWING_NS,
//             "blip"
//         );

//         if(blips.length === 0){
//             continue;
//         }

//         const relationshipId = blips[0].getAttributeNS(
//             OFFICE_REL_NS,
//             "embed"
//         );

//         const imagePath = relationships[relationshipId];

//         if(!imagePath){
//             continue;
//         }

//         const imageFile = zip.file(imagePath);

//         if(!imageFile){
//             continue;
//         }

//         const extension = imagePath.split(".").pop().toLowerCase();

//         const mimeTypes = {
//             png: "image/png",
//             jpg: "image/jpeg",
//             jpeg: "image/jpeg",
//             gif: "image/gif",
//             bmp: "image/bmp",
//             webp: "image/webp"
//         };

//         const mime = mimeTypes[extension];

//         if(!mime){
//             continue;
//         }

//         const base64 = await imageFile.async("base64");

//         images.push(
//             `data:${mime};base64,${base64}`
//         );
//     }

//     return images;
// }

// async function getSlideText(zip, slidePath){
//     const xml = await zip.file(slidePath).async("string");
//     const doc = parseXml(xml);

//     const paragraphs = doc.getElementsByTagNameNS(
//         DRAWING_NS,
//         "p"
//     );

//     const texts = [];

//     for(const paragraph of paragraphs){
//         const textElements = paragraph.getElementsByTagNameNS(
//             DRAWING_NS,
//             "t"
//         );

//         let text = "";

//         for(const element of textElements){
//             text += element.textContent;
//         }

//         if(text.trim()){
//             texts.push(text);
//         }
//     }

//     return texts;
// }

async function convertPptx(arrayBuffer) {
    const zip = await JSZip.loadAsync(arrayBuffer);

    const presentationFile = zip.file(
        "ppt/presentation.xml"
    );
    
    if(!presentationFile){
        throw new Error(
            "Powerpointの構成ファイルが見つかりません"
        );
    }

    const presentationXml =
        await presentationFile.async("string");

    const presentationDoc = parseXml(presentationXml);

    const slideSize = firstDescendendantNS(
        presentationDoc,
        PRESENTATION_NS,
        "sldSz"
    );

    if(!slideSize){
        throw new Error(
            "スライドサイズが取得できません"
        );
    }

    const dimensions = {
        width: Number(slideSize.getAttribute("cx")),
        height: Number(slideSize.getAttribute("cy"))
    };

    const presentationRels = await getRelationships(
        zip,
        "ppt/_rels/presentation.xml.rels",
        "ppt/"
    );

    const slideIds = presentationDoc.getElementsByTagNameNS(
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

        const objects = await getSlideObjects(
            zip,
            slidePath
        );

        slides.push({
            objects
        });

    }

    return {
        slides,
        dimensions
    };
}

function displaySlides(slides, dimensions){
    documentFrame.hidden = true;
    pptxPreview.hidden = false;

    pptxPreview.replaceChildren();

    slides.forEach((slide, index) => {
        const slideElement = document.createElement("section");
        slideElement.className = "slide";

        const title = document.createElement("h2");
        title.textContent = `スライド ${index + 1}`;

        slideElement.appendChild(title);

        const canvas = document.createElement("div");
        canvas.className = "slide-canvas";

        canvas.style.aspectRatio = 
            `${dimensions.width} / ${dimensions.height}`;

        for(const object of slide.objects){
            const element = document.createElement(
                object.type === "image" ? "img" : "div"
            );

            element.classList.add("slide-object");

            element.style.left =
                `${object.x / dimensions.width * 100}%`;

            element.style.top =
                `${object.y / dimensions.height * 100}%`;

            element.style.width =
                `${object.cx / dimensions.width * 100}%`;

            element.style.height =
                `${object.cy / dimensions.height * 100}%`;

            if(object.type === "text"){
                element.classList.add("slide-text");

                element.textContent = object.text;

                element.style.color = object.color;
            } else{
                element.src = object.src;
                element.alt = "スライド内の画像";
            }

            canvas.appendChild(element);
            
        }

        slideElement.appendChild(canvas);
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
        ).split(";")[0].trim().toLowerCase();
            

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

        const isPdf =
            lowerFilename.endsWith(".pdf") ||
            contentType === "application/pdf";

        if(isPdf){
            throw new Error("PDFファイルはプレビュー対象外です.");
        }

        console.log("取得ファイル名:", filename);
        console.log("取得MIMEタイプ:", contentType);
        console.log("DOCX判定:", isDocx);
        console.log("DOC判定:", isDoc);
        console.log("PPTX:", isPptx);

        const arrayBuffer = await response.arrayBuffer();

        originalFile = new Blob([arrayBuffer], {
            type: contentType || "application/octet-stream"
        });

        originalFilename = filename;
        downloadButton.disabled = false;

        let html;

        if(isPptx){
            if(typeof JSZip === "undefined"){
                throw new Error("JSZipが読まれていません");
            }

            status.textContent = "PowerPointを解析しています...";

            const pptxData = await convertPptx(arrayBuffer);

            const slides = pptxData.slides;

            if(slides.length === 0){
                throw new Error("slide is missing...");
            }

            displaySlides(
                slides,
                pptxData.dimensions
            );

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

downloadButton.addEventListener("click", () => {
    if(!originalFile) return;

    const downloadUrl = URL.createObjectURL(originalFile);

    const link = document.createElement("a");
    link.href = downloadUrl;
    link.download = originalFilename;

    document.body.appendChild(link);
    link.click();
    link.remove();

    setTimeout(() => {
        URL.revokeObjectURL(downloadUrl);
    }, 1000);
});

loadFile();