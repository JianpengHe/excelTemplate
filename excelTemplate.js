const zlib = require('zlib');
const fs = require('fs');
const xlsxPath = process.argv[2], fnName = process.argv[3];
if (!xlsxPath || !fnName) {
    console.log("请带上2个参数，分别是Excel文件地址和导出的函数名");
    console.log("如", `node "${process.argv[1]}" example.xlsx ObjToExcel`);
    process.exit();
}
const files = [];
const zipParse = (ReadStream, zipFileCallback, finishCallback = console.log, bufListMaxSize = 30 * 1024 * 1024) => {
    let bufList = [], bufListSize = 0, size = 0, userDestroyed = false;
    ReadStream.on("data", chunck => {
        bufList.push(chunck);
        bufListSize += chunck.length;
        size += chunck.length;
        if (bufListSize > bufListMaxSize) {
            ReadStream.pause();
            bufList = [zipParse.readZip(Buffer.concat(bufList), zipFileCallback, 1)];
            if (bufList[0] === false) {
                ReadStream.destroy();
                userDestroyed = true;
                return;
            }
            bufListSize = bufList[0].length;
            ReadStream.resume();
        }
    });
    ReadStream.on("error", finishCallback);
    ReadStream.on("close", () => {
        finishCallback(null, size, !userDestroyed && zipParse.readZip(Buffer.concat(bufList), zipFileCallback, 1));
    });
};
zipParse.readZip = (buf, show) => {
    if (buf.length < 30) { return buf; }
    const p = buf.indexOf(zipParse.zipHead);
    if (p < 0 || buf.length - p < 30) { return buf; }
    const fileNameLen = buf.readUInt16LE(p + 26), compressedSize = buf.readUInt32LE(p + 18), compressedStart = p + 30 + fileNameLen + buf.readUInt16LE(p + 28);
    if (buf.length < compressedSize + compressedStart) { return buf; }
    files.push({
        compressedBuffer: buf.slice(compressedStart, compressedStart + compressedSize),
        compressedSize,
        uncompressedSize: buf.readUInt32LE(p + 22),
        fileName: buf.slice(p + 30, p + 30 + fileNameLen).toString(),
        compressionMethod: buf.readUInt16LE(p + 8),
        compressedFileBuffer: buf.slice(p, compressedStart + compressedSize),
        fileNameLen,
        compressedStart
    });
    return zipParse.readZip(buf.slice(compressedStart + compressedSize));
}
zipParse.zipHead = Buffer.from([0x50, 0x4b, 0x03, 0x04]);

const zipFileCallback = async info => {
    console.log(info);
};
let listStart = 0;
const list1 = Buffer.from([0x50, 0x4b, 1, 2, 0x2d, 0]);
zipParse(fs.createReadStream(xlsxPath), zipFileCallback, (err, size, buf) => {
    let byte = 0, sharedStrings = {}, list = [],
        o = files.map(a => {
            if (a.fileName.includes("sharedStrings")) {
                sharedStrings = a;
                return Buffer.alloc(0)
            }
            list.push(list1);
            list.push(a.compressedFileBuffer.slice(4, 28));
            list.push(Buffer.alloc(12, 0));
            const start = Buffer.alloc(4);
            start.writeInt32LE(byte);
            list.push(start);
            list.push(a.compressedFileBuffer.slice(30, a.fileNameLen + 30));
            byte += a.compressedFileBuffer.length;
            return a.compressedFileBuffer;
        });
    o.push((a => {
        fn = fn.replace("$text", "'" + zlib.inflateRawSync(a.compressedBuffer).toString().replace(/\r|\n/g, "") + "'");
        list.push(list1);
        a.compressedFileBuffer[8] = 0;
        a.compressedFileBuffer.writeUInt32LE(0, 14);
        list.push(a.compressedFileBuffer.slice(4, 18));
        list = Buffer.concat(list);
        const partB = a.compressedFileBuffer.slice(26, a.compressedStart);
        listStart += partB.length + 8;
        update("$partB", partB);
        update("$partC", list);
        list = [];
        list.push(a.compressedFileBuffer.slice(26, 30));
        list.push(Buffer.alloc(10, 0));
        const start = Buffer.alloc(4);
        start.writeInt32LE(byte);
        list.push(start);
        list.push(a.compressedFileBuffer.slice(30, a.fileNameLen + 30));
        return a.compressedFileBuffer.slice(0, 18);
    })(sharedStrings));
    o = Buffer.concat(o);
    listStart += o.length;
    fn = fn.replace("$start", listStart);
    update("$partA", o);
    list = Buffer.concat(list);
    buf = buf.slice(buf.indexOf(Buffer.from([0x50, 0x4b, 5, 6])), buf.length - 6);
    update("$partD", Buffer.concat([list, buf]));
    fs.writeFileSync(fnName + ".js", `(${fn})("${fnName}");`);
    console.log("完成", "在同目录", fnName + ".js", "浏览器引入后调用", fnName + "({});即可使用");
});


let fn = (function (b) { if (!window.atob || !window.TextEncoder) { alert("您的浏览器版本过旧，建议使用最新版的谷歌浏览器"); return } var d = function (r, t) { var s = m; for (var q in r) { s = s.replace("{{" + q + "}}", r[q]) } s = new TextEncoder().encode(s); var e = [j, new Uint32Array([s.length]), new Uint32Array([s.length]), h, s, g, new Uint32Array([s.length]), new Uint32Array([s.length]), f, new Uint32Array([a + s.length]), new Uint16Array([0])]; if (t) { var p = window.URL.createObjectURL(new Blob(e)), n = document.createElement("a"); document.body.appendChild(n); n.href = p; n.download = t; n.click(); window.URL.revokeObjectURL(p) } return e }, c = function (p) { var o = window.atob((p + "====".substr(0, (4 - p.length % 4) % 4)).replace(/\-/g, "+").replace(/_/g, "/")), n = new Uint8Array(o.length); for (var e = 0; e < o.length; ++e) { n[e] = o.charCodeAt(e) } return n }; var j = c($partA), m = $text, h = c($partB), g = c($partC), f = c($partD), a = $start; window[b] = d }).toString();
const update = (name, buf) => { fn = fn.replace(name, "'" + buf.toString("base64") + "'"); };