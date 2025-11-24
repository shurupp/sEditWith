/**
 * openLinkedImage.jsx
 * Professional, concise InDesign → Photoshop / Illustrator opener
 * No guides, no alerts unless something is wrong.
 * Select any graphic frame (or the image itself) and run.
 */
(function () {
    /* ---------- helpers ---------- */
    const RASTER = ['.psd', '.jpg', '.jpeg', '.tif', '.tiff', '.png'];
    const VECTOR = ['.ai', '.eps', '.svg'];

    const extOf = p => File(p).name.toLowerCase().slice(-4);

    /* ---------- validation ---------- */
    if (app.selection.length !== 1) return;          // silent exit
    const sel = app.selection[0];

    const img = sel.constructor.name === 'Image'
        ? sel
        : (sel.allGraphics && sel.allGraphics.length)
            ? sel.allGraphics[0]
            : null;

    if (!img || !img.itemLink) {
        alert('Select a frame that contains a linked image.');
        return;
    }

    /* ---------- routing ---------- */
    const path = img.itemLink.filePath;
    const ext  = extOf(path);
    const host = RASTER.includes(ext) ? 'photoshop'
               : VECTOR.includes(ext) ? 'illustrator'
               : null;

    if (!host) {
        alert('Unsupported file type: ' + ext);
        return;
    }

    /* ---------- open ---------- */
    const bt = new BridgeTalk();
    bt.target = host;
    bt.body   = 'app.open(new File("' + path.replace(/\\/g, '/') + '"));';
    bt.onError = e => $.writeln('BridgeTalk error: ' + e.body);
    bt.send();
})();