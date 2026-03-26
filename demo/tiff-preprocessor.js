async function preprocessTiff(blob) {
    const files = fflate.unzipSync(new Uint8Array(await blob.arrayBuffer()));
    const tiffEntries = Object.entries(files).filter(([fileName]) => /[.]tiff?$/i.test(fileName));

    if (tiffEntries.length == 0)
        return blob;

    for (let [fileName, buffer] of tiffEntries) {
        const tiff = new Tiff({ buffer });
        const pngBlob = await new Promise(res => tiff.toCanvas().toBlob(result => res(result), "image/png"));
        files[fileName] = new Uint8Array(await pngBlob.arrayBuffer());
    }

    return new Blob([fflate.zipSync(files)], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });
}
