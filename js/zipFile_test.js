const fs = require('fs');
const archiver = require('archiver');

const zipFile = ({ filePath, name = 'example', cb }) => {
    const output = fs.createWriteStream(filePath + '/' + name + '.zip');
    const archive = archiver('zip', {
        zlib: { level: 9 } // Sets the compression level.
    });

    output.on('close', function () {
        console.log(archive.pointer() + ' total bytes');
        console.log('archiver has been finalized and the output file descriptor has closed.');
        // 返回成功删除 output 文件夹内容
        // fs.rmSync(filePath + "/" + name, { recursive: true, force: true });
        cb && cb()
    });

    archive.on('error', function (err) {
        throw err;
    });

    archive.pipe(output);
    archive.directory(filePath + name, false);

    archive.finalize();
}

module.exports = { zipFile }