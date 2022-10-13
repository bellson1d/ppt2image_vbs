const fs = require('fs');
const { exec } = require("child_process");

const zipFile = ({ fileOutput, name = 'example', cb }) => {
    const directoryPath = fileOutput + '/' + nameWithoutTail;
    exec(
        `tar.exe -a -c -f ${directoryPath + '.zip'} ${directoryPath}`,
        (error, stdout, stderr) => {
            if (error) {
                console.log(`error: ${error.message}`);
                existFile(res, zipFileName, nameWithoutTail);
            }
            if (stderr) {
                console.log(`stderr: ${stderr}`);
                existFile(res, zipFileName, nameWithoutTail);
            }
            console.log(`Convert to Zip: ${stdout}`);

            fs.rmSync(directoryPath, { recursive: true, force: true });
        }
    )
}

module.exports = { zipFile }