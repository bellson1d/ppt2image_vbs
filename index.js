const express = require("express");
const https = require("https"); // or 'https' for https:// URLs
const path = require("path");
const axios = require('axios')
const bodyParser = require("body-parser");
const { exec } = require("child_process");
const fs = require("fs");
const cookieParser = require("cookie-parser");

const app = express();
const port = 3000;

app.use(bodyParser.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());

// 模拟文件下载地址
const tempFile = "temp_";
// 文件暂存地址
const fileOutput = path.resolve(__dirname + "/output");

//检测是否存在某个文件
const existFile = (res, pathName, originFileName = "") => {
    console.log(originFileName);
    if (!fs.existsSync(pathName)) {
        if (originFileName) {
            // 某一步骤失败，删除相关临时文件
            exec(`rm ${originFileName}*`, (error, stdout, stderr) => {
                if (!error && !stderr) {
                    console.log(`Clean ${originFileName} etc files success!`);
                } else {
                    console.log(`Clean ${originFileName} etc files failed!`);
                }
            });
        }

        res.send({
            code: 500,
            message: "Server Error! Please contact admin. " + pathName,
        });
        return false;
    }
    return true;
};

app.post("/v2/convert/pptx", async function (req, res) {
    const {
        fileUrl,
        fileName,
        detailId,
        urlOrigin = "https://drive-dev.felo.me",
    } = req.body;

    if (!fileName) {
        res.send({ code: 500, message: "No fileName" });
        return;
    } else if (!/\.pptx?$/.test(fileName)) {
        res.send({ code: 500, message: "File type is not ppt(x)!" });
        return;
    } else if (!fileUrl) {
        res.send({ code: 500, message: "No fileUrl" });
        return;
    } else if (!detailId) {
        res.send({ code: 500, message: "No detailId" });
        return;
    }
    res.send({ code: 200, message: "File processing" });
    console.log("fileName", fileName);

    const fileNameTemp = `${tempFile}${detailId}.pptx`;
    // 不包含文件后缀的名称
    const nameWithoutTail = fileNameTemp.replace(/\.pptx?/, "");
    // images 文件地址
    const imagesPath = fileOutput + "/" + nameWithoutTail;

    // 创建对应文件夹
    if (!fs.existsSync(imagesPath)) {
        fs.mkdirSync(imagesPath)
    }

    // pptx 文件地址
    const filePath = imagesPath + "/" + fileNameTemp;
    const file = fs.createWriteStream(filePath);

    console.log("fileUrl: ", fileUrl);
    // 从 AWS S3 获取 pptx 文件
    https.get(fileUrl, function (response, error) {
        if (error) {
            console.log(error);
            res.send({ code: 500, message: "Fetch pptx file error!" });
        }

        response.pipe(file);

        file.on("finish", () => {
            file.close();
            console.log("Download Completed");

            // 转换成 image 的文件名
            const newImageFileName = imagesPath + "/" + nameWithoutTail + "-0001.jpg";
            // 将 pptx 转换成 image 文件
            exec(`cscript ./vbs/ppt2image.vbs ${filePath}`, (error, stdout, stderr) => {
                if (error) {
                    console.log(`error: ${error.message}`);
                    existFile(res, newImageFileName, nameWithoutTail);
                }
                if (stderr) {
                    console.log(`stderr: ${stderr}`);
                    existFile(res, newImageFileName, nameWithoutTail);
                }
                console.log(`Convert to JPG success: ${stdout}`);

                const afterZipCallback = () => {
                    // PPT创建成功回调
                    const cbUrl = urlOrigin + `/api/v4/open/detail/convert/ppt`;
                    console.log("cbUrl", cbUrl);

                    axios
                        .post(cbUrl, { DetailId: Number(detailId) })
                        .then((res) => {
                            console.log(`statusCode: ${res.status}`);
                            console.log(res.data);
                        })
                        .catch((error) => {
                            console.log(error);
                            res.send({
                                code: 500,
                                message: "Callback request failed!",
                            });
                        });
                }

                // 将多个 jpg 转换成一个 zip 文件
                const directoryPath = fileOutput + '/' + nameWithoutTail;
                // 转换成 zip 的文件名
                const zipFileName = directoryPath + ".zip";
                const zipName = nameWithoutTail + '.zip'

                exec(
                    `cd ${fileOutput} && tar.exe -cf  ${zipName} ${nameWithoutTail}`,
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
                        afterZipCallback();
                    }
                )
            });
        });
    });
});

app.get("/v2/download/zip", function ({ query, body }, res) {
    const { detailId } = query;

    if (!detailId) res.send({ code: 500, message: "No detailId" });

    // 不包含文件后缀的名称
    const nameWithoutTail = `${tempFile}${detailId}`;
    const fileNameTemp = `${nameWithoutTail}.zip`;
    const filePath = fileOutput + "/" + fileNameTemp

    if (existFile(res, filePath, nameWithoutTail)) {
        const options = {
            root: fileOutput,
        };

        res.sendFile(fileNameTemp, options, function (err) {
            if (err) {
                console.log(err);
                res.send({ code: 500, message: "SendFile failed!" });
            } else {
                console.log("Sent:", nameWithoutTail + ".zip");

                // 返回成功删除 output 文件夹内容
                fs.rmSync(filePath, { recursive: true, force: true });
            }
        });
    }
});


app.get("*", function ({ query, body }, res) {
    res.send("Opps! No request handler!");
});

app.listen(port, () => {
    console.log(`Example app listening on port ${port}`);
});
