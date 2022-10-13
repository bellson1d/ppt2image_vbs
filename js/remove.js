const findRemoveSync = require('find-remove')
const path = require("path");

findRemoveSync(path.resolve(__dirname, '../output'), { files: '121.*' })

// Get-ChildItem 'E:\LearnProjects\ppt2image_vbs\output' | Where{$_.Name -Match "121"} | Remove-Item
