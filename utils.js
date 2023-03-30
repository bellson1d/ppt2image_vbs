// 获取当前时间和指定时间差
export const timeDuration = (startTime) => {
    return (new Date().getTime() - startTime) / 1000
}
