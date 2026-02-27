const fs = require('fs');
const path = require('path');

/**
 * 自动清理指定目录下超过指定天数的文件
 * @param {string} dirPath - 要清理的目录路径
 * @param {number} daysToKeep - 保留的天数
 */
function cleanOldFiles(dirPath, daysToKeep = 30) {
    if (!fs.existsSync(dirPath)) {
        return;
    }

    const now = Date.now();
    const maxAgeMs = daysToKeep * 24 * 60 * 60 * 1000;

    try {
        const files = fs.readdirSync(dirPath);

        files.forEach(file => {
            const filePath = path.join(dirPath, file);
            const stats = fs.statSync(filePath);

            if (stats.isDirectory()) {
                // 递归清理子目录
                cleanOldFiles(filePath, daysToKeep);
                
                // 如果子目录为空了，可以考虑删除子目录，这里为了防止误删先不删外层基于日期的目录
                try {
                    if (fs.readdirSync(filePath).length === 0) {
                        fs.rmdirSync(filePath);
                        console.log(`[清理垃圾] 已删除空目录: ${filePath}`);
                    }
                } catch (e) {
                    // 忽略删除目录错误
                }
            } else {
                // 检查文件修改时间
                const ageMs = now - stats.mtimeMs;
                if (ageMs > maxAgeMs) {
                    try {
                        fs.unlinkSync(filePath);
                        console.log(`[清理垃圾] 已删除过期文件 (${Math.round(ageMs / 86400000)}天): ${filePath}`);
                    } catch (err) {
                        console.error(`[清理垃圾] 无法删除文件 ${filePath}:`, err.message);
                    }
                }
            }
        });
    } catch (error) {
        console.error(`[清理垃圾] 读取目录 ${dirPath} 失败:`, error.message);
    }
}

/**
 * 每天凌晨自动运行清理任务
 * @param {string} baseDataDir - 数据根目录
 * @param {number} days - 保留天数
 */
function startDailyCleanup(baseDataDir, days = 30) {
    // 启动时先执行一次
    console.log(`[清理垃圾] 服务启动，执行初始文件清理... (保留 ${days} 天)`);
    cleanOldFiles(baseDataDir, days);

    // 计算到下一个凌晨 2 点的毫秒数
    function scheduleNext() {
        const now = new Date();
        const next = new Date(now);
        next.setHours(2, 0, 0, 0); // 设置为凌晨 2 点
        if (now > next) {
            next.setDate(next.getDate() + 1); // 如果今天 2 点已过，设置为明天 2 点
        }
        const delay = next.getTime() - now.getTime();

        setTimeout(() => {
            console.log(`[清理垃圾] 运行每日定期清理任务...`);
            cleanOldFiles(baseDataDir, days);
            scheduleNext(); // 安排下一次
        }, delay);
    }

    scheduleNext();
}

module.exports = {
    cleanOldFiles,
    startDailyCleanup
};
