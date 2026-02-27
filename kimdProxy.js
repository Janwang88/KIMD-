const fetch = require('node-fetch');
const https = require('https');

// 创建一个忽略SSL证书验证的agent（仅用于开发环境）
const httpsAgent = new https.Agent({
    rejectUnauthorized: false
});

const BASE_URL = 'https://chajian.kimd.cn:9999';
const BASE_IP_URL = 'https://153.35.130.19:9999';

/**
 * 带有超时的 Fetch 请求
 */
async function fetchWithTimeout(url, options, timeoutMs = 20000) {
    return await Promise.race([
        fetch(url, options),
        new Promise((_, reject) => {
            setTimeout(() => reject(new Error('TIMEOUT')), timeoutMs);
        })
    ]);
}

/**
 * 带有重试机制的 Fetch 请求
 */
async function fetchWithRetry(url, options, retries = 2) {
    let lastErr = null;
    for (let i = 0; i <= retries; i++) {
        try {
            return await fetchWithTimeout(url, options);
        } catch (err) {
            lastErr = err;
            // 简单退避
            await new Promise(r => setTimeout(r, 500 * (i + 1)));
        }
    }
    throw lastErr;
}

/**
 * 先尝试域名请求，失败则重试 IP 直连
 */
async function fetchWithFallback(path, options) {
    const url = `${BASE_URL}${path}`;
    try {
        return await fetchWithRetry(url, options);
    } catch (err) {
        // DNS 解析失败或超时，尝试 IP 直连
        if (err && (err.code === 'ENOTFOUND' || err.message === 'TIMEOUT' || err.code === 'ECONNRESET')) {
            const headers = Object.assign({}, options?.headers || {});
            // 使用 IP 直连时，补充 Host 头保持与域名一致
            if (!headers['Host']) {
                headers['Host'] = 'chajian.kimd.cn:9999';
            }
            const ipUrl = `${BASE_IP_URL}${path}`;
            return await fetchWithRetry(ipUrl, { ...options, headers });
        }
        throw err;
    }
}

module.exports = {
    fetchWithFallback,
    httpsAgent
};
