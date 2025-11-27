// 启动服务器并自动打开浏览器脚本
// 运行方法: node start-server.js

import { execSync, exec } from 'child_process';
import { platform } from 'os';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { dirname, resolve } from 'path';

// 在ES模块中模拟__dirname
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

console.log('正在启动 AI 小说平台...');
console.log('=================================');

// 检查是否安装了依赖
console.log('\n检查依赖...');
try {
  require('nuxt');
  console.log('✓ Nuxt 已安装');
} catch (e) {
  console.log('✗ Nuxt 未安装，正在安装依赖...');
  try {
    execSync('npm install', { stdio: 'inherit' });
    console.log('✓ 依赖安装成功');
  } catch (error) {
    console.error('✗ 依赖安装失败，请手动运行 npm install');
    process.exit(1);
  }
}

// 确保 covers 目录存在
console.log('\n检查必要目录...');
try {
  const coversDir = resolve(__dirname, 'public', 'covers');
  if (!fs.existsSync(coversDir)) {
    fs.mkdirSync(coversDir, { recursive: true });
    console.log('✓ 创建 covers 目录成功');
  } else {
    console.log('✓ covers 目录已存在');
  }
}
catch (error) {
  console.error('✗ 创建目录失败:', error.message);
}

// 启动开发服务器
console.log('\n启动开发服务器...');
console.log('访问地址: http://localhost:3000');
console.log('按 Ctrl+C 停止服务器');
console.log('=================================');

// 启动服务器
const serverProcess = exec('npm run dev');

serverProcess.stdout.on('data', (data) => {
  console.log(data.toString());
});

serverProcess.stderr.on('data', (data) => {
  console.error(data.toString());
});

// 延迟打开浏览器
setTimeout(() => {
  try {
    let command;
    switch (platform()) {
      case 'win32':
        command = 'start http://localhost:3000';
        break;
      case 'darwin':
        command = 'open http://localhost:3000';
        break;
      case 'linux':
        command = 'xdg-open http://localhost:3000';
        break;
      default:
        return;
    }
    exec(command);
    console.log('✓ 已自动打开浏览器');
  } catch (error) {
    console.log('提示: 请手动打开浏览器访问 http://localhost:3000');
  }
}, 3000);
