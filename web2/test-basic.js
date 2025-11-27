// 基本功能测试脚本
// 运行方法: node test-basic.js

const https = require('https');
const http = require('http');
const { execSync } = require('child_process');

console.log('开始测试 AI 小说平台基本功能...');
console.log('=================================');

// 测试数据库初始化
try {
  console.log('\n1. 测试数据库初始化...');
  const dbInitResult = execSync('node -e "require(\'./server/database\').initDatabase()"', { encoding: 'utf8' });
  console.log('✓ 数据库初始化成功');
} catch (error) {
  console.error('✗ 数据库初始化失败:', error.message);
}

// 测试API接口
testApi('/api/novels', '获取小说列表');

testApi('/api/auth/login', '登录接口', 'POST', JSON.stringify({
  username: 'testuser',
  password: '123456'
}));

// 测试页面路由
console.log('\n3. 测试页面路由 (需要启动开发服务器)...');
console.log('- 首页: http://localhost:3000/');
console.log('- 小说列表: http://localhost:3000/novels');
console.log('- 登录页面: http://localhost:3000/login');
console.log('- 注册页面: http://localhost:3000/register');

console.log('\n测试完成! 请启动开发服务器并在浏览器中验证页面功能。');
console.log('启动命令: npm run dev');

// 辅助函数: 测试API接口
function testApi(path, description, method = 'GET', body = null) {
  return new Promise((resolve) => {
    console.log(`\n2. 测试${description}...`);
    
    const options = {
      hostname: 'localhost',
      port: 3000,
      path: path,
      method: method,
      headers: {
        'Content-Type': 'application/json'
      }
    };

    const req = http.request(options, (res) => {
      let data = '';
      
      res.on('data', (chunk) => {
        data += chunk;
      });
      
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          console.log(`✓ ${description} 响应状态码: ${res.statusCode}`);
          try {
            const jsonData = JSON.parse(data);
            console.log(`  响应数据格式: ${typeof jsonData}`);
            if (Array.isArray(jsonData)) {
              console.log(`  数据长度: ${jsonData.length}`);
            } else if (typeof jsonData === 'object') {
              console.log(`  数据键: ${Object.keys(jsonData).join(', ')}`);
            }
          } catch (e) {
            console.log('  响应不是有效的JSON');
          }
        } else {
          console.log(`✗ ${description} 失败，状态码: ${res.statusCode}`);
        }
        resolve();
      });
    });
    
    req.on('error', (error) => {
      console.log(`✗ ${description} 请求失败: ${error.message}`);
      console.log('  请确保开发服务器已启动');
      resolve();
    });
    
    if (body) {
      req.write(body);
    }
    
    req.end();
  });
}
