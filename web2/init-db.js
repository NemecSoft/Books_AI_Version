// 手动初始化数据库脚本
const sqlite3 = require('better-sqlite3');
const path = require('path');
const fs = require('fs');
const bcrypt = require('bcryptjs');

// 确保数据目录存在
const dataDir = path.resolve('./data');
console.log('正在使用数据目录:', dataDir);
if (!fs.existsSync(dataDir)) {
  try {
    fs.mkdirSync(dataDir, { recursive: true });
    console.log('数据目录创建成功:', dataDir);
  } catch (err) {
    console.error('创建数据目录失败:', err);
    process.exit(1);
  }
}

// 创建数据库连接
const dbPath = path.join(dataDir, 'books.db');
console.log('数据库路径:', dbPath);

// 删除现有数据库文件
if (fs.existsSync(dbPath)) {
  try {
    fs.unlinkSync(dbPath);
    console.log('已删除现有数据库文件');
  } catch (err) {
    console.error('删除现有数据库文件失败:', err);
    process.exit(1);
  }
}

// 创建新的数据库连接
const db = sqlite3(dbPath);

// 初始化数据库表结构
console.log('开始初始化数据库表结构...');

try {
  // 用户表
  db.exec(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT NOT NULL UNIQUE,
      password TEXT NOT NULL,
      email TEXT NOT NULL UNIQUE,
      role TEXT DEFAULT 'user',
      created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
  `);
  console.log('✓ 用户表创建成功');

  // 小说表
  db.exec(`
    CREATE TABLE IF NOT EXISTS novels (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      title TEXT NOT NULL,
      author TEXT NOT NULL,
      description TEXT,
      cover_image TEXT,
      category TEXT,
      is_paid BOOLEAN DEFAULT 0,
      created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
      updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
  `);
  console.log('✓ 小说表创建成功');

  // 章节表
  db.exec(`
    CREATE TABLE IF NOT EXISTS chapters (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      novel_id INTEGER NOT NULL,
      title TEXT NOT NULL,
      content TEXT NOT NULL,
      chapter_number INTEGER NOT NULL,
      is_paid BOOLEAN DEFAULT 0,
      created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
      updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY (novel_id) REFERENCES novels(id)
    );
  `);
  console.log('✓ 章节表创建成功');

  // 插入测试用户数据
  const insertUser = db.prepare(
    `INSERT INTO users (username, email, password, role, created_at)
     VALUES (@username, @email, @password, @role, @created_at)`
  );
  
  // 密码：123456（使用bcrypt加密）
  const now = new Date().toISOString();
  const password = '123456';
  const hashedPassword = bcrypt.hashSync(password, 10);
  
  const testUser = {
    username: 'testuser',
    email: 'test@example.com',
    password: hashedPassword,
    role: 'user',
    created_at: now
  };
  
  const adminUser = {
    username: 'admin',
    email: 'admin@example.com',
    password: hashedPassword,
    role: 'admin',
    created_at: now
  };
  
  insertUser.run(testUser);
  insertUser.run(adminUser);
  console.log('✓ 测试用户数据插入成功');

  console.log('\n数据库初始化完成！');
  console.log('\n测试账号：');
  console.log('  用户名：testuser，密码：123456（普通用户）');
  console.log('  用户名：admin，密码：123456（管理员）');
  
} catch (error) {
  console.error('数据库初始化失败:', error);
  process.exit(1);
} finally {
  // 关闭数据库连接
  db.close();
}