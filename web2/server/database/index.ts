// 数据库连接管理
import sqlite3 from 'better-sqlite3'
import path from 'path'
import fs from 'fs'

// 确保数据目录存在 - 使用绝对路径
const dataDir = path.resolve('d:\\AI\\books\\web2\\data')
console.log('正在使用数据目录:', dataDir)
if (!fs.existsSync(dataDir)) {
  try {
    fs.mkdirSync(dataDir, { recursive: true })
    console.log('数据目录创建成功:', dataDir)
  } catch (err) {
    console.error('创建数据目录失败:', err)
  }
}

// 创建数据库连接
const dbPath = path.join(dataDir, 'books.db')
console.log('数据库路径:', dbPath)

// 创建数据库连接
const db = sqlite3(dbPath)

// 初始化数据库表结构
export async function initDatabase() {
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
    `)

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
    `)

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
    `)

    // 事件列表表
    db.exec(`
      CREATE TABLE IF NOT EXISTS events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        chapter_id INTEGER NOT NULL,
        event_type TEXT NOT NULL,
        event_text TEXT NOT NULL,
        timestamp TEXT,
        FOREIGN KEY (chapter_id) REFERENCES chapters(id)
      );
    `)

    // 简要事件表
    db.exec(`
      CREATE TABLE IF NOT EXISTS brief_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        novel_id INTEGER NOT NULL,
        event_text TEXT NOT NULL,
        order_number INTEGER NOT NULL,
        FOREIGN KEY (novel_id) REFERENCES novels(id)
      );
    `)

    // AI内容表
    db.exec(`
      CREATE TABLE IF NOT EXISTS ai_contents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        chapter_id INTEGER NOT NULL,
        content_type TEXT NOT NULL, -- 'image' 或 'video'
        content_url TEXT NOT NULL,
        description TEXT,
        FOREIGN KEY (chapter_id) REFERENCES chapters(id)
      );
    `)

    // 广告表
    db.exec(`
      CREATE TABLE IF NOT EXISTS ads (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ad_type TEXT NOT NULL,
        content TEXT NOT NULL,
        target_url TEXT,
        is_active BOOLEAN DEFAULT 1,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
    `)

    // 支付记录表
    db.exec(`
      CREATE TABLE IF NOT EXISTS payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        novel_id INTEGER,
        amount REAL NOT NULL,
        payment_method TEXT NOT NULL,
        status TEXT DEFAULT 'pending',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (user_id) REFERENCES users(id),
        FOREIGN KEY (novel_id) REFERENCES novels(id)
      );
    `)

    // 用户阅读进度表
    db.exec(`
      CREATE TABLE IF NOT EXISTS reading_progress (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        novel_id INTEGER NOT NULL,
        chapter_id INTEGER NOT NULL,
        last_read_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (user_id) REFERENCES users(id),
        FOREIGN KEY (novel_id) REFERENCES novels(id),
        FOREIGN KEY (chapter_id) REFERENCES chapters(id)
      );
    `)

    // 用户收藏表
    db.exec(`
      CREATE TABLE IF NOT EXISTS favorites (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        novel_id INTEGER NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (user_id) REFERENCES users(id),
        FOREIGN KEY (novel_id) REFERENCES novels(id),
        UNIQUE(user_id, novel_id)
      );
    `)

    // 用户权限表
    db.exec(`
      CREATE TABLE IF NOT EXISTS user_permissions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        permission_type TEXT NOT NULL,
        target_id INTEGER,
        FOREIGN KEY (user_id) REFERENCES users(id)
      );
    `)

    console.log('数据库初始化完成')
  } catch (error) {
    console.error('数据库初始化失败:', error)
    throw error
  }
}

// 导出数据库实例
export { db }