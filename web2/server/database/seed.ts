import { Database } from 'better-sqlite3';
import { db } from './index';

// 模拟数据初始化函数
export async function seedData() {
  console.log('开始初始化模拟数据...');
  
  try {
    // 检查是否已有数据
    const novelCount = db.prepare('SELECT COUNT(*) as count FROM novels').get() as { count: number };
    if (novelCount.count > 0) {
      console.log('数据库中已有数据，跳过初始化');
      return;
    }

    // 插入模拟小说数据
    const novelsData = [
      {
        title: '星际探险记',
        author: 'AI作家',
        description: '这是一部关于星际探险的科幻小说，讲述了人类探索未知宇宙的冒险故事。',
        category: '科幻',
        cover_url: '/covers/sci-fi-1.jpg',
        is_paid: false,
        created_at: new Date(),
        updated_at: new Date()
      },
      {
        title: '魔法世界之旅',
        author: 'AI作家',
        description: '一个普通少年意外进入魔法世界，开启了一段奇幻冒险。',
        category: '奇幻',
        cover_url: '/covers/fantasy-1.jpg',
        is_paid: true,
        created_at: new Date(),
        updated_at: new Date()
      },
      {
        title: '都市传说',
        author: 'AI作家',
        description: '发生在现代都市中的一系列神秘事件，探索人性的黑暗面。',
        category: '悬疑',
        cover_url: '/covers/mystery-1.jpg',
        is_paid: false,
        created_at: new Date(),
        updated_at: new Date()
      },
      {
        title: '历史长河',
        author: 'AI作家',
        description: '穿越时空，见证历史上的重大事件，体验不同时代的生活。',
        category: '历史',
        cover_url: '/covers/history-1.jpg',
        is_paid: true,
        created_at: new Date(),
        updated_at: new Date()
      },
      {
        title: '未来战士',
        author: 'AI作家',
        description: '在未来世界，人工智能与人类的战争一触即发，一个普通士兵的成长历程。',
        category: '科幻',
        cover_url: '/covers/sci-fi-2.jpg',
        is_paid: false,
        created_at: new Date(),
        updated_at: new Date()
      }
    ];

    // 插入小说
    const insertNovel = db.prepare(
      `INSERT INTO novels (title, author, description, category, cover_image, is_paid, created_at, updated_at)
       VALUES (@title, @author, @description, @category, @cover_url, @is_paid, @created_at, @updated_at)
       RETURNING id`
    );

    for (const novel of novelsData) {
      const novelId = insertNovel.get(novel) as { id: number };
      console.log(`插入小说: ${novel.title}, ID: ${novelId.id}`);
      
      // 为每部小说插入3个章节
      const insertChapter = db.prepare(
        `INSERT INTO chapters (novel_id, title, content, chapter_number, is_paid, created_at, updated_at)
         VALUES (@novel_id, @title, @content, @chapter_number, @is_paid, @created_at, @created_at)`
      );
      
      for (let i = 1; i <= 3; i++) {
        const isPaid = novel.is_paid && i > 1; // 付费小说从第二章开始收费
        const chapter = {
          novel_id: novelId.id,
          title: `第${i}章 ${getChapterTitle(novel.category, i)}`,
          content: generateChapterContent(novel.title, i),
          chapter_number: i,
          word_count: Math.floor(Math.random() * 2000) + 1000,
          is_paid,
          created_at: new Date()
        };
        
        insertChapter.run(chapter);
        console.log(`  插入章节: ${chapter.title} (${isPaid ? '付费' : '免费'})`);
      }
    }

    // 插入模拟用户数据（测试账号）
    const insertUser = db.prepare(
      `INSERT INTO users (username, email, password, role, created_at)
       VALUES (@username, @email, @password, @role, @created_at)`
    );
    
    // 密码：123456（已加密）
    const testUser = {
      username: 'testuser',
      email: 'test@example.com',
      password: '$2a$10$w7Vj2e5xG8eX5v6L7z8y9OeZ9d8c7v6b5n4m3l2k1j0h9g8f7e6d5c4b3a',
      role: 'user',
      created_at: new Date()
    };
    
    const adminUser = {
      username: 'admin',
      email: 'admin@example.com',
      password: '$2a$10$w7Vj2e5xG8eX5v6L7z8y9OeZ9d8c7v6b5n4m3l2k1j0h9g8f7e6d5c4b3a',
      role: 'admin',
      created_at: new Date()
    };
    
    insertUser.run(testUser);
    insertUser.run(adminUser);
    console.log('插入测试用户: testuser / admin (密码: 123456)');
    
    console.log('模拟数据初始化完成！');
  } catch (error) {
    console.error('初始化模拟数据失败:', error);
  }
}

// 根据分类生成章节标题
function getChapterTitle(category: string, chapterNumber: number): string {
  const titles: Record<string, string[]> = {
    '科幻': ['启程', '未知星球', '外星文明'],
    '奇幻': ['魔法学院', '咒语学习', '第一次冒险'],
    '悬疑': ['神秘事件', '线索追踪', '真相大白'],
    '历史': ['初入宫廷', '权力斗争', '历史转折']
  };
  
  const categoryTitles = titles[category] || ['开始', '发展', '高潮'];
  return categoryTitles[chapterNumber - 1] || `章节${chapterNumber}`;
}

// 生成章节内容
function generateChapterContent(novelTitle: string, chapterNumber: number): string {
  return `这是《${novelTitle}》第${chapterNumber}章的内容。\n\n` +
    `在这一章中，主角经历了许多冒险和挑战。这是一段由AI生成的示例文本，\n` +
    `用于展示小说阅读界面。实际内容将更加丰富和精彩。\n\n` +
    `小说是文学的重要形式，通过文字构建丰富多彩的世界，\n` +
    `带给读者无限的想象空间。\n\n` +
    `这里是更多的示例文本，用来填充章节内容...\n` +
    `这里是更多的示例文本，用来填充章节内容...\n` +
    `这里是更多的示例文本，用来填充章节内容...`;
}