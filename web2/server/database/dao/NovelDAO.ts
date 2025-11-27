// 小说数据访问对象
import { BaseDAO } from './BaseDAO'

export interface Novel {
  id: number
  title: string
  author: string
  description: string
  cover_image: string
  category: string
  is_paid: boolean
  created_at: string
  updated_at: string
}

export interface Chapter {
  id: number
  novel_id: number
  title: string
  content: string
  chapter_number: number
  is_paid: boolean
  created_at: string
  updated_at: string
}

export class NovelDAO extends BaseDAO<Novel> {
  constructor() {
    super('novels')
  }

  // 根据分类查找小说
  findByCategory(category: string, page: number = 1, pageSize: number = 10): { novels: Novel[], total: number } {
    const offset = (page - 1) * pageSize
    
    const novels = this.query(
      `SELECT * FROM ${this.tableName} WHERE category = ? ORDER BY created_at DESC LIMIT ? OFFSET ?`,
      [category, pageSize, offset]
    )

    const total = this.query(
      `SELECT COUNT(*) as count FROM ${this.tableName} WHERE category = ?`,
      [category]
    )[0].count

    return { novels, total: Number(total) }
  }

  // 搜索小说
  search(keyword: string, page: number = 1, pageSize: number = 10): { novels: Novel[], total: number } {
    const offset = (page - 1) * pageSize
    const searchPattern = `%${keyword}%`
    
    const novels = this.query(
      `SELECT * FROM ${this.tableName} WHERE title LIKE ? OR author LIKE ? OR description LIKE ? ORDER BY created_at DESC LIMIT ? OFFSET ?`,
      [searchPattern, searchPattern, searchPattern, pageSize, offset]
    )

    const total = this.query(
      `SELECT COUNT(*) as count FROM ${this.tableName} WHERE title LIKE ? OR author LIKE ? OR description LIKE ?`,
      [searchPattern, searchPattern, searchPattern]
    )[0].count

    return { novels, total: Number(total) }
  }

  // 获取最新小说
  getLatestNovels(limit: number = 10): Novel[] {
    return this.query(
      `SELECT * FROM ${this.tableName} ORDER BY created_at DESC LIMIT ?`,
      [limit]
    )
  }

  // 获取热门小说（按章节数）
  getPopularNovels(limit: number = 10): Novel[] {
    return this.query(
      `SELECT n.*, COUNT(c.id) as chapter_count FROM ${this.tableName} n
       LEFT JOIN chapters c ON n.id = c.novel_id
       GROUP BY n.id
       ORDER BY chapter_count DESC, n.created_at DESC
       LIMIT ?`,
      [limit]
    )
  }
}

// 章节DAO
export class ChapterDAO extends BaseDAO<Chapter> {
  constructor() {
    super('chapters')
  }

  // 获取小说的所有章节
  getChaptersByNovelId(novelId: number): Chapter[] {
    return this.query(
      `SELECT * FROM ${this.tableName} WHERE novel_id = ? ORDER BY chapter_number ASC`,
      [novelId]
    )
  }

  // 获取小说的付费章节
  getPaidChaptersByNovelId(novelId: number): Chapter[] {
    return this.query(
      `SELECT * FROM ${this.tableName} WHERE novel_id = ? AND is_paid = 1 ORDER BY chapter_number ASC`,
      [novelId]
    )
  }

  // 获取小说的免费章节
  getFreeChaptersByNovelId(novelId: number): Chapter[] {
    return this.query(
      `SELECT * FROM ${this.tableName} WHERE novel_id = ? AND is_paid = 0 ORDER BY chapter_number ASC`,
      [novelId]
    )
  }

  // 获取小说的最新章节
  getLatestChapterByNovelId(novelId: number): Chapter | undefined {
    const chapters = this.query(
      `SELECT * FROM ${this.tableName} WHERE novel_id = ? ORDER BY chapter_number DESC LIMIT 1`,
      [novelId]
    )
    return chapters[0]
  }

  // 获取章节总数
  getChapterCountByNovelId(novelId: number): number {
    const result = this.query(
      `SELECT COUNT(*) as count FROM ${this.tableName} WHERE novel_id = ?`,
      [novelId]
    )
    return Number(result[0].count)
  }
}