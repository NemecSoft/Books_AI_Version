// 基础数据访问对象类
import { db } from '../index'

export class BaseDAO<T> {
  protected tableName: string

  constructor(tableName: string) {
    this.tableName = tableName
  }

  // 查询所有记录
  findAll(): T[] {
    const sql = `SELECT * FROM ${this.tableName}`
    return db.prepare(sql).all() as T[]
  }

  // 根据ID查询记录
  findById(id: number): T | undefined {
    const sql = `SELECT * FROM ${this.tableName} WHERE id = ?`
    return db.prepare(sql).get(id) as T | undefined
  }

  // 插入记录
  insert(data: Partial<T>): number {
    const fields = Object.keys(data)
    const placeholders = fields.map(() => '?').join(', ')
    const sql = `INSERT INTO ${this.tableName} (${fields.join(', ')}) VALUES (${placeholders})`
    const info = db.prepare(sql).run(...Object.values(data))
    return info.lastInsertRowid as number
  }

  // 更新记录
  update(id: number, data: Partial<T>): boolean {
    const fields = Object.keys(data)
    const updates = fields.map(field => `${field} = ?`).join(', ')
    const sql = `UPDATE ${this.tableName} SET ${updates} WHERE id = ?`
    const info = db.prepare(sql).run(...Object.values(data), id)
    return info.changes > 0
  }

  // 删除记录
  delete(id: number): boolean {
    const sql = `DELETE FROM ${this.tableName} WHERE id = ?`
    const info = db.prepare(sql).run(id)
    return info.changes > 0
  }

  // 根据条件查询
  findByCondition(condition: Record<string, any>): T[] {
    const whereClauses = Object.keys(condition).map(key => `${key} = ?`)
    const sql = `SELECT * FROM ${this.tableName} WHERE ${whereClauses.join(' AND ')}`
    return db.prepare(sql).all(...Object.values(condition)) as T[]
  }

  // 执行原始SQL查询
  query(sql: string, params: any[] = []): any[] {
    return db.prepare(sql).all(...params)
  }

  // 执行原始SQL语句（更新、删除等）
  execute(sql: string, params: any[] = []): boolean {
    const info = db.prepare(sql).run(...params)
    return info.changes >= 0
  }
}