// 用户数据访问对象
import { BaseDAO } from './BaseDAO'
import bcrypt from 'bcryptjs'

export interface User {
  id: number
  username: string
  password: string
  email: string
  role: string
  created_at: string
}

export class UserDAO extends BaseDAO<User> {
  constructor() {
    super('users')
  }

  // 根据用户名查找用户
  findByUsername(username: string): User | undefined {
    return this.findByCondition({ username })[0]
  }

  // 根据邮箱查找用户
  findByEmail(email: string): User | undefined {
    return this.findByCondition({ email })[0]
  }

  // 创建新用户（密码加密）
  createUser(userData: Omit<User, 'id' | 'created_at'>): number {
    const hashedPassword = bcrypt.hashSync(userData.password, 10)
    return this.insert({
      ...userData,
      password: hashedPassword
    })
  }

  // 验证用户密码
  verifyPassword(user: User, password: string): boolean {
    return bcrypt.compareSync(password, user.password)
  }

  // 更新用户密码
  updatePassword(userId: number, newPassword: string): boolean {
    const hashedPassword = bcrypt.hashSync(newPassword, 10)
    return this.update(userId, { password: hashedPassword })
  }

  // 获取用户列表（分页）
  getUserList(page: number = 1, pageSize: number = 10): { users: User[], total: number } {
    const offset = (page - 1) * pageSize
    
    const users = this.query(
      `SELECT * FROM ${this.tableName} ORDER BY created_at DESC LIMIT ? OFFSET ?`,
      [pageSize, offset]
    )

    const total = this.query(`SELECT COUNT(*) as count FROM ${this.tableName}`)[0].count

    return { users, total: Number(total) }
  }
}