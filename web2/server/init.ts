// 服务器初始化
import { initDatabase } from './database'
import { seedData } from './database/seed'

export default defineNitroPlugin(async () => {
  try {
    // 初始化数据库
    await initDatabase()
    console.log('服务器初始化完成')
    
    // 初始化模拟数据
    await seedData()
  } catch (error) {
    console.error('服务器初始化失败:', error)
    throw error
  }
})