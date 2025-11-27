// API路由入口文件
export default defineEventHandler(async (event) => {
  return {
    status: 'success',
    message: '小说阅读平台API服务运行中',
    version: '1.0.0'
  }
})