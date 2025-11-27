import { defineEventHandler, createError, getCookie } from 'h3';
import jwt from 'jsonwebtoken';

export default defineEventHandler((event) => {
  // 获取当前请求路径
  let path = event.path;
  console.log(`服务器中间件检查路径: ${path}`);
  
  // 去掉查询参数，只保留路径部分
  if (path.includes('?')) {
    path = path.split('?')[0];
    console.log(`去掉查询参数后的路径: ${path}`);
  }
  
  // 不需要认证的路径列表
  const publicPaths = [
    // 页面路由
    '/', 
    '/login', 
    '/register', 
    '/novels', 
    '/novels/*',  // 添加小说详情页面的通配符
    
    // API路由
    '/api/novels',
    '/api/novels/*',
    '/api/auth/login',
    '/api/auth/register',
    '/api/test',
    '/api/hello',
    
    // 静态资源路径
    '/_nuxt/*',
    '/@vite/*',
    '/@nuxt/*',
    '/assets/*',
    '/public/*',
    '/covers/*'  // 小说封面图片
  ];
  
  // 检查当前路径是否需要认证
  const requiresAuth = !publicPaths.some(publicPath => {
    if (publicPath.endsWith('/*')) {
      const basePath = publicPath.slice(0, -2);
      // 改进的通配符匹配：确保路径要么是basePath本身，要么是basePath后跟斜杠和其他内容
      const matches = path === basePath || 
                    (path.startsWith(basePath + '/'));
      console.log(`检查通配符路径: ${publicPath}, 基础路径: ${basePath}, 匹配结果: ${matches}`);
      return matches;
    }
    const matches = path === publicPath;
    console.log(`检查精确路径: ${publicPath}, 匹配结果: ${matches}`);
    return matches;
  });
  
  // 额外检查：静态资源文件（如.js, .css, .png等）不需要认证
  const isStaticAsset = /\.(js|css|png|jpg|jpeg|gif|svg|ico|woff|woff2|ttf|eot)$/i.test(path);
  if (isStaticAsset) {
    console.log(`路径 ${path} 是静态资源，无需认证`);
    return;
  }
  
  console.log(`路径 ${path} 需要认证: ${requiresAuth}`);
  
  // 如果不需要认证，直接通过
  if (!requiresAuth) {
    return;
  }
  
  // 对于需要认证的路径，检查token
  const token = getCookie(event, 'token');
  
  if (!token) {
    throw createError({
      statusCode: 401,
      message: '未登录，请先登录'
    });
  }

  try {
    // 验证token
    const decoded = jwt.verify(token, process.env.JWT_SECRET || 'your-secret-key');
    
    // 将用户信息存储到event上下文中
    (event.context as any).user = decoded;
    
    return;
  } catch (error) {
    throw createError({
      statusCode: 401,
      statusMessage: '登录已过期，请重新登录'
    });
  }
});