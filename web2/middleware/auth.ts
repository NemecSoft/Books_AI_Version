export default defineNuxtRouteMiddleware(async (to) => {
  console.log(`客户端中间件检查路由: ${to.path}`);
  
  // 不需要登录的页面
  const publicPages = ['/login', '/register', '/', '/novels', '/novels/*'];
  
  // 检查当前路由是否需要登录
  const requiresAuth = !publicPages.some(page => {
    if (page.endsWith('/*')) {
      const basePath = page.slice(0, -2);
      const matches = to.path.startsWith(basePath);
      console.log(`检查通配符路径: ${page}, 匹配结果: ${matches}`);
      return matches;
    }
    const matches = to.path === page;
    console.log(`检查精确路径: ${page}, 匹配结果: ${matches}`);
    return matches;
  });
  
  console.log(`路由 ${to.path} 需要认证: ${requiresAuth}`);
  
  // 如果需要登录，检查用户是否已登录
  if (requiresAuth) {
    try {
      console.log('正在检查用户登录状态...');
      const response = await fetch('/api/auth/me', {
        credentials: 'include'
      });
      
      console.log(`/api/auth/me 响应状态: ${response.status}`);
      
      if (!response.ok) {
        console.log(`用户未登录，重定向到登录页`);
        // 未登录，重定向到登录页
        return navigateTo(`/login?redirect=${encodeURIComponent(to.path)}`);
      }
      console.log('用户已登录');
    } catch (error) {
      console.error('检查登录状态时出错:', error);
      // 网络错误，也重定向到登录页
      return navigateTo(`/login?redirect=${encodeURIComponent(to.path)}`);
    }
  }
});