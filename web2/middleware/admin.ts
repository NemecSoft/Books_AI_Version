export default defineNuxtRouteMiddleware(async (to) => {
  try {
    const response = await fetch('/api/auth/me', {
      credentials: 'include'
    });
    
    if (!response.ok) {
      // 未登录，重定向到登录页
      return navigateTo(`/login?redirect=${encodeURIComponent(to.path)}`);
    }
    
    const result = await response.json();
    const user = result.data.user;
    
    // 检查用户是否为管理员
    if (user.role !== 'admin') {
      // 不是管理员，重定向到首页
      return navigateTo('/');
    }
  } catch (error) {
    // 发生错误，重定向到登录页
    return navigateTo(`/login?redirect=${encodeURIComponent(to.path)}`);
  }
});