import { defineEventHandler, deleteCookie } from 'h3';

export default defineEventHandler((event) => {
  // 删除token cookie
  deleteCookie(event, 'token', {
    path: '/'
  });

  return {
    success: true,
    message: '登出成功'
  };
});