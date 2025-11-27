import { defineEventHandler } from 'h3';

export default defineEventHandler((event) => {
  console.log('测试API被调用');
  return {
    success: true,
    message: '测试API工作正常',
    timestamp: new Date().toISOString(),
    data: {
      books: [
        { id: 1, title: '测试小说1', author: '测试作者1' },
        { id: 2, title: '测试小说2', author: '测试作者2' }
      ]
    }
  };
});