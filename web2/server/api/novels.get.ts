import { defineEventHandler } from 'h3';

export default defineEventHandler((event) => {
  console.log('处理小说列表请求 - 硬编码数据');
  
  // 直接返回硬编码的小说数据
  const mockNovels = [
    {
      id: 1,
      title: '星际漫游',
      author: '刘慈欣',
      category: '科幻',
      description: '一部关于星际旅行的科幻小说',
      cover: 'https://example.com/cover1.jpg',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    },
    {
      id: 2,
      title: '魔法世界',
      author: 'JK罗琳',
      category: '奇幻',
      description: '一个充满魔法的奇幻世界',
      cover: 'https://example.com/cover2.jpg',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    },
    {
      id: 3,
      title: '战争与和平',
      author: '托尔斯泰',
      category: '文学',
      description: '一部经典的文学作品',
      cover: 'https://example.com/cover3.jpg',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    }
  ];
  
  // 确保返回正确格式的对象
  return {
    success: true,
    data: {
      novels: mockNovels,
      pagination: {
        page: 1,
        pageSize: 8,
        total: 3,
        totalPages: 1
      }
    }
  };
});