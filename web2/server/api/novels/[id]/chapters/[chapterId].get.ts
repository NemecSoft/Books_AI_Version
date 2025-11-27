import { defineEventHandler, getRouterParams, createError, getCookie } from 'h3';
import { ChapterDAO } from '../../../../database/dao/NovelDAO';
import jwt from 'jsonwebtoken';

export default defineEventHandler(async (event) => {
  const params = getRouterParams(event);
  const novelId = params.id as string;
  const chapterId = params.chapterId as string;
  
  if (!novelId || !chapterId) {
    throw createError({
      statusCode: 400,
      statusMessage: '小说ID和章节ID不能为空'
    });
  }
  
  const chapterDAO = new ChapterDAO();
  
  try {
    // 获取章节信息
    const chapter = await chapterDAO.findById(parseInt(chapterId));
    
    if (!chapter) {
      throw createError({
        statusCode: 404,
        statusMessage: '章节不存在'
      });
    }
    
    // 检查章节是否属于该小说
    if (chapter.novel_id !== parseInt(novelId)) {
      throw createError({
        statusCode: 400,
        statusMessage: '章节不属于该小说'
      });
    }
    
    // 对于付费章节，检查用户是否登录
    if (chapter.is_paid) {
      const token = getCookie(event, 'token');
      
      if (!token) {
        throw createError({
          statusCode: 401,
          statusMessage: '需要登录才能阅读付费章节'
        });
      }
      
      try {
        jwt.verify(token, process.env.JWT_SECRET || 'your-secret-key');
        // 这里可以添加检查用户是否已购买该章节的逻辑
      } catch (error) {
        throw createError({
          statusCode: 401,
          statusMessage: '登录已过期，请重新登录'
        });
      }
    }
    
    return {
      success: true,
      data: {
        chapter
      }
    };
  } catch (error: any) {
    throw createError({
      statusCode: error.statusCode || 500,
      statusMessage: error.statusMessage || '获取章节内容失败'
    });
  }
});