import { defineEventHandler, getRouterParams } from 'h3';
import { NovelDAO } from '../../../database/dao/NovelDAO';
import { ChapterDAO } from '../../../database/dao/NovelDAO';

export default defineEventHandler(async (event) => {
  const params = getRouterParams(event);
  const novelId = params.id as string;
  
  if (!novelId) {
    return {
      success: false,
      message: '小说ID不能为空'
    };
  }
  
  const novelDAO = new NovelDAO();
  const chapterDAO = new ChapterDAO();
  
  try {
    // 获取小说详情
    const novel = await novelDAO.findById(parseInt(novelId));
    
    if (!novel) {
      return {
        success: false,
        message: '小说不存在'
      };
    }
    
    // 获取小说章节列表
    const chapters = await chapterDAO.findByNovelId(parseInt(novelId), 0, 50);
    
    return {
      success: true,
      data: {
        novel,
        chapters
      }
    };
  } catch (error) {
    return {
      success: false,
      message: '获取小说详情失败',
      error: error instanceof Error ? error.message : String(error)
    };
  }
});