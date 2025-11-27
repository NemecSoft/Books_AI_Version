import { defineEventHandler, createError, getCookie } from 'h3';
import jwt from 'jsonwebtoken';
import { UserDAO } from '../../database/dao/UserDAO';

export default defineEventHandler(async (event) => {
  try {
    // 从cookie获取token
    const token = getCookie(event, 'token');
    
    if (!token) {
      throw createError({
        statusCode: 401,
        message: '未登录'
      });
    }

    // 验证token
    let decoded: any;
    try {
      decoded = jwt.verify(token, process.env.JWT_SECRET || 'your-secret-key');
    } catch (error) {
      throw createError({
        statusCode: 401,
        statusMessage: '无效的令牌'
      });
    }

    // 获取用户信息
    const userDAO = new UserDAO();
    const user = await userDAO.findById(decoded.id);
    
    if (!user) {
      throw createError({
        statusCode: 404,
        statusMessage: '用户不存在'
      });
    }

    // 返回用户信息（不包含密码）
    const { password: _, ...userInfo } = user;
    
    return {
      success: true,
      data: {
        user: userInfo
      }
    };
  } catch (error: any) {
    throw createError({
      statusCode: error.statusCode || 500,
      statusMessage: error.message || '获取用户信息失败'
    });
  }
});