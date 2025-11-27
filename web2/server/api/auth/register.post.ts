import { defineEventHandler, readBody, createError } from 'h3';
import { UserDAO } from '../../database/dao/UserDAO';

export default defineEventHandler(async (event) => {
  try {
    const body = await readBody(event);
    const { username, email, password } = body;

    // 验证输入
    if (!username || !email || !password) {
      throw createError({
        statusCode: 400,
        message: '用户名、邮箱和密码不能为空'
      });
    }

    // 检查用户名和邮箱是否已存在
    const userDAO = new UserDAO();
    
    if (await userDAO.findByUsername(username)) {
      throw createError({
        statusCode: 400,
        statusMessage: '用户名已存在'
      });
    }

    if (await userDAO.findByEmail(email)) {
      throw createError({
        statusCode: 400,
        statusMessage: '邮箱已被注册'
      });
    }

    // 创建新用户
    const newUser = await userDAO.createUser({
      username,
      email,
      password,
      role: 'user', // 默认角色为普通用户
      created_at: new Date(),
      updated_at: new Date()
    });

    // 返回用户信息（不包含密码）
    const { password: _, ...userInfo } = newUser;

    return {
      success: true,
      message: '注册成功',
      data: {
        user: userInfo
      }
    };
  } catch (error: any) {
    throw createError({
      statusCode: error.statusCode || 500,
      statusMessage: error.message || '注册失败'
    });
  }
});