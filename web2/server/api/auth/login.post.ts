import { defineEventHandler, readBody, createError, setCookie } from 'h3';
import jwt from 'jsonwebtoken';
import { UserDAO } from '../../database/dao/UserDAO';

export default defineEventHandler(async (event) => {
  try {
    const body = await readBody(event);
    const { username, password } = body;

    // 验证输入
    if (!username || !password) {
      throw createError({
        statusCode: 400,
        message: '用户名和密码不能为空'
      });
    }

    // 查找用户并验证密码
    const userDAO = new UserDAO();
    const user = await userDAO.findByUsername(username);
    
    if (!user || !(await userDAO.verifyPassword(user, password))) {
      throw createError({
        statusCode: 401,
        message: '用户名或密码错误'
      });
    }

    // 生成JWT令牌
    const token = jwt.sign(
      {
        id: user.id,
        username: user.username,
        role: user.role
      },
      process.env.JWT_SECRET || 'your-secret-key',
      {
        expiresIn: '24h'
      }
    );

    // 设置cookie
    setCookie(event, 'token', token, {
      httpOnly: true,
      secure: process.env.NODE_ENV === 'production',
      maxAge: 24 * 60 * 60,
      path: '/'
    });

    // 返回用户信息（不包含密码）
    const { password: _, ...userInfo } = user;
    
    return {
      success: true,
      message: '登录成功',
      data: {
        user: userInfo,
        token
      }
    };
  } catch (error: any) {
    throw createError({
      statusCode: error.statusCode || 500,
      statusMessage: error.message || '登录失败'
    });
  }
});